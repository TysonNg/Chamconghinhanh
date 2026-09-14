"""Supplementary photo records; image bytes and capture metadata are immutable."""
import hashlib
import io
import json
import sqlite3
import uuid
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

from flask import Blueprint, jsonify, request, send_file
from PIL import Image, UnidentifiedImageError

LABEL = 'Ảnh bổ sung — không xác nhận thời điểm chụp'


def register_batches(app, data_dir):
    root = Path(data_dir) / 'supplement_output'
    root.mkdir(parents=True, exist_ok=True)
    db = Path(data_dir) / 'supplement_batches.sqlite3'

    def connect():
        c = sqlite3.connect(db, timeout=30)
        c.row_factory = sqlite3.Row
        return c

    with connect() as c:
        c.execute('CREATE TABLE IF NOT EXISTS batches (id TEXT PRIMARY KEY, document TEXT NOT NULL)')

    def load(c, bid):
        row = c.execute('SELECT document FROM batches WHERE id=?', (bid,)).fetchone()
        if not row:
            raise LookupError('Không tìm thấy batch')
        return json.loads(row[0])

    def save(c, b):
        c.execute('INSERT OR REPLACE INTO batches VALUES (?, ?)', (b['id'], json.dumps(b, ensure_ascii=False)))

    def log(b, action, details):
        b['history'].append({'at': datetime.now(timezone.utc).isoformat(), 'action': action, 'details': details})

    def verify(b):
        for item in b['items']:
            path = root / b['id'] / item['archive_name']
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != item['sha256']:
                raise RuntimeError('Ảnh lưu trữ đã thay đổi hoặc bị mất; không thể duyệt/xuất')

    bp = Blueprint('supplement_batches', __name__)

    @app.before_request
    def retire_timestamp_mutations():
        if request.method == 'POST' and request.path in (
                '/api/supplement/upload', '/api/supplement/process'):
            return jsonify(error='Luồng đổi timestamp đã ngừng. Dùng hồ sơ ảnh bổ sung tại /api/supplement/batches; ảnh và EXIF được giữ nguyên.'), 410

    def parse_day(value):
        if not isinstance(value, str):
            raise ValueError('Ngày hồ sơ phải có định dạng YYYY-MM-DD')
        return date.fromisoformat(value).isoformat()

    def payload():
        data = request.get_json()
        if not isinstance(data, dict):
            raise ValueError('Nội dung JSON phải là object')
        return data

    @bp.errorhandler(ValueError)
    def invalid(e):
        return jsonify(error=str(e)), 400

    @bp.errorhandler(LookupError)
    def missing(e):
        return jsonify(error=str(e)), 404

    @bp.errorhandler(RuntimeError)
    def conflict(e):
        return jsonify(error=str(e)), 409

    @bp.route('/api/supplement/batches', methods=['GET', 'POST'])
    def batches():
        if request.method == 'GET':
            with connect() as c:
                return jsonify(batches=[json.loads(r[0]) for r in c.execute('SELECT document FROM batches ORDER BY rowid DESC')])
        files = request.files.getlist('photos')
        if not 1 <= len(files) <= 31:
            raise ValueError('Chọn từ 1 đến 31 ảnh')
        employee = request.form.get('employee', '').strip()
        if not employee or len(employee) > 200:
            raise ValueError('Tên nhân viên không hợp lệ')
        day = parse_day(request.form.get('supplement_date', ''))
        prepared = []
        total = 0
        for f in files:
            raw = f.read(20 * 1024 * 1024 + 1)
            total += len(raw)
            if total > 100 * 1024 * 1024:
                raise ValueError('Tổng ảnh tối đa 100MB')
            if len(raw) > 20 * 1024 * 1024:
                raise ValueError('Mỗi ảnh tối đa 20MB')
            try:
                with Image.open(io.BytesIO(raw)) as im:
                    fmt = im.format
                    if fmt not in ('JPEG', 'PNG', 'WEBP', 'BMP') or im.width * im.height > 40000000:
                        raise ValueError('Định dạng hoặc kích thước ảnh không hỗ trợ')
                    im.verify()
                with Image.open(io.BytesIO(raw)) as decoded:
                    decoded.load()
            except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as e:
                raise ValueError('File ảnh không hợp lệ') from e
            ext = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp', 'BMP': '.bmp'}[fmt]
            iid = uuid.uuid4().hex
            prepared.append((raw, {'id': iid, 'archive_name': iid + ext,
                'source_name': Path(f.filename or 'photo').name, 'supplement_date': day,
                'note': '', 'sha256': hashlib.sha256(raw).hexdigest()}))
        bid = uuid.uuid4().hex
        folder = root / bid
        folder.mkdir()
        b = {'id': bid, 'employee': employee, 'label': LABEL, 'status': 'draft',
             'items': [i for _, i in prepared], 'history': []}
        try:
            for raw, item in prepared:
                (folder / item['archive_name']).write_bytes(raw)
            log(b, 'created', {'count': len(prepared)})
            with connect() as c:
                save(c, b)
        except Exception:
            for _, item in prepared:
                (folder / item['archive_name']).unlink(missing_ok=True)
            folder.rmdir()
            raise
        return jsonify(batch=b), 201

    @bp.get('/api/supplement/batches/<bid>')
    def detail(bid):
        with connect() as c:
            return jsonify(batch=load(c, bid))

    @bp.patch('/api/supplement/batches/<bid>/items/<iid>')
    def update(bid, iid):
        data = payload()
        day = parse_day(data.get('supplement_date', ''))
        note = data.get('note', '')
        if not isinstance(note, str) or len(note) > 2000:
            raise ValueError('Ghi chú tối đa 2000 ký tự')
        with connect() as c:
            c.execute('BEGIN IMMEDIATE')
            b = load(c, bid)
            item = next((i for i in b['items'] if i['id'] == iid), None)
            if item is None:
                raise LookupError('Không tìm thấy ảnh')
            old = dict(item)
            item.update(supplement_date=day, note=note)
            b['status'] = 'draft'
            log(b, 'item_updated', {'item_id': iid, 'before': old, 'after': dict(item)})
            save(c, b)
        return jsonify(batch=b)

    @bp.get('/api/supplement/batches/<bid>/items/<iid>/image')
    def preview(bid, iid):
        with connect() as c:
            b = load(c, bid)
        item = next((i for i in b['items'] if i['id'] == iid), None)
        if item is None:
            raise LookupError('Không tìm thấy ảnh')
        return send_file(root / b['id'] / item['archive_name'])

    @bp.post('/api/supplement/batches/<bid>/approve')
    def approve(bid):
        reviewer = payload().get('reviewer', '')
        if not isinstance(reviewer, str) or not reviewer.strip() or len(reviewer) > 200:
            raise ValueError('Nhập tên người duyệt')
        with connect() as c:
            c.execute('BEGIN IMMEDIATE')
            b = load(c, bid)
            verify(b)
            b['status'] = 'approved'
            log(b, 'approved', {'reviewer': reviewer.strip()})
            save(c, b)
        return jsonify(batch=b)

    @bp.get('/api/supplement/batches/<bid>/download')
    def download(bid):
        with connect() as c:
            b = load(c, bid)
        if b['status'] != 'approved':
            raise RuntimeError('Cần duyệt batch trước khi xuất')
        verify(b)
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w', zipfile.ZIP_STORED) as z:
            z.writestr('manifest.json', json.dumps(b, ensure_ascii=False, indent=2))
            z.writestr('README.txt', LABEL + '\nNgày bổ sung trong manifest là ngày hồ sơ, không phải ngày chụp.\nẢnh và EXIF được giữ nguyên.\n')
            for item in b['items']:
                raw = (root / bid / item['archive_name']).read_bytes()
                if hashlib.sha256(raw).hexdigest() != item['sha256']:
                    raise RuntimeError('Ảnh thay đổi trong khi xuất')
                z.writestr(item['archive_name'], raw)
        stream.seek(0)
        return send_file(stream, mimetype='application/zip', as_attachment=True, download_name=f'anh-bo-sung-{bid}.zip')

    app.register_blueprint(bp)
