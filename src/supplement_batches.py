"""Supplementary photo records; image bytes and capture metadata are immutable."""
import hashlib
import io
import json
import re
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

    # ==================== FAKE PHOTO ENGINE & STAGING ROUTES ====================
    from src.fake_photo_service import FakePhotoService
    fake_service = FakePhotoService(data_dir)

    @bp.post('/api/supplement/records')
    def create_supplement_records():
        files = request.files.getlist('photos')
        project = request.form.get('project', '').strip()
        employee = request.form.get('employee', '').strip()
        if not project or not employee or not 1 <= len(files) <= 31:
            raise ValueError('Chọn dự án, nhân viên và từ 1 đến 31 ảnh')
        try:
            configs = json.loads(request.form.get('configs', '[]'))
        except (TypeError, json.JSONDecodeError):
            raise ValueError('Cấu hình ngày bổ sung không hợp lệ')
        if not isinstance(configs, list) or len(configs) != len(files):
            raise ValueError('Mỗi ảnh cần có ngày đề nghị bổ sung')
        prepared = []
        total = 0
        default_shift = request.form.get('shift')
        for file, config in zip(files, configs):
            if not isinstance(config, dict):
                raise ValueError('Cấu hình ảnh không hợp lệ')
            day = parse_day(config.get('target_date'))
            target_time = config.get('target_time', '').strip()
            if not target_time and default_shift:
                target_time = fake_service.generate_random_time(default_shift)
            elif target_time.count(':') == 1:
                target_time = f"{target_time}:00"

            raw = file.read(20 * 1024 * 1024 + 1)
            total += len(raw)
            if len(raw) > 20 * 1024 * 1024 or total > 100 * 1024 * 1024:
                raise ValueError('Mỗi ảnh tối đa 20MB, tổng tối đa 100MB')
            try:
                with Image.open(io.BytesIO(raw)) as image:
                    if image.format not in ('JPEG', 'PNG', 'WEBP', 'BMP') or image.width * image.height > 40000000:
                        raise ValueError('Định dạng hoặc kích thước ảnh không hỗ trợ')
                    image.verify()
                with Image.open(io.BytesIO(raw)) as image:
                    image.load()
            except (OSError, Image.DecompressionBombError) as exc:
                raise ValueError('File ảnh không hợp lệ') from exc
            prepared.append((raw, day, target_time, Path(file.filename or 'photo').name))

        items = []
        for raw, day, t_time, name in prepared:
            item = fake_service.create_fake_photo(
                raw_bytes=raw,
                project=project,
                employee=employee,
                target_date=day,
                target_time=t_time,
                original_name=name,
                replace_timestamp=True,
                modify_exif=True
            )
            try:
                insp = fake_service.ai_service.inspect_photo(raw)
                if insp and isinstance(insp, dict):
                    insp['original_visible_date'] = insp.get('visible_date')
                    item['inspection'].update(insp)
                    with fake_service._connect() as c:
                        c.execute('UPDATE staging_photos SET inspection_json=? WHERE id=?', (json.dumps(item['inspection'], ensure_ascii=False), item['id']))
            except Exception:
                pass
            items.append(item)
        return jsonify(success=True, count=len(items), items=items), 201

    # ==================== AI CONFIG & TEST ROUTES ====================
    @bp.route('/api/supplement/ai-config', methods=['GET', 'POST'])
    def ai_config():
        if request.method == 'GET':
            return jsonify(fake_service.ai_service.get_public_config())

        data = request.get_json() or {}
        # Nếu người dùng giữ nguyên chuỗi đã che (masked) hoặc để trống, giữ lại key cũ
        existing = fake_service.ai_service.get_full_config()
        for key_field in ['gemini_api_key', 'openai_api_key']:
            val = data.get(key_field, '')
            if not val or '...' in val:
                data[key_field] = existing.get(key_field, '')

        success = fake_service.ai_service.save_config(data)
        if success:
            return jsonify(success=True, config=fake_service.ai_service.get_public_config())
        return jsonify(error='Không thể lưu cấu hình AI'), 500

    @bp.route('/api/supplement/ai-test', methods=['POST'])
    def ai_test():
        data = request.get_json() or {}
        provider = data.get('provider', 'gemini')
        model = data.get('model', '')
        api_key = data.get('api_key', '').strip()

        # Nếu không truyền key mới hoặc truyền key đã che, dùng key đã lưu
        existing = fake_service.ai_service.get_full_config()
        if not model:
            model = existing.get(provider + '_model', '')
        if not api_key or '...' in api_key:
            if provider == 'gemini':
                api_key = existing.get('gemini_api_key', '')
            elif provider == 'openai':
                api_key = existing.get('openai_api_key', '')

        ok, msg = fake_service.ai_service.test_connection(provider, api_key, model)
        return jsonify(success=ok, message=msg)

    @bp.route('/api/supplement/fake-photos', methods=['POST'])
    def create_fake_photos():
        files = request.files.getlist('photos')
        if not files:
            return jsonify(error='Vui lòng chọn ít nhất 1 ảnh'), 400
        if len(files) > 31:
            return jsonify(error='Tối đa 31 ảnh mỗi lần'), 400

        project = request.form.get('project', '').strip()
        employee = request.form.get('employee', '').strip()
        if not project or not employee:
            return jsonify(error='Vui lòng chọn dự án và nhân viên'), 400

        replace_ts = request.form.get('replace_timestamp', 'true').lower() in ('true', '1')
        mod_exif = request.form.get('modify_exif', 'true').lower() in ('true', '1')
        location = request.form.get('location_name', project).strip()

        raw_configs = request.form.get('configs', '')
        configs_map = {}
        if raw_configs:
            try:
                parsed = json.loads(raw_configs)
                if isinstance(parsed, list):
                    for idx, item in enumerate(parsed):
                        configs_map[idx] = item
            except Exception:
                pass

        created_items = []
        default_date = request.form.get('default_date', date.today().isoformat())
        default_time = request.form.get('default_time', '08:00:00')

        for idx, f in enumerate(files):
            raw = f.read()
            cfg = configs_map.get(idx, {})
            t_date = cfg.get('target_date') or request.form.get(f'date_{idx}') or default_date
            t_time = cfg.get('target_time') or request.form.get(f'time_{idx}') or default_time

            if len(t_time.split(':')) == 2:
                t_time = f"{t_time}:00"

            item = fake_service.create_fake_photo(
                raw_bytes=raw,
                project=project,
                employee=employee,
                target_date=t_date,
                target_time=t_time,
                original_name=Path(f.filename or f"photo_{idx}.jpg").name,
                replace_timestamp=replace_ts,
                modify_exif=mod_exif,
                location_name=location
            )
            created_items.append(item)

        return jsonify(success=True, count=len(created_items), items=created_items), 201

    @bp.route('/api/supplement/staging', methods=['GET'])
    def get_staging():
        proj = request.args.get('project')
        emp = request.args.get('employee')
        items = fake_service.get_staging_photos(project=proj, employee=emp)
        return jsonify(success=True, count=len(items), items=items)

    @bp.route('/api/supplement/staging/<photo_id>/image', methods=['GET'])
    def get_staging_image(photo_id):
        img_path = fake_service.get_staging_image_path(photo_id)
        if not img_path or not img_path.exists():
            return jsonify(error='Không tìm thấy ảnh'), 404
        return send_file(img_path)

    @bp.route('/api/supplement/staging/<photo_id>/watermark-crop', methods=['GET'])
    def get_staging_watermark_crop(photo_id):
        crop = fake_service.get_watermark_crop_bytes(photo_id)
        if not crop:
            return jsonify(error='Không tìm thấy crop watermark từ ảnh gốc'), 404
        return send_file(io.BytesIO(crop), mimetype='image/jpeg')

    @bp.route('/api/supplement/apply-to-attendance', methods=['POST'])
    def apply_to_attendance():
        data = request.get_json() or {}
        ids = data.get('ids', [])
        delete_after = data.get('delete_after', True)
        res = fake_service.apply_to_attendance(photo_ids=ids, delete_after=delete_after)
        return jsonify(res)

    @bp.route('/api/supplement/delete-staging', methods=['POST'])
    def delete_staging():
        data = request.get_json() or {}
        ids = data.get('ids', [])
        deleted = fake_service.delete_staging_photos(photo_ids=ids)
        return jsonify(success=True, deleted_count=deleted)

    @bp.route('/api/supplement/download-staging/<photo_id>', methods=['GET'])
    def download_staging_single(photo_id):
        img_path = fake_service.get_staging_image_path(photo_id)
        if not img_path or not img_path.exists():
            return jsonify(error='Không tìm thấy ảnh'), 404
        return send_file(img_path, as_attachment=True, download_name=img_path.name)

    @bp.route('/api/supplement/download-staging-zip', methods=['POST'])
    def download_staging_zip():
        data = request.get_json() or {}
        ids = data.get('ids', [])
        items = fake_service.get_staging_photos()
        if ids:
            items = [it for it in items if it['id'] in ids]
        if not items:
            return jsonify(error='Không có ảnh để tải'), 400

        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as z:
            for it in items:
                p = fake_service.get_staging_image_path(it['id'])
                if p and p.exists():
                    safe_emp = re.sub(r'[<>:"/\\|?* ]', '_', it['employee'])
                    safe_time = it['target_time'].replace(':', '-')
                    z_name = f"{safe_emp}_{it['target_date']}_{safe_time}_{it['id']}{p.suffix}"
                    z.write(p, arcname=z_name)
        stream.seek(0)
        return send_file(stream, mimetype='application/zip', as_attachment=True, download_name='anh-fake-bo-sung.zip')

    @bp.route('/api/supplement/staging/<photo_id>/regenerate', methods=['POST'])
    def regenerate_staging(photo_id):
        data = request.get_json(silent=True) or {}
        confirmed_lines = data.get('confirmed_lines')
        if confirmed_lines is not None and not isinstance(confirmed_lines, list):
            return jsonify(status='invalid_request', error='confirmed_lines phải là danh sách'), 400
        result = fake_service.regenerate_staging_photo(
            photo_id, confirmed_lines=confirmed_lines
        )
        status = result.get('status') if isinstance(result, dict) else None
        if status == 'completed':
            return jsonify(success=True, status=status, item=result)
        if status == 'needs_confirmation':
            return jsonify(success=False, **result), 409
        if status == 'verification_failed':
            return jsonify(success=False, **result), 422
        if status in ('original_missing', 'not_found'):
            return jsonify(success=False, **result), 404
        return jsonify(success=False, status='failed', error='Không thể tạo lại ảnh này'), 500

    @bp.route('/api/supplement/check-ai', methods=['POST'])
    def check_ai_photo():
        from src.anti_ai_detector import inspect_ai_markers
        import uuid

        photo_id = None
        if request.is_json:
            data = request.get_json() or {}
            photo_id = data.get('photo_id')
        else:
            photo_id = request.form.get('photo_id')

        target_path = None
        if photo_id:
            target_path = fake_service.get_staging_image_path(photo_id)
        elif 'photo' in request.files:
            f = request.files['photo']
            temp_check = fake_service.temp_dir / f"check_{uuid.uuid4().hex[:8]}.jpg"
            f.save(temp_check)
            target_path = temp_check

        if not target_path or not target_path.exists():
            return jsonify(error='Không tìm thấy file ảnh để kiểm tra'), 404

        report = inspect_ai_markers(str(target_path))
        return jsonify(success=True, report=report)

    app.register_blueprint(bp)
