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
from src.supplement_evidence import EvidenceConnection

LABEL = 'Ảnh bổ sung — không xác nhận thời điểm chụp'


def register_batches(app, data_dir, registry_provider=None, matcher_provider=None, service=None):
    root = Path(data_dir) / 'supplement_output'
    root.mkdir(parents=True, exist_ok=True)
    db = Path(data_dir) / 'supplement_batches.sqlite3'

    def connect():
        c = sqlite3.connect(db, timeout=30, factory=EvidenceConnection)
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
        if not isinstance(value, str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}',value):
            raise ValueError('Ngày hồ sơ phải có định dạng YYYY-MM-DD')
        return date.fromisoformat(value).isoformat()

    def payload():
        data = request.get_json()
        if not isinstance(data, dict):
            raise ValueError('Nội dung JSON phải là object')
        return data

    @bp.errorhandler(ValueError)
    def invalid(e):
        return jsonify(success=False,error=str(e)), 400

    @bp.errorhandler(LookupError)
    def missing(e):
        return jsonify(success=False,error=str(e)), 404

    @bp.errorhandler(RuntimeError)
    def conflict(e):
        return jsonify(success=False,error=str(e)), 409

    @bp.route('/api/supplement/batches', methods=['GET', 'POST'])
    def batches():
        if request.method == 'GET':
            with connect() as c:
                return jsonify(batches=[json.loads(r[0]) for r in c.execute('SELECT document FROM batches ORDER BY rowid DESC')])
        files = request.files.getlist('photos')
        if not 1 <= len(files) <= 31:
            raise ValueError('Chọn từ 1 đến 31 ảnh')
        day = parse_day(request.form.get('supplement_date', ''))
        project_id = request.form.get('project_id', '').strip()
        employee_id = request.form.get('employee_id', '').strip()
        project, identity = fake_service._identity(project_id, employee_id, date.fromisoformat(day))
        employee = identity['display_name']
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
        b = {'id': bid, 'employee': employee, 'project_id': project_id, 'employee_id': employee_id, 'label': LABEL, 'status': 'stored',
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
            if b.get('project_id') and b.get('employee_id'):
                fake_service._identity(b['project_id'], b['employee_id'], date.fromisoformat(day))
            old = dict(item)
            item.update(supplement_date=day, note=note)
            b['checks_valid'] = False
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
        return jsonify(success=False, error='Tool cá nhân không cần duyệt; có thể tải hồ sơ trực tiếp.'), 410

    @bp.get('/api/supplement/batches/<bid>/download')
    def download(bid):
        with connect() as c:
            b = load(c, bid)
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
    fake_service = service or FakePhotoService(data_dir, registry_provider=registry_provider, matcher_provider=matcher_provider)
    app.extensions['supplement_service'] = fake_service
    from src.supplement_evidence import parse_target_time, selected_ids

    @bp.post('/api/supplement/records')
    def create_supplement_records():
        files = request.files.getlist('photos')
        project_id = request.form.get('project_id', '').strip()
        employee_id = request.form.get('employee_id', '').strip()
        if not project_id or not employee_id or not 1 <= len(files) <= 31:
            raise ValueError('Chọn dự án, nhân viên bằng ID và từ 1 đến 31 ảnh')
        try:
            configs = json.loads(request.form.get('configs', '[]'))
        except (TypeError, json.JSONDecodeError):
            raise ValueError('Cấu hình ngày bổ sung không hợp lệ')
        if not isinstance(configs, list) or len(configs) != len(files):
            raise ValueError('Mỗi ảnh cần có ngày đề nghị bổ sung')
        prepared = []
        total = 0
        for file, config in zip(files, configs):
            if not isinstance(config, dict):
                raise ValueError('Cấu hình ảnh không hợp lệ')
            day = parse_day(config.get('target_date'))
            target_time = parse_target_time(config.get('target_time'), provided='target_time' in config)
            project, employee = fake_service._identity(project_id, employee_id, date.fromisoformat(day))

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

        items = fake_service.create_records(prepared,project_id,employee_id,
                replace_timestamp=request.form.get('replace_timestamp','true').lower() in ('true','1'),
                modify_exif=request.form.get('modify_exif','true').lower() in ('true','1'))
        return jsonify(success=True, count=len(items), items=items), 201

    @bp.post('/api/supplement/records/<photo_id>/check')
    def check_record(photo_id):
        return jsonify(success=True, item=fake_service.inspect_record(photo_id))

    @bp.patch('/api/supplement/records/<photo_id>')
    def update_record(photo_id):
        data = payload()
        with fake_service._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            row = c.execute('SELECT * FROM staging_photos WHERE id=?',(photo_id,)).fetchone()
            if not row:
                raise LookupError('Không tìm thấy hồ sơ')
            item = dict(row)
            if item['applied_path']:
                raise RuntimeError('Hồ sơ đã áp dụng không được đổi nguồn hoặc danh tính')
            day = parse_day(data.get('target_date',item['target_date']))
            pid, eid = data.get('project_id',item['project_id']), data.get('employee_id',item['employee_id'])
            project, employee = fake_service._identity(pid,eid,date.fromisoformat(day))
            source_hash = item['source_sha256']
            if data.get('confirm_legacy_source') is True and not source_hash:
                source_hash = hashlib.sha256(fake_service._source_path(item).read_bytes()).hexdigest()
                c.execute('UPDATE staging_photos SET source_provenance=? WHERE id=?',('legacy_baseline',photo_id))
                fake_service._audit(c,photo_id,'legacy_source_baseline',{'sha256':source_hash,'not_original_intake_hash':True})
            target_time = parse_target_time(data.get('target_time',item['target_time']), provided='target_time' in data)
            c.execute('UPDATE staging_photos SET project_id=?,employee_id=?,project=?,employee=?,target_date=?,target_time=?,source_sha256=?,validation_json=? WHERE id=?',
                      (pid,eid,project['storage_dir'],employee['display_name'],day,target_time,source_hash,'{}',photo_id))
            fake_service._audit(c,photo_id,'updated',{'before':item,'requested':data})
        return jsonify(success=True,item=fake_service.inspect_record(photo_id))

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
        # Legacy route shares the same IDs, date/time and upload validation.
        return create_supplement_records()

    @bp.route('/api/supplement/staging', methods=['GET'])
    def get_staging():
        proj = request.args.get('project')
        emp = request.args.get('employee')
        items = fake_service.get_staging_photos(project=proj, employee=emp,
                    project_id=request.args.get('project_id'), employee_id=request.args.get('employee_id'))
        return jsonify(success=True, count=len(items), items=items)

    @bp.route('/api/supplement/staging/<photo_id>/image', methods=['GET'])
    def get_staging_image(photo_id):
        img_path = fake_service.get_staging_image_path(photo_id)
        if not img_path or not img_path.exists():
            return jsonify(error='Không tìm thấy ảnh'), 404
        return send_file(img_path)

    @bp.get('/api/supplement/staging/<photo_id>/original')
    def original_image(photo_id):
        with fake_service._connect() as c:
            row = c.execute('SELECT * FROM staging_photos WHERE id=?',(photo_id,)).fetchone()
        if not row:
            raise LookupError('Không tìm thấy hồ sơ')
        item = dict(row)
        if item.get('intake_operation_id'):
            with fake_service._connect() as c:
                op=c.execute('SELECT state FROM supplement_operations WHERE id=?',(item['intake_operation_id'],)).fetchone()
            if not op or op[0] != 'committed':
                raise LookupError('Đợt tạo hồ sơ chưa hoàn tất')
        path = fake_service._source_path(item)
        if not path.is_file():
            raise LookupError('Không còn ảnh gốc')
        if item['source_sha256'] and hashlib.sha256(path.read_bytes()).hexdigest() != item['source_sha256']:
            raise RuntimeError('Ảnh gốc đã thay đổi so với lúc tiếp nhận')
        return send_file(path, as_attachment=request.args.get('download') == '1', download_name=path.name)

    @bp.route('/api/supplement/staging/<photo_id>/watermark-crop', methods=['GET'])
    def get_staging_watermark_crop(photo_id):
        crop = fake_service.get_watermark_crop_bytes(photo_id)
        if not crop:
            return jsonify(error='Không tìm thấy crop watermark từ ảnh gốc'), 404
        return send_file(io.BytesIO(crop), mimetype='image/jpeg')

    @bp.route('/api/supplement/apply-to-attendance', methods=['POST'])
    def apply_to_attendance():
        data = payload()
        ids = selected_ids(data.get('ids'))
        delete_after = data.get('delete_after', True)
        res = fake_service.apply_to_attendance(photo_ids=ids, delete_after=delete_after)
        return jsonify(res), (200 if res.get('success') else 422)

    @bp.route('/api/supplement/delete-staging', methods=['POST'])
    def delete_staging():
        data = payload()
        ids = selected_ids(data.get('ids'))
        deleted = fake_service.delete_staging_photos(photo_ids=ids)
        return jsonify(success=True, deleted_count=deleted)

    @bp.route('/api/supplement/download-staging/<photo_id>', methods=['GET'])
    def download_staging_single(photo_id):
        img_path = fake_service.get_staging_image_path(photo_id)
        if not img_path or not img_path.exists():
            return jsonify(error='Không tìm thấy ảnh'), 404
        img_path, raw = fake_service.read_verified_derived(photo_id)
        return send_file(io.BytesIO(raw), as_attachment=True, download_name=img_path.name)

    @bp.route('/api/supplement/download-staging-zip', methods=['POST'])
    def download_staging_zip():
        data = payload()
        ids = selected_ids(data.get('ids'))
        lookup = {item['id']:item for item in fake_service.get_staging_photos()}
        if any(photo_id not in lookup for photo_id in ids):
            raise RuntimeError('Danh sách có hồ sơ không tồn tại hoặc đã áp dụng')
        prepared = [(lookup[photo_id], *fake_service.read_verified_derived(photo_id)) for photo_id in ids]
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as z:
            for item, path, raw in prepared:
                safe_emp = re.sub(r'[<>:"/\\|?* ]', '_', item['employee'])
                safe_time = (item['target_time'] or '').replace(':', '-')
                name = f"{safe_emp}_{item['target_date']}_{safe_time}_{item['id']}{path.suffix}"
                z.writestr(name, raw)
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
        if status == 'invalid_request':
            return jsonify(success=False, **result), 400
        if status == 'integrity_failed':
            return jsonify(success=False, **result), 409
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
