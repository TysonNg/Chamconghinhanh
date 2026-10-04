"""Immutable supplementary evidence, validation, and recoverable file publication."""
import hashlib
import json
import os
import re
import shutil
import sqlite3
import uuid
import threading
from functools import wraps
from datetime import datetime, timezone
from pathlib import Path

_LOCKS = {}
_LOCKS_GUARD = threading.Lock()
_EVIDENCE_DBS = {}
_LOCK_DEPTH = threading.local()


def serialized(method):
    """Serialize publication/deletion/recovery across threads and app processes."""
    @wraps(method)
    def run(self, *args, **kwargs):
        key = str(self.db_path.resolve())
        with _LOCKS_GUARD:
            lock = _LOCKS.setdefault(key, threading.RLock())
        with lock:
            active = getattr(_LOCK_DEPTH, "active", set())
            if key in active:
                return method(self, *args, **kwargs)
            with (self.data_dir / ".supplement-write.lock").open("a+b") as handle:
                handle.seek(0, 2)
                if handle.tell() == 0:
                    handle.write(b"0")
                    handle.flush()
                handle.seek(0)
                try:
                    if os.name == "nt":
                        import msvcrt
                        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                    else:
                        import fcntl
                        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                except OSError as exc:
                    raise RuntimeError("Đang có thao tác ảnh khác; hãy thử lại") from exc
                try:
                    _LOCK_DEPTH.active = active | {key}
                    return method(self, *args, **kwargs)
                finally:
                    _LOCK_DEPTH.active = active
                    handle.seek(0)
                    if os.name == "nt":
                        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                    else:
                        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
    return run

from PIL import Image
from src.attendance_dates import image_date_status, parse_attendance_date, canonical_day_path, resolve_day_folder


class EvidenceConnection(sqlite3.Connection):
    """Commit/rollback and close deterministically (SQLite locks persist on Windows)."""
    def __exit__(self, *args):
        try:
            return super().__exit__(*args)
        finally:
            self.close()


def now():
    return datetime.now(timezone.utc).isoformat()


def parse_target_time(value, *, provided=False):
    if value is None and provided:
        raise ValueError('Giờ không được là null; bỏ trường nếu chưa nhập giờ')
    if value is None or value == '':
        return None
    if not isinstance(value, str) or not re.fullmatch(r'\d{2}:\d{2}(:\d{2})?', value.strip()):
        raise ValueError('Giờ phải có định dạng HH:MM hoặc HH:MM:SS')
    parts = [int(p) for p in value.strip().split(':')]
    if parts[0] > 23 or parts[1] > 59 or (len(parts) == 3 and parts[2] > 59):
        raise ValueError('Giờ không hợp lệ')
    return value.strip() + (':00' if len(parts) == 2 else '')


def capture_datetime(raw):
    import io
    try:
        with Image.open(io.BytesIO(raw)) as image:
            exif = image.getexif()
            value = exif.get(36867) or exif.get_ifd(34665).get(36867)
            if isinstance(value, bytes):
                value = value.decode('ascii')
            if value:
                return datetime.strptime(value.rstrip('\x00'), '%Y:%m:%d %H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
    except (OSError, ValueError, KeyError, AttributeError, UnicodeError):
        pass
    return None


def selected_ids(value):
    if not isinstance(value, list) or not value or len(value) > 1000 or not all(isinstance(i, str) and i.strip() for i in value):
        raise ValueError('Chọn rõ danh sách ID ảnh cần thao tác (tối đa 1000)')
    return list(dict.fromkeys(value))


def registered_supplement(path):
    """Identify our retained records even when their metadata is missing/broken."""
    path = Path(path).resolve()
    if not re.fullmatch(r'[a-f0-9]{12}', path.stem):
        return False
    databases = [db for root,db in list(_EVIDENCE_DBS.items()) if path.is_relative_to(root)]
    for db in databases:
        try:
            with sqlite3.connect(Path(db).as_uri() + '?mode=ro', uri=True, factory=EvidenceConnection) as c:
                if c.execute('SELECT 1 FROM staging_photos WHERE id=?',(path.stem,)).fetchone():
                    return True
        except sqlite3.Error:
            return True
    return False


def evidence_is_visible(path):
    """Readers must not consume files from an uncommitted apply operation."""
    path = Path(path)
    sidecar = path.with_suffix(path.suffix + '.json')
    if not sidecar.exists():
        return not registered_supplement(path)
    try:
        meta = json.loads(sidecar.read_text(encoding='utf-8'))
        if not isinstance(meta, dict):
            return not registered_supplement(path)
        operation_id = meta.get('operation_id')
        if not operation_id:
            return meta.get('source') != 'supplement_original' and not registered_supplement(path)
        input_root = next((p for p in path.parents if p.name == 'input_images'), None)
        db = Path(meta['operation_db']).resolve()
        if input_root is None or not db.is_relative_to(input_root.parent.resolve()) or not db.is_file():
            return False
        with sqlite3.connect(db.as_uri() + '?mode=ro', uri=True, factory=EvidenceConnection) as c:
            row = c.execute('SELECT state FROM supplement_operations WHERE id=?', (operation_id,)).fetchone()
        return bool(row and row[0] == 'committed' and meta.get('sha256') == hashlib.sha256(path.read_bytes()).hexdigest())
    except (OSError, ValueError, KeyError, sqlite3.Error, TypeError):
        return not registered_supplement(path) and not bool(locals().get('operation_id'))


class SupplementEvidenceMixin:
    @serialized
    def _backup_evidence(self):
        if not self.db_path.exists():
            return
        with self._connect() as c:
            table = c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='supplement_schema'").fetchone()
            if table and c.execute('SELECT 1 FROM supplement_schema WHERE version=2').fetchone():
                return
        backup = self.data_dir / 'backups' / ('before-evidence-v1-' + uuid.uuid4().hex)
        backup.mkdir(parents=True)
        with self._connect() as source, sqlite3.connect(backup / self.db_path.name,factory=EvidenceConnection) as target:
            source.backup(target)
        archive_db = self.data_dir / 'supplement_batches.sqlite3'
        if archive_db.exists():
            with sqlite3.connect(archive_db, factory=EvidenceConnection) as source, sqlite3.connect(backup / archive_db.name, factory=EvidenceConnection) as target:
                source.backup(target)
        for folder in (self.raw_dir, self.staging_dir, self.data_dir / 'supplement_output'):
            if folder.exists():
                shutil.copytree(folder, backup / folder.name)

    @serialized
    def _init_evidence(self):
        _EVIDENCE_DBS[self.input_images_dir.resolve()] = self.db_path.resolve()
        with self._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            c.execute('CREATE TABLE IF NOT EXISTS supplement_schema(version INTEGER PRIMARY KEY, installed_at TEXT NOT NULL)')
            columns = {r['name']: r for r in c.execute('PRAGMA table_info(staging_photos)')}
            additions = {'project_id': "TEXT NOT NULL DEFAULT ''", 'employee_id': "TEXT NOT NULL DEFAULT ''",
                         'source_sha256': "TEXT NOT NULL DEFAULT ''", 'derived_sha256': "TEXT NOT NULL DEFAULT ''",
                         'validation_json': "TEXT NOT NULL DEFAULT '{}'", 'exif_status': "TEXT NOT NULL DEFAULT 'legacy'",
                         'intake_operation_id': "TEXT NOT NULL DEFAULT ''", 'source_provenance': "TEXT NOT NULL DEFAULT 'legacy_unverified'"}
            for column, declaration in additions.items():
                if column not in columns:
                    c.execute(f'ALTER TABLE staging_photos ADD COLUMN {column} {declaration}')
            columns = list(c.execute('PRAGMA table_info(staging_photos)'))
            if next(r for r in columns if r['name'] == 'target_time')['notnull']:
                definitions = []
                for r in columns:
                    definition = '"' + r['name'] + '" ' + r['type']
                    if r['pk']:
                        definition += ' PRIMARY KEY'
                    if r['notnull'] and r['name'] != 'target_time':
                        definition += ' NOT NULL'
                    if r['dflt_value'] is not None:
                        definition += ' DEFAULT ' + r['dflt_value']
                    definitions.append(definition)
                c.execute('CREATE TABLE staging_photos_v1 (' + ','.join(definitions) + ')')
                names = ','.join('"' + r['name'] + '"' for r in columns)
                c.execute(f'INSERT INTO staging_photos_v1 ({names}) SELECT {names} FROM staging_photos')
                c.execute('DROP TABLE staging_photos')
                c.execute('ALTER TABLE staging_photos_v1 RENAME TO staging_photos')
            c.execute("UPDATE staging_photos SET target_time=NULL WHERE target_time=''")
            c.execute('CREATE INDEX IF NOT EXISTS staging_source_hash ON staging_photos(source_sha256)')
            c.execute('CREATE TABLE IF NOT EXISTS supplement_operations(id TEXT PRIMARY KEY, state TEXT NOT NULL, document TEXT NOT NULL, created_at TEXT NOT NULL)')
            c.execute('CREATE TABLE IF NOT EXISTS supplement_history(id INTEGER PRIMARY KEY, record_id TEXT NOT NULL, action TEXT NOT NULL, at TEXT NOT NULL, details TEXT NOT NULL)')
            c.execute('CREATE TABLE IF NOT EXISTS applied_evidence(project_id TEXT NOT NULL, target_date TEXT NOT NULL, source_sha256 TEXT NOT NULL, employee_id TEXT NOT NULL, dest_path TEXT NOT NULL, operation_id TEXT NOT NULL, PRIMARY KEY(project_id,target_date,source_sha256))')
            c.execute('INSERT OR IGNORE INTO supplement_schema VALUES(2,?)', (now(),))
        self._recover_operations()

    def _audit(self, c, record_id, action, details):
        c.execute('INSERT INTO supplement_history(record_id,action,at,details) VALUES(?,?,?,?)',
                  (record_id, action, now(), json.dumps(details, ensure_ascii=False)))

    def _source_path(self, item):
        name = item.get('original_file_name') or f"{item['id']}_raw.jpg"
        path = self.raw_dir / name
        if not path.resolve().is_relative_to(self.raw_dir.resolve()) or path.is_symlink():
            raise ValueError('Đường dẫn ảnh gốc không hợp lệ')
        return path

    def _identity(self, project_id, employee_id, day):
        registry = self.registry_provider() if self.registry_provider else None
        if not registry or not isinstance(project_id,str) or not project_id.strip() or not isinstance(employee_id,str) or not employee_id.strip():
            raise ValueError('Chọn dự án và nhân viên bằng ID; hồ sơ cũ cần gắn lại danh tính')
        project = registry.get_project(project_id)
        employee = registry.get_employee(employee_id)
        if project['project_id'] != project_id:
            raise ValueError('Dự án phải được chọn bằng ID, không dùng tên')
        if not project['active']:
            raise ValueError('Dự án đã được lưu trữ')
        if not registry.is_member(project_id, employee_id, day):
            raise ValueError('Nhân viên không thuộc dự án tại ngày bổ sung')
        return project, employee

    def _face_context(self, item, matcher):
        day = parse_attendance_date(item['target_date'])
        return {'model':matcher.model_name, 'detector':matcher.detector_backend,
                'metric':matcher.distance_metric, 'threshold':matcher._get_default_threshold(),
                'source_sha256':item['source_sha256'],
                'portraits':[{'path':str(p),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
                    for p in self.registry_provider().portrait_paths(item['project_id'],item['employee_id'],day)]}

    def _validate_evidence(self, item, *, check_face=False):
        result = {'date_status':'unknown', 'face_status':'not_run', 'integrity_status':'unverified',
                  'storage_status':'available', 'can_apply':False, 'reasons':[],
                  'processing':{'watermark':item.get('watermark_status','legacy'), 'exif':item.get('exif_status','legacy')}}
        try:
            original = self._source_path(item)
            if not original.is_file():
                result['storage_status'] = 'original_missing'
                raise ValueError('Không còn ảnh gốc')
            derived = self.staging_dir / item['file_name']
            if not derived.is_file():
                result['storage_status'] = 'derived_missing'
                result['processing'] = {'watermark':'missing', 'exif':'missing'}
            elif item.get('derived_sha256') and hashlib.sha256(derived.read_bytes()).hexdigest() != item['derived_sha256']:
                result['storage_status'] = 'derived_changed'
                result['processing'] = {'watermark':'verification_failed', 'exif':'unknown'}
            raw = original.read_bytes()
            actual_hash = hashlib.sha256(raw).hexdigest()
            if not item.get('source_sha256'):
                raise ValueError('Hồ sơ cũ chưa có hash tiếp nhận; cần xác nhận nguồn')
            if actual_hash != item['source_sha256']:
                result['integrity_status'] = 'mismatch'
                raise ValueError('Ảnh gốc đã thay đổi so với lúc tiếp nhận')
            result['integrity_status'] = 'baseline_verified' if item.get('source_provenance')=='legacy_baseline' else 'verified'
            target = parse_attendance_date(item['target_date'])
            result['date_status'] = image_date_status(original, target)
            inspection = json.loads(item.get('inspection_json') or '{}')
            visible = inspection.get('original_visible_date')
            if visible:
                if parse_attendance_date(visible) != target:
                    result['date_status'] = 'mismatch'
                elif result['date_status'] == 'unknown':
                    result['date_status'] = 'consistent'
            project, employee = self._identity(item.get('project_id'), item.get('employee_id'), target)
            if check_face:
                matcher = self.matcher_provider() if self.matcher_provider else None
                if matcher is None:
                    result['face_status'] = 'error'
                    result['reasons'].append('Bộ nhận diện chưa sẵn sàng')
                else:
                    context = self._face_context(item, matcher)
                    match = matcher.match_employee_in_photo(project_id=project['project_id'],
                            employee_id=employee['employee_id'], attendance_date=target, image_path=str(original))
                    result['face_status'] = match.status
                    result['face_distance'] = match.distance
                    if self._face_context(item,matcher) != context:
                        raise RuntimeError('Chân dung hoặc model thay đổi trong lúc kiểm tra')
                    result['face_evidence'] = {**context, 'checked_at':now()}
                    if match.status != 'matched':
                        result['reasons'].append(match.reason or 'Khuôn mặt chưa đủ điều kiện')
            else:
                previous = json.loads(item.get('validation_json') or '{}')
                result['face_status'] = previous.get('face_status', 'not_run')
                result['face_distance'] = previous.get('face_distance')
                saved_context = dict(previous.get('face_evidence') or {})
                saved_context.pop('checked_at',None)
                if saved_context:
                    matcher = self.matcher_provider() if self.matcher_provider else None
                    if matcher is None or saved_context != self._face_context(item,matcher):
                        result['face_status'] = 'not_run'
                        result['reasons'].append('Nguồn/chân dung/model đã thay đổi; cần kiểm tra lại')
                elif result['face_status'] == 'matched':
                    result['face_status'] = 'not_run'
            if result['date_status'] != 'consistent':
                result['reasons'].append('Ngày ảnh gốc lệch hồ sơ' if result['date_status'] == 'mismatch' else 'Chưa xác định được ngày ảnh gốc')
            result['can_apply'] = result['date_status'] == 'consistent' and result['face_status'] == 'matched'
            if item.get('intake_operation_id'):
                with self._connect() as c:
                    intake=c.execute('SELECT state FROM supplement_operations WHERE id=?',(item['intake_operation_id'],)).fetchone()
                if not intake or intake[0] != 'committed':
                    result['can_apply'] = False
                    result['reasons'].append('Đợt tạo hồ sơ chưa hoàn tất')
        except (ValueError, OSError, KeyError, TypeError, AttributeError) as exc:
            result['reasons'].append(str(exc))
            if result['integrity_status'] in ('verified','baseline_verified'):
                result['face_status'] = 'identity_unresolved'
        except Exception as exc:
            result['face_status'] = 'error'
            result['reasons'].append('Lỗi kiểm tra: ' + type(exc).__name__)
        return result

    def _public_record(self, item, validation=None):
        item = dict(item)
        validation = validation or self._validate_evidence(item)
        for column, key in [('inspection_json','inspection'), ('watermark_ocr_json','watermark_ocr'),
                            ('confirmed_watermark_json','confirmed_watermark'), ('generation_meta_json','generation_meta')]:
            if column in item:
                try:
                    item[key] = json.loads(item.pop(column) or '{}')
                except (ValueError, TypeError):
                    item[key] = {'status':'invalid_metadata'}
        item.update(validation)
        item.pop('validation_json', None)
        item['target_time'] = item.get('target_time') or None
        path = self.staging_dir / item['file_name']
        item['url'] = f"/api/supplement/staging/{item['id']}/image" if path.is_file() else None
        item['derived_url'] = item['url']
        try:
            item['original_url'] = f"/api/supplement/staging/{item['id']}/original" if self._source_path(item).is_file() else None
        except ValueError:
            item['original_url'] = None
        item['size_kb'] = round(path.stat().st_size / 1024, 1) if path.is_file() else 0
        return item

    @serialized
    def inspect_record(self, photo_id, *, check_face=True):
        with self._connect() as c:
            row = c.execute('SELECT * FROM staging_photos WHERE id=?', (photo_id,)).fetchone()
        if not row:
            raise LookupError('Không tìm thấy hồ sơ')
        item = dict(row)
        result = self._validate_evidence(item, check_face=check_face)
        if check_face:
            with self._connect() as c:
                c.execute('UPDATE staging_photos SET validation_json=? WHERE id=?', (json.dumps(result,ensure_ascii=False), photo_id))
                self._audit(c, photo_id, 'checked', result)
        return self._public_record(item, result)

    def get_staging_photos(self, project=None, employee=None, *, project_id=None, employee_id=None):
        query = "SELECT * FROM staging_photos WHERE COALESCE(applied_at,'')='' AND (intake_operation_id='' OR EXISTS(SELECT 1 FROM supplement_operations o WHERE o.id=staging_photos.intake_operation_id AND o.state='committed'))"
        params = []
        for column, value in [('project',project),('employee',employee),('project_id',project_id),('employee_id',employee_id)]:
            if value:
                query += ' AND ' + column + '=?'
                params.append(value)
        with self._connect() as c:
            rows = c.execute(query + ' ORDER BY created_at DESC', params).fetchall()
        return [self._public_record(dict(row)) for row in rows]

    def _reserve_intake_file(self, operation_id, path, raw):
        path = Path(path)
        if not operation_id:
            with path.open('xb') as handle:
                handle.write(raw)
            return
        stage = self.temp_dir / operation_id
        stage.mkdir(exist_ok=True)
        temp = stage / path.name
        with temp.open('xb') as handle:
            handle.write(raw)
            handle.flush()
            os.fsync(handle.fileno())
        stat = temp.stat()
        with self._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            row=c.execute('SELECT document,state FROM supplement_operations WHERE id=?',(operation_id,)).fetchone()
            if not row or row['state'] != 'prepared':
                raise RuntimeError('Đợt tiếp nhận không còn hoạt động')
            doc=json.loads(row['document'])
            doc['files'].append({'dest':str(path),'stage_path':str(temp),'device':stat.st_dev,'inode':stat.st_ino})
            c.execute('UPDATE supplement_operations SET document=? WHERE id=?',(json.dumps(doc),operation_id))
        os.link(temp,path)  # Atomic exclusive publication, never overwrite a foreign file.

    @serialized
    def create_records(self, prepared, project_id, employee_id, *, replace_timestamp=True, modify_exif=True):
        self._recover_operations()
        operation={'id':uuid.uuid4().hex,'kind':'create','files':[],
                   'record_ids':[uuid.uuid4().hex[:12] for _ in prepared]}
        with self._connect() as c:
            c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',(operation['id'],'prepared',json.dumps(operation),now()))
        try:
            items=[]
            for (raw,day,target_time,name),record_id in zip(prepared,operation['record_ids']):
                project,employee=self._identity(project_id,employee_id,parse_attendance_date(day))
                item=self.create_supplement_photo(raw,project['storage_dir'],employee['display_name'],day,name,target_time,
                    project_id=project_id,employee_id=employee_id,replace_timestamp=replace_timestamp,
                    modify_exif=modify_exif,record_id=record_id,intake_operation_id=operation['id'])
                items.append(item)
                with (self.staging_dir / item['file_name']).open('rb+') as handle:
                    os.fsync(handle.fileno())
            with self._connect() as c:
                c.execute("UPDATE supplement_operations SET state='committed' WHERE id=?",(operation['id'],))
                for record_id in operation['record_ids']:
                    self._audit(c,record_id,'intake_committed',{'operation_id':operation['id']})
            for item in items:
                item['reasons']=[reason for reason in item['reasons'] if reason!='Đợt tạo hồ sơ chưa hoàn tất']
                item['can_apply']=item['date_status']=='consistent' and item['face_status']=='matched' and not item['reasons']
            shutil.rmtree(self.temp_dir/operation['id'],ignore_errors=True)
            return items
        except Exception as exc:
            self._recover_operations()
            raise RuntimeError('Chưa lưu đợt ảnh; đã phục hồi toàn bộ: '+str(exc)) from exc

    @serialized
    def mutate_attendance(self, action):
        return action()

    def _remove_owned_files(self, doc, root):
        for entry in doc.get('files',[]):
            path=Path(entry['dest'])
            if not path.resolve().is_relative_to(root.resolve()) or path.is_symlink():
                raise RuntimeError('Đường dẫn phục hồi không hợp lệ')
            if path.exists() and 'inode' in entry:
                stat=path.stat()
                if (stat.st_dev,stat.st_ino)==(entry['device'],entry['inode']):
                    path.unlink()

    def _owned_cleanup(self, document):
        # Older journals cannot establish image ownership: preserve evidence for repair.
        for entry in document.get('files', []):
            dest = Path(entry['dest'])
            sidecar = dest.with_suffix(dest.suffix + '.json')
            if dest.exists() and 'inode' not in entry and sidecar.exists():
                metadata = json.loads(sidecar.read_text(encoding='utf-8'))
                if metadata.get('operation_id') == document['id']:
                    raise RuntimeError('Unresolved recovery: missing file ownership evidence in older journal')
        for entry in document.get('files', []):
            dest = Path(entry['dest'])
            if not dest.resolve().is_relative_to(self.input_images_dir.resolve()):
                raise RuntimeError('Không thể phục hồi đường dẫn ngoài thư mục ảnh')
            sidecar = dest.with_suffix(dest.suffix + '.json')
            # Match the staged file identity. A sidecar does not prove ownership of the image.
            if dest.exists() and 'inode' in entry:
                stat=dest.stat()
                if (stat.st_dev,stat.st_ino)==(entry['device'],entry['inode']):
                    dest.unlink()
            if sidecar.exists():
                metadata=json.loads(sidecar.read_text(encoding='utf-8'))
                if metadata.get('operation_id')==document['id']:
                    sidecar.unlink()

    def _recover_operations(self):
        with self._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            for row in c.execute("SELECT * FROM supplement_operations WHERE state='prepared'").fetchall():
                doc = json.loads(row['document'])
                if doc.get('kind') == 'create':
                    self._remove_owned_files(doc,self.data_dir)
                    c.execute('DELETE FROM staging_photos WHERE intake_operation_id=?',(row['id'],))
                elif doc.get('kind') == 'delete':
                    for entry in doc['files']:
                        source, trash = Path(entry['source']), Path(entry['trash'])
                        if not source.resolve().is_relative_to(self.data_dir.resolve()) or not trash.resolve().is_relative_to(self.temp_dir.resolve()):
                            raise RuntimeError('Đường dẫn phục hồi xóa không hợp lệ')
                        if trash.exists():
                            if 'inode' in entry and (trash.stat().st_dev,trash.stat().st_ino)!=(entry['device'],entry['inode']):
                                raise RuntimeError('File phục hồi không thuộc thao tác xóa')
                            if source.exists():
                                raise RuntimeError('Nguồn đã tồn tại; không ghi đè khi phục hồi')
                            os.replace(trash, source)
                else:
                    self._owned_cleanup(doc)
                c.execute("UPDATE supplement_operations SET state='rolled_back' WHERE id=?", (row['id'],))
                self._audit(c, '', 'recovered', {'operation_id':row['id']})
                stage=self.temp_dir/row['id']
                if stage.resolve().is_relative_to(self.temp_dir.resolve()):
                    shutil.rmtree(stage,ignore_errors=True)

        with self._connect() as c:
            for row in c.execute("SELECT * FROM supplement_operations WHERE state='committed'").fetchall():
                doc=json.loads(row['document'])
                if doc.get('kind')=='delete':
                    for entry in doc['files']:
                        trash=Path(entry['trash'])
                        if not trash.resolve().is_relative_to(self.temp_dir.resolve()) or trash.is_symlink():
                            raise RuntimeError('Đường dẫn dọn dẹp không hợp lệ')
                        if trash.exists() and 'inode' in entry and (trash.stat().st_dev,trash.stat().st_ino)==(entry['device'],entry['inode']):
                            try:
                                trash.unlink()
                            except OSError:
                                pass  # Retry next recovery; original deletion is already committed.
                stage=self.temp_dir/row['id']
                if stage.resolve().is_relative_to(self.temp_dir.resolve()):
                    shutil.rmtree(stage,ignore_errors=True)

    @serialized
    def apply_to_attendance(self, photo_ids, delete_after=True):
        ids = selected_ids(photo_ids)
        self._recover_operations()
        with self._connect() as c:
            rows = c.execute('SELECT * FROM staging_photos WHERE id IN ('+','.join('?' for _ in ids)+')',ids).fetchall()
        found = {r['id'] for r in rows}
        rejected = [{'id':i,'reason':'Không tìm thấy hồ sơ'} for i in ids if i not in found]
        validated = []
        for row in rows:
            item = dict(row)
            check = self._validate_evidence(item, check_face=True)
            item['validation_json'] = json.dumps(check,ensure_ascii=False)
            with self._connect() as c:
                c.execute('UPDATE staging_photos SET validation_json=? WHERE id=?', (item['validation_json'],item['id']))
                self._audit(c,item['id'],'checked',check)
            if not check['can_apply']:
                rejected.append({'id':item['id'], **check, 'reason':'; '.join(check['reasons'])})
            else:
                validated.append(item)
        if rejected:
            return {'success':False,'count':0,'rejected':rejected,'message':'Chưa áp dụng: '+ '; '.join(r['reason'] for r in rejected)}
        operation = {'id':uuid.uuid4().hex,'kind':'apply','files':[]}
        applied, pending, groups = [], [], {}
        stage=self.temp_dir/operation['id']
        try:
            stage.mkdir()
            with self._connect() as c:
                c.execute('BEGIN IMMEDIATE')
                for snapshot in validated:
                    row = c.execute('SELECT * FROM staging_photos WHERE id=?',(snapshot['id'],)).fetchone()
                    if not row:
                        raise RuntimeError('Hồ sơ đã bị xóa trong lúc kiểm tra')
                    item = dict(row)
                    source = self._source_path(item)
                    raw = source.read_bytes()
                    if hashlib.sha256(raw).hexdigest() != item['source_sha256'] or item['source_sha256'] != snapshot['source_sha256']:
                        raise RuntimeError('Ảnh gốc thay đổi trong lúc kiểm tra')
                    day = parse_attendance_date(item['target_date'])
                    project, employee = self._identity(item['project_id'],item['employee_id'],day)
                    key = (item['project_id'],item['target_date'],item['source_sha256'])
                    existing = c.execute('SELECT * FROM applied_evidence WHERE project_id=? AND target_date=? AND source_sha256=?',key).fetchone()
                    duplicate = groups.get(key)
                    if existing or duplicate:
                        prior = dict(existing) if existing else duplicate
                        if prior['employee_id'] != item['employee_id']:
                            raise RuntimeError('Cùng ảnh gốc đã gắn với nhân viên khác trong dự án/ngày')
                        dest = Path(prior['dest_path'])
                        if existing and (not dest.is_file() or hashlib.sha256(dest.read_bytes()).hexdigest() != item['source_sha256'] or not evidence_is_visible(dest)):
                            raise RuntimeError('Ảnh đích đã mất hoặc thay đổi; cần kiểm tra')
                    else:
                        root = self.input_images_dir / project['storage_dir']
                        resolution = resolve_day_folder(root, day)
                        if resolution.status not in ('found','missing'):
                            raise RuntimeError(resolution.reason)
                        folder = resolution.path or canonical_day_path(root, day)
                        dest = folder / (item['id'] + source.suffix)
                        if not dest.resolve().is_relative_to(self.input_images_dir.resolve()) or dest.is_symlink():
                            raise RuntimeError('Đường dẫn áp dụng không hợp lệ')
                        if dest.exists() or dest.with_suffix(dest.suffix+'.json').exists():
                            raise RuntimeError('Đích đã tồn tại nhưng chưa có nhật ký áp dụng; cần kiểm tra')
                        temp=stage/dest.name
                        with temp.open('xb') as handle:
                            handle.write(raw)
                            handle.flush()
                            os.fsync(handle.fileno())
                        stat=temp.stat()
                        metadata={'derived':False,'source':'supplement_original','record_id':item['id'],
                            'project_id':item['project_id'],'employee_id':item['employee_id'], 'requested_date':item['target_date'],
                            'sha256':item['source_sha256'], 'operation_id':operation['id'],'operation_db':str(self.db_path.resolve()),
                            'original_visible_date':json.loads(item['inspection_json'] or '{}').get('original_visible_date')}
                        meta_temp=stage/(dest.name+'.json')
                        meta_temp.write_text(json.dumps(metadata,ensure_ascii=False),encoding='utf-8')
                        with meta_temp.open('rb+') as handle:
                            os.fsync(handle.fileno())
                        operation['files'].append({'dest':str(dest),'sha256':item['source_sha256'],
                            'stage_path':str(temp),'device':stat.st_dev,'inode':stat.st_ino})
                        pending.append((item,raw,dest))
                        groups[key] = {'employee_id':item['employee_id'],'dest_path':str(dest)}
                    applied.append({'id':item['id'],'dest_path':str(dest),'filename':dest.name,'project':project['storage_dir'],'day':item['target_date']})
                c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)', (operation['id'],'prepared',json.dumps(operation),now()))
            # The persistent prepare journal precedes publication. BEGIN IMMEDIATE serializes apply/delete/recovery.
            with self._connect() as c:
                c.execute('BEGIN IMMEDIATE')
                op_state=c.execute('SELECT state FROM supplement_operations WHERE id=?',(operation['id'],)).fetchone()[0]
                if op_state != 'prepared':
                    raise RuntimeError('Thao tác đã được phục hồi; hãy thử lại')
                for item in validated:
                    current=c.execute('SELECT * FROM staging_photos WHERE id=?',(item['id'],)).fetchone()
                    if not current or dict(current) != item:
                        raise RuntimeError('Hồ sơ thay đổi trước khi công bố; hãy kiểm tra lại')
                    self._identity(item['project_id'],item['employee_id'],parse_attendance_date(item['target_date']))
                    saved_context=dict(json.loads(item['validation_json'])['face_evidence'])
                    saved_context.pop('checked_at',None)
                    matcher=self.matcher_provider()
                    if saved_context != self._face_context(item,matcher):
                        raise RuntimeError('Chân dung hoặc model thay đổi trước khi công bố')
                    if hashlib.sha256(self._source_path(item).read_bytes()).hexdigest()!=item['source_sha256']:
                        raise RuntimeError('Ảnh gốc thay đổi trước khi công bố')
                for item,raw,dest in pending:
                    dest.parent.mkdir(parents=True,exist_ok=True)
                    sidecar=dest.with_suffix(dest.suffix+'.json')
                    os.link(stage/(dest.name+'.json'),sidecar)
                    os.link(stage/dest.name,dest)
                    c.execute('INSERT INTO applied_evidence VALUES(?,?,?,?,?,?)',
                        (item['project_id'],item['target_date'],item['source_sha256'],item['employee_id'],str(dest),operation['id']))
                new_paths={str(dest) for _,_,dest in pending}
                for entry in applied:
                    dest=Path(entry['dest_path'])
                    item=next(i for i in validated if i['id']==entry['id'])
                    if not dest.is_file() or hashlib.sha256(dest.read_bytes()).hexdigest()!=item['source_sha256']:
                        raise RuntimeError('Ảnh đích thay đổi trước khi hoàn tất')
                    if str(dest) not in new_paths and not evidence_is_visible(dest):
                        raise RuntimeError('Ảnh đích chưa đủ bằng chứng commit')
                for entry in applied:
                    c.execute("UPDATE staging_photos SET applied_path=?,applied_at=CASE WHEN COALESCE(applied_at,'')='' THEN ? ELSE applied_at END WHERE id=?",
                              (entry['dest_path'],now(),entry['id']))
                    self._audit(c,entry['id'],'applied',{'operation_id':operation['id'],**entry})
                c.execute("UPDATE supplement_operations SET state='committed' WHERE id=?",(operation['id'],))
            shutil.rmtree(stage,ignore_errors=True)
            return {'success':True,'count':len(applied),'applied':applied,'rejected':[],'message':f'Đã áp dụng {len(applied)} hồ sơ; giữ nguyên ảnh gốc.'}
        except Exception as exc:
            logger_message = str(exc)
            try:
                self._recover_operations()
            except Exception as recovery_error:
                logger_message += '; phục hồi chưa hoàn tất: '+str(recovery_error)
            with self._connect() as c:
                self._audit(c,'','apply_failed',{'operation_id':operation['id'],'error':logger_message})
            shutil.rmtree(stage,ignore_errors=True)
            return {'success':False,'count':0,'rejected':[],'message':logger_message}

    @serialized
    def delete_staging_photos(self, photo_ids):
        ids=selected_ids(photo_ids)
        self._recover_operations()
        operation={'id':uuid.uuid4().hex,'kind':'delete','files':[]}
        with self._connect() as c:
            c.execute('BEGIN IMMEDIATE')
            rows=c.execute('SELECT * FROM staging_photos WHERE id IN ('+','.join('?' for _ in ids)+')',ids).fetchall()
            if len(rows)!=len(ids) or any(r['applied_path'] for r in rows):
                raise RuntimeError('Có hồ sơ không tồn tại hoặc đã áp dụng; chưa xóa ảnh nào')
            for row in rows:
                item=dict(row)
                for source in (self.staging_dir/item['file_name'],self._source_path(item)):
                    if not source.is_file() or source.is_symlink() or not source.resolve().is_relative_to(self.data_dir.resolve()):
                        raise RuntimeError('Ảnh cần xóa bị mất hoặc đường dẫn không hợp lệ; chưa xóa hồ sơ')
                    stat=source.stat()
                    operation['files'].append({'source':str(source),'trash':str(self.temp_dir/(operation['id']+'_'+source.name)),
                                               'device':stat.st_dev,'inode':stat.st_ino})
            c.execute('INSERT INTO supplement_operations VALUES(?,?,?,?)',(operation['id'],'prepared',json.dumps(operation),now()))
        try:
            with self._connect() as c:
                c.execute('BEGIN IMMEDIATE')
                state=c.execute('SELECT state FROM supplement_operations WHERE id=?',(operation['id'],)).fetchone()[0]
                current=c.execute('SELECT * FROM staging_photos WHERE id IN ('+','.join('?' for _ in ids)+')',ids).fetchall()
                if state!='prepared' or len(current)!=len(ids) or any(r['applied_path'] for r in current):
                    raise RuntimeError('Hồ sơ thay đổi trong lúc xóa; chưa xóa ảnh nào')
                for entry in operation['files']:
                    stat=Path(entry['source']).stat()
                    if (stat.st_dev,stat.st_ino)!=(entry['device'],entry['inode']):
                        raise RuntimeError('File nguồn thay đổi trước khi xóa')
                    os.replace(entry['source'],entry['trash'])
                for row in current:
                    self._audit(c,row['id'],'deleted',dict(row))
                    c.execute('DELETE FROM staging_photos WHERE id=?',(row['id'],))
                c.execute("UPDATE supplement_operations SET state='committed' WHERE id=?",(operation['id'],))
        except Exception:
            self._recover_operations()
            raise RuntimeError('Không thể xóa; đã phục hồi hồ sơ, hãy thử lại')
        for entry in operation['files']:
            try:
                Path(entry['trash']).unlink(missing_ok=True)
            except OSError as exc:
                with self._connect() as c:
                    self._audit(c,'','delete_cleanup_pending',{'operation_id':operation['id'],'error':str(exc)})
        return len(ids)
