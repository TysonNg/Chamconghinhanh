# -*- coding: utf-8 -*-
"""
Fake Photo Service - Tạo ảnh bổ sung chấm công với Watermark ngày giờ & EXIF
"""

import os
import re
import io
import uuid
import shutil
import sqlite3
import logging
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from PIL import Image, ImageOps
import piexif

from src.watermark_engine import ExifEditor
from src.smart_watermark_replacer import SmartWatermarkReplacer
from src.ai_timestamp_service import AITimestampService
from src.supplement_evidence import SupplementEvidenceMixin, parse_target_time, capture_datetime

logger = logging.getLogger(__name__)


class FakePhotoService(SupplementEvidenceMixin):
    def __init__(self, data_dir: str, registry_provider=None, matcher_provider=None):
        self.registry_provider = registry_provider
        self.matcher_provider = matcher_provider
        self.data_dir = Path(data_dir)
        self.staging_dir = self.data_dir / 'supplement_staging'
        self.raw_dir = self.data_dir / 'supplement_staging_raw'
        self.temp_dir = self.data_dir / 'supplement_temp'
        self.db_path = self.data_dir / 'supplement_batches.sqlite3'
        
        self.staging_dir.mkdir(parents=True, exist_ok=True)
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        
        # Thư mục input_images của hệ thống chấm công
        # data_dir thường là BASE_DIR/supplement_data
        base_candidate = self.data_dir.parent
        if (base_candidate / 'input_images').exists() or not (self.data_dir / 'input_images').exists():
            self.input_images_dir = base_candidate / 'input_images'
        else:
            self.input_images_dir = self.data_dir / 'input_images'
        self.input_images_dir.mkdir(parents=True, exist_ok=True)
        
        self.ai_service = AITimestampService(str(self.data_dir))
        self.smart_replacer = SmartWatermarkReplacer(ai_service=self.ai_service)
        self.exif_editor = ExifEditor()
        
        self._backup_evidence()
        self._init_db()
        self._init_evidence()

    def _connect(self):
        c = sqlite3.connect(self.db_path, timeout=30)
        c.row_factory = sqlite3.Row
        return c

    def _init_db(self):
        with self._connect() as c:
            c.execute('''
                CREATE TABLE IF NOT EXISTS staging_photos (
                    id TEXT PRIMARY KEY,
                    project TEXT NOT NULL,
                    employee TEXT NOT NULL,
                    target_date TEXT NOT NULL,
                    target_time TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    file_name TEXT NOT NULL,
                    created_at TEXT NOT NULL
                )
            ''')
            columns = {r[1] for r in c.execute('PRAGMA table_info(staging_photos)')}
            if 'inspection_json' not in columns:
                c.execute("ALTER TABLE staging_photos ADD COLUMN inspection_json TEXT NOT NULL DEFAULT '{}'")
            migrations = {
                'applied_path': "TEXT NOT NULL DEFAULT ''",
                'applied_at': "TEXT NOT NULL DEFAULT ''",
                'original_file_name': "TEXT NOT NULL DEFAULT ''",
                'watermark_status': "TEXT NOT NULL DEFAULT 'legacy'",
                'watermark_ocr_json': "TEXT NOT NULL DEFAULT '{}'",
                'confirmed_watermark_json': "TEXT NOT NULL DEFAULT '{}'",
                'generation_meta_json': "TEXT NOT NULL DEFAULT '{}'",
            }
            for column, declaration in migrations.items():
                if column not in columns:
                    c.execute(f"ALTER TABLE staging_photos ADD COLUMN {column} {declaration}")

    @staticmethod
    def _original_extension(original_name: str, raw_bytes: bytes) -> str:
        suffix = Path(original_name or '').suffix.lower()
        if suffix in ('.jpg', '.jpeg', '.png', '.webp', '.bmp'):
            return suffix
        try:
            with Image.open(io.BytesIO(raw_bytes)) as image:
                return '.' + (image.format or 'jpg').lower().replace('jpeg', 'jpg')
        except Exception:
            return '.jpg'

    @staticmethod
    def _target_watermark_lines(analysis: Dict, target_date: str, target_time: str) -> List[str]:
        source = list(analysis.get('confirmed_lines') or analysis.get('suggested_lines') or [])
        if not source:
            return []
        new_timestamp = str(analysis.get('new_timestamp') or '').strip()
        if not new_timestamp:
            dt = datetime.strptime(target_date, '%Y-%m-%d')
            new_timestamp = f"{dt.day:02d} Th{dt.month}, {dt.year} {target_time}"
        return [new_timestamp, *source[1:]]

    def create_supplement_photo(self, raw_bytes, project, employee, target_date, original_name,
                                target_time=None, *, project_id="", employee_id="",
                                replace_timestamp=True, modify_exif=True):
        """Keep observed evidence separate from requested changes to a derived image."""
        target_time = parse_target_time(target_time)
        if project_id or employee_id:
            p, e = self._identity(project_id, employee_id, datetime.strptime(target_date, '%Y-%m-%d').date())
            project, employee = p['storage_dir'], e['display_name']
        item = self.create_fake_photo(raw_bytes, project, employee, target_date, target_time, original_name,
                    replace_timestamp=bool(target_time and replace_timestamp),
                    modify_exif=bool(target_time and modify_exif), project_id=project_id, employee_id=employee_id)
        inspection = dict(item['inspection'])
        try:
            observed = self.ai_service.inspect_photo(raw_bytes)
            if isinstance(observed, dict):
                inspection.update(observed)
                inspection['original_visible_date'] = observed.get('visible_date')
                inspection['visible_date'] = observed.get('visible_date')
                inspection['visible_time'] = observed.get('visible_time')
        except Exception as exc:
            inspection.update(status='failed', message='Không đọc được ngày trên ảnh: ' + type(exc).__name__)
        inspection['exif_datetime'] = capture_datetime(raw_bytes)
        inspection['date_mismatch'] = bool(inspection['exif_datetime'] and not inspection['exif_datetime'].startswith(target_date))
        with self._connect() as c:
            c.execute('UPDATE staging_photos SET inspection_json=? WHERE id=?',
                      (json.dumps(inspection, ensure_ascii=False), item['id']))
        return self.inspect_record(item['id'], check_face=True)

    def create_fake_photo(
        self,
        raw_bytes: bytes,
        project: str,
        employee: str,
        target_date: str,
        target_time: str,
        original_name: str,
        replace_timestamp: bool = True,
        modify_exif: bool = True,
        location_name: str = "",
        # Legacy params (ignored, kept for API compat)
        remove_old_watermark: bool = True,
        add_watermark: bool = True,
        project_id: str = "",
        employee_id: str = "",
    ) -> Dict:
        """
        Xử lý 1 ảnh: Phát hiện text timestamp cũ → Xóa chỉ text → Vẽ text mới → Sửa EXIF.
        Giữ nguyên 100% ảnh gốc (mặt, nền), chỉ thay đổi text ngày/giờ.
        """
        target_time = parse_target_time(target_time)
        datetime.strptime(target_date, '%Y-%m-%d')
        replace_timestamp = bool(replace_timestamp and target_time)
        modify_exif = bool(modify_exif and target_time)
        photo_id = uuid.uuid4().hex[:12]
        temp_in = self.temp_dir / f"{photo_id}_raw.jpg"
        temp_replaced = self.temp_dir / f"{photo_id}_replaced.jpg"
        final_file = self.staging_dir / f"{photo_id}.jpg"
        original_file_name = f"{photo_id}_original{self._original_extension(original_name, raw_bytes)}"
        original_file = self.raw_dir / original_file_name

        try:
            # 0. Lưu bytes upload bất biến. Mọi lần gen lại chỉ được đọc file này.
            original_file.write_bytes(raw_bytes)

            # 1. Đọc ảnh và chuẩn hóa sang RGB JPEG
            with Image.open(io.BytesIO(raw_bytes)) as img:
                img = ImageOps.exif_transpose(img)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(temp_in, format='JPEG', quality=95)

            current_path = str(temp_in)
            watermark_status = 'not_requested'
            watermark_ocr = {}
            confirmed_watermark = {}
            generation_meta = {}
            exif_status = 'not_requested'

            # 2. OCR toàn bộ watermark (Cloud AI hoặc EasyOCR nội bộ) và xử lý watermark.
            if replace_timestamp:
                try:
                    watermark_ocr = self.ai_service.analyze_watermark_block(
                        raw_bytes, target_date, target_time
                    )
                    watermark_status = watermark_ocr.get('status', 'needs_confirmation')
                    has_lines = bool(watermark_ocr.get('confirmed_lines') or watermark_ocr.get('suggested_lines'))
                    if not has_lines:
                        watermark_status = 'verification_failed'
                        generation_meta = {'error': 'full_watermark_ocr_required'}
                    elif watermark_status == 'confirmed' and self.ai_service.is_configured():
                        lines = self._target_watermark_lines(watermark_ocr, target_date, target_time)
                        block = dict(watermark_ocr)
                        block['confirmed_lines'] = lines
                        generated = self.ai_service.generate_verified_watermark_crop(raw_bytes, block)
                        generation_meta = {k: v for k, v in generated.items() if k != 'image_bytes'}
                        watermark_status = generated.get('status', 'verification_failed')
                        if watermark_status == 'completed' and generated.get('image_bytes'):
                            temp_replaced.write_bytes(generated['image_bytes'])
                            current_path = str(temp_replaced)
                            confirmed_watermark = {'lines': lines}

                    # Fallback: nếu AI block generation thất bại nhưng có thông tin watermark,
                    # thử SmartWatermarkReplacer (xóa nét mỏng + vẽ lại bằng PIL nội bộ)
                    if has_lines and watermark_status != 'completed':
                        try:
                            logger.info('AI block generation chưa thành công, thử fallback SmartWatermarkReplacer...')
                            fallback_ok = self.smart_replacer.replace_timestamp(
                                str(temp_in), str(temp_replaced), target_date, target_time
                            )
                            if fallback_ok and temp_replaced.exists():
                                current_path = str(temp_replaced)
                                watermark_status = 'completed'
                                generation_meta['fallback'] = 'smart_replacer'
                                logger.info('Fallback SmartWatermarkReplacer thành công!')
                            else:
                                watermark_status = 'verification_failed'
                        except Exception as fb_err:
                            watermark_status = 'verification_failed'
                            logger.warning(f'Fallback SmartWatermarkReplacer thất bại: {fb_err}')
                except Exception as e:
                    watermark_status = 'verification_failed'
                    generation_meta = {'error': type(e).__name__}
                    logger.warning(f"Lỗi tạo watermark block: {e}, giữ nguyên ảnh gốc")

            # 3. Sửa EXIF metadata nếu được bật
            if modify_exif and watermark_status in ('completed', 'not_requested'):
                try:
                    date_parts = target_date.split('-')
                    exif_dt = f"{date_parts[0]}:{date_parts[1]}:{date_parts[2]} {target_time}"
                    exif_bytes_val = exif_dt.encode('utf-8')

                    try:
                        exif_dict = piexif.load(current_path)
                    except Exception:
                        exif_dict = {'0th': {}, 'Exif': {}, 'GPS': {}, '1st': {}}

                    exif_dict['0th'][piexif.ImageIFD.DateTime] = exif_bytes_val
                    exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = exif_bytes_val
                    exif_dict['Exif'][piexif.ExifIFD.DateTimeDigitized] = exif_bytes_val

                    exif_raw = piexif.dump(exif_dict)
                    piexif.insert(exif_raw, current_path)
                    exif_status = 'completed'
                except Exception as e:
                    logger.warning(f"Lỗi chèn piexif: {e}, fallback regex...")
                    fallback_exif = self.exif_editor.modify_exif_date(current_path, current_path, target_date, target_time)
                    exif_status = 'completed' if fallback_exif else 'failed'


            # Copy kết quả cuối cùng vào Staging
            if watermark_status == 'completed' or (not replace_timestamp and exif_status == 'completed'):
                shutil.copy2(current_path, str(final_file))
            else:
                final_file.write_bytes(raw_bytes)

            source_sha256 = hashlib.sha256(raw_bytes).hexdigest()
            with self._connect() as c:
                duplicate_ids = [row[0] for row in c.execute('SELECT id FROM staging_photos WHERE source_sha256=?', (source_sha256,))]
            orig_exif_dt = capture_datetime(raw_bytes)
            inspection = {
                'status': 'not_run', 'visible_date': None, 'visible_time': None,
                'original_visible_date': None, 'exif_datetime': orig_exif_dt,
                'date_mismatch': bool(orig_exif_dt and not orig_exif_dt.startswith(target_date)),
                'message': 'Đã lưu ảnh gốc; kết quả xử lý được báo riêng', 'duplicate_ids': duplicate_ids
            }

            # Lưu vào Database
            record = {
                'id': photo_id,
                'project': project,
                'employee': employee,
                'target_date': target_date,
                'target_time': target_time,
                'original_name': original_name,
                'file_name': f"{photo_id}.jpg",
                'created_at': datetime.now(timezone.utc).isoformat(),
                'inspection': inspection,
                'original_file_name': original_file_name,
                'watermark_status': watermark_status,
                'watermark_ocr': watermark_ocr,
                'confirmed_watermark': confirmed_watermark,
                'generation_meta': generation_meta,
                'project_id': project_id, 'employee_id': employee_id,
                'source_sha256': source_sha256, 'derived_sha256': hashlib.sha256(final_file.read_bytes()).hexdigest(),
                'exif_status': exif_status,
            }

            with self._connect() as c:
                c.execute('''
                    INSERT INTO staging_photos (
                        id, project, employee, target_date, target_time, original_name,
                        file_name, created_at, inspection_json, original_file_name,
                        watermark_status, watermark_ocr_json, confirmed_watermark_json,
                        generation_meta_json, project_id, employee_id, source_sha256, derived_sha256, exif_status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    record['id'], record['project'], record['employee'],
                    record['target_date'], record['target_time'],
                    record['original_name'], record['file_name'],
                    record['created_at'],
                    json.dumps(inspection, ensure_ascii=False),
                    original_file_name,
                    watermark_status,
                    json.dumps(watermark_ocr, ensure_ascii=False),
                    json.dumps(confirmed_watermark, ensure_ascii=False),
                    json.dumps(generation_meta, ensure_ascii=False),
                    project_id, employee_id, source_sha256, record['derived_sha256'], exif_status,
                ))
                self._audit(c, photo_id, 'created', {'source_sha256': source_sha256, 'project_id':project_id, 'employee_id':employee_id})

            record['url'] = f"/api/supplement/staging/{photo_id}/image"
            return self._public_record(record)

        finally:
            # Dọn dẹp file tạm
            for p in [temp_in, temp_replaced]:
                try:
                    if p.exists():
                        p.unlink(missing_ok=True)
                except Exception:
                    pass

    def get_staging_image_path(self, photo_id: str) -> Optional[Path]:
        """Lấy đường dẫn file ảnh trong staging"""
        with self._connect() as c:
            row = c.execute('SELECT file_name FROM staging_photos WHERE id=?', (photo_id,)).fetchone()
        if not row:
            return None
        p = self.staging_dir / row[0]
        return p if p.exists() else None

    def get_watermark_crop_bytes(self, photo_id: str) -> Optional[bytes]:
        with self._connect() as c:
            row = c.execute(
                'SELECT original_file_name, watermark_ocr_json FROM staging_photos WHERE id=?',
                (photo_id,),
            ).fetchone()
        if not row:
            return None
        original_name = row[0] or f'{photo_id}_raw.jpg'
        original_path = self.raw_dir / original_name
        if not original_path.exists():
            legacy_candidates = list(self.raw_dir.glob(f"{photo_id}_original.*")) + [self.temp_dir / f"{photo_id}_raw.jpg"]
            found = next((c for c in legacy_candidates if c.exists()), None)
            if found:
                original_path = found
            else:
                return None
        try:
            analysis = json.loads(row[1] or '{}')
            if not analysis.get('block_box_2d') and not analysis.get('box_2d'):
                analysis['block_box_2d'] = [700, 0, 1000, 1000]
            with Image.open(original_path) as image:
                image = ImageOps.exif_transpose(image).convert('RGB')
                crop, _ = self.ai_service.extract_watermark_crop(image, analysis)
                stream = io.BytesIO()
                crop.save(stream, 'JPEG', quality=95)
                return stream.getvalue()
        except Exception as exc:
            logger.warning('Không thể tạo crop watermark %s: %s', photo_id, exc)
            return None

    def regenerate_staging_photo(
        self, photo_id: str, confirmed_lines: Optional[List[str]] = None
    ) -> Dict:
        """Regenerate from the immutable upload only; never chain AI outputs."""
        with self._connect() as c:
            row = c.execute('SELECT * FROM staging_photos WHERE id=?', (photo_id,)).fetchone()
        if not row:
            logger.error(f"Không tìm thấy photo_id={photo_id} trong database")
            return {'status': 'not_found'}

        record = dict(row)
        target_date = record['target_date']
        target_time = record['target_time']
        original_name = record.get('original_file_name') or ''
        raw_file = self.raw_dir / original_name if original_name else self.raw_dir / f"{photo_id}_raw.jpg"
        staging_file = self.staging_dir / f"{photo_id}.jpg"
        if not raw_file.exists():
            logger.error(f"Không tìm thấy ảnh gốc bất biến cho photo_id={photo_id}")
            return {'status': 'original_missing', 'photo_id': photo_id}

        raw_bytes = raw_file.read_bytes()
        if not target_time:
            return {'status':'invalid_request', 'error':'Nhập giờ trước khi sửa watermark/EXIF'}
        if not record.get('source_sha256') or hashlib.sha256(raw_bytes).hexdigest() != record['source_sha256']:
            return {'status':'integrity_failed', 'error':'Ảnh gốc chưa được xác nhận toàn vẹn hoặc đã thay đổi'}
        temp_candidate = self.temp_dir / f"{photo_id}_candidate.jpg"

        try:
            try:
                analysis = json.loads(record.get('watermark_ocr_json') or '{}')
            except Exception:
                analysis = {}
            try:
                stored_confirmation = json.loads(record.get('confirmed_watermark_json') or '{}')
            except Exception:
                stored_confirmation = {}

            has_valid_analysis = bool(analysis.get('block_box_2d') and (analysis.get('confirmed_lines') or analysis.get('suggested_lines')))
            if not has_valid_analysis:
                analysis = self.ai_service.analyze_watermark_block(raw_bytes, target_date, target_time)

            lines = [
                self.ai_service.normalize_watermark_text(line)
                for line in (confirmed_lines or stored_confirmation.get('lines') or [])
                if self.ai_service.normalize_watermark_text(line)
            ]
            if not lines:
                lines = self._target_watermark_lines(analysis, target_date, target_time)

            if len(lines) < 1:
                meta = {'error': 'full_watermark_ocr_required'}
                with self._connect() as c:
                    c.execute(
                        '''UPDATE staging_photos SET watermark_status=?, watermark_ocr_json=?,
                           generation_meta_json=? WHERE id=?''',
                        (
                            'verification_failed', json.dumps(analysis, ensure_ascii=False),
                            json.dumps(meta, ensure_ascii=False), photo_id,
                        ),
                    )
                return {
                    'status': 'verification_failed', 'photo_id': photo_id,
                    'generation_meta': meta,
                }
            else:
                block = dict(analysis)
                block['confirmed_lines'] = lines
                generated = self.ai_service.generate_verified_watermark_crop(raw_bytes, block)

            # Fallback SmartWatermarkReplacer khi AI image generation thất bại (chỉ khi không có confirmed_lines thủ công)
            if not confirmed_lines and (generated.get('status') != 'completed' or not generated.get('image_bytes')):
                try:
                    logger.info('AI generation thất bại khi regenerate, thử fallback SmartWatermarkReplacer...')
                    temp_in = self.temp_dir / f"{photo_id}_raw.jpg"
                    with Image.open(io.BytesIO(raw_bytes)) as img:
                        img = ImageOps.exif_transpose(img)
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        img.save(temp_in, format='JPEG', quality=95)
                    fallback_out = self.temp_dir / f"{photo_id}_fallback.jpg"
                    fallback_ok = self.smart_replacer.replace_timestamp(
                        str(temp_in), str(fallback_out), target_date, target_time
                    )
                    if fallback_ok and fallback_out.exists():
                        generated = {
                            'status': 'completed',
                            'image_bytes': fallback_out.read_bytes(),
                            'fallback': 'smart_replacer',
                        }
                        logger.info('Fallback SmartWatermarkReplacer thành công khi regenerate!')
                    try:
                        temp_in.unlink(missing_ok=True)
                        fallback_out.unlink(missing_ok=True)
                    except Exception:
                        pass
                except Exception as fb_err:
                    logger.warning(f'Fallback SmartWatermarkReplacer thất bại khi regenerate: {fb_err}')

            status = generated.get('status', 'verification_failed')
            meta = {key: value for key, value in generated.items() if key != 'image_bytes'}
            if status != 'completed' or not generated.get('image_bytes'):
                with self._connect() as c:
                    c.execute(
                        '''UPDATE staging_photos SET watermark_status=?, watermark_ocr_json=?,
                           confirmed_watermark_json=?, generation_meta_json=? WHERE id=?''',
                        (
                            status, json.dumps(analysis, ensure_ascii=False),
                            json.dumps({'lines': lines}, ensure_ascii=False),
                            json.dumps(meta, ensure_ascii=False), photo_id,
                        ),
                    )
                return {'status': status, 'photo_id': photo_id, 'generation_meta': meta}

            temp_candidate.write_bytes(generated['image_bytes'])

            # Sửa EXIF
            exif_status = 'not_requested'
            try:
                date_parts = target_date.split('-')
                exif_dt = f"{date_parts[0]}:{date_parts[1]}:{date_parts[2]} {target_time}"
                exif_bytes_val = exif_dt.encode('utf-8')
                try:
                    exif_dict = piexif.load(str(temp_candidate))
                except Exception:
                    exif_dict = {'0th': {}, 'Exif': {}, 'GPS': {}, '1st': {}}
                exif_dict['0th'][piexif.ImageIFD.DateTime] = exif_bytes_val
                exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = exif_bytes_val
                exif_dict['Exif'][piexif.ExifIFD.DateTimeDigitized] = exif_bytes_val
                exif_raw = piexif.dump(exif_dict)
                piexif.insert(exif_raw, str(temp_candidate))
                exif_status = 'completed'
            except Exception as e:
                exif_status = 'failed'
                logger.warning(f"Lỗi chèn piexif khi regenerate: {e}")

            # Chỉ thay kết quả hiện tại sau khi candidate đã vượt hậu kiểm.
            os.replace(str(temp_candidate), str(staging_file))

            now_iso = datetime.now(timezone.utc).isoformat()
            with self._connect() as c:
                c.execute(
                    '''UPDATE staging_photos SET created_at=?, watermark_status=?,
                       watermark_ocr_json=?, confirmed_watermark_json=?, generation_meta_json=?
                       WHERE id=?''',
                    (
                        now_iso, 'completed', json.dumps(analysis, ensure_ascii=False),
                        json.dumps({'lines': lines}, ensure_ascii=False),
                        json.dumps(meta, ensure_ascii=False), photo_id,
                    ),
                )

            with self._connect() as c:
                c.execute('UPDATE staging_photos SET derived_sha256=?,exif_status=? WHERE id=?',
                          (hashlib.sha256(staging_file.read_bytes()).hexdigest(), exif_status, photo_id))
                self._audit(c, photo_id, 'regenerated', {'watermark':'completed','exif':exif_status})
            record['exif_status'] = exif_status
            record['created_at'] = now_iso
            record['url'] = f"/api/supplement/staging/{photo_id}/image"
            record['size_kb'] = round(staging_file.stat().st_size / 1024, 1)
            record['inspection'] = json.loads(record.get('inspection_json', '{}'))
            record['watermark_status'] = 'completed'
            record['watermark_ocr'] = analysis
            record['confirmed_watermark'] = {'lines': lines}
            record['generation_meta'] = meta
            record['status'] = 'completed'
            logger.info(f"Đã regenerate thành công photo_id={photo_id}")
            return self._public_record(record)

        finally:
            try:
                temp_candidate.unlink(missing_ok=True)
            except Exception:
                pass

