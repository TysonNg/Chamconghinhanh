# -*- coding: utf-8 -*-
"""
Photo Supplement Manager - Quản lý quy trình bổ sung ảnh chấm công
Pipeline: Upload ảnh chụp bù → Xóa watermark cũ → Tạo watermark mới → Sửa EXIF
"""

import os
import uuid
import json
import shutil
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from PIL import Image

from src.watermark_engine import (
    WatermarkScanner, WatermarkRemover, WatermarkGenerator, ExifEditor
)
from src.photo_edit_core import EditRequestNormalizer, ImageSourceService, TimestampAnalyzer

logger = logging.getLogger(__name__)


class SupplementRecord:
    """Một bản ghi bổ sung ảnh chấm công"""
    
    def __init__(self, record_id: str = None):
        self.id = record_id or str(uuid.uuid4())[:8]
        self.employee_name = ""
        self.target_date = ""       # YYYY-MM-DD
        self.target_time = ""       # HH:MM:SS
        self.original_path = ""     # Ảnh chụp bù gốc
        self.processed_path = ""    # Ảnh đã xử lý
        self.status = "pending"     # pending, processing, done, error
        self.error_message = ""
        self.created_at = datetime.now().isoformat()
        self.processed_at = ""
        
        # Tùy chọn watermark
        self.watermark_style = "timestamp_camera"
        self.watermark_position = "bottom-left"
        self.watermark_color = (255, 165, 0)
        self.location_name = ""
        self.gps_coords = ""
        self.remove_old_watermark = True
        self.modify_exif = True
    
    def to_dict(self) -> Dict:
        return {
            'id': self.id,
            'employee_name': self.employee_name,
            'target_date': self.target_date,
            'target_time': self.target_time,
            'original_path': self.original_path,
            'processed_path': self.processed_path,
            'status': self.status,
            'error_message': self.error_message,
            'created_at': self.created_at,
            'processed_at': self.processed_at,
            'watermark_style': self.watermark_style,
            'watermark_position': self.watermark_position,
            'watermark_color': list(self.watermark_color),
            'location_name': self.location_name,
            'gps_coords': self.gps_coords,
            'remove_old_watermark': self.remove_old_watermark,
            'modify_exif': self.modify_exif,
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'SupplementRecord':
        record = cls(data.get('id'))
        record.employee_name = data.get('employee_name', '')
        record.target_date = data.get('target_date', '')
        record.target_time = data.get('target_time', '')
        record.original_path = data.get('original_path', '')
        record.processed_path = data.get('processed_path', '')
        record.status = data.get('status', 'pending')
        record.error_message = data.get('error_message', '')
        record.created_at = data.get('created_at', '')
        record.processed_at = data.get('processed_at', '')
        record.watermark_style = data.get('watermark_style', 'timestamp_camera')
        record.watermark_position = data.get('watermark_position', 'bottom-left')
        color = data.get('watermark_color', [255, 165, 0])
        record.watermark_color = tuple(color) if isinstance(color, list) else color
        record.location_name = data.get('location_name', '')
        record.gps_coords = data.get('gps_coords', '')
        record.remove_old_watermark = data.get('remove_old_watermark', True)
        record.modify_exif = data.get('modify_exif', True)
        return record


class PhotoSupplement:
    """
    Quản lý toàn bộ quy trình bổ sung ảnh chấm công.
    
    Workflow:
    1. Upload ảnh chụp bù
    2. Xóa watermark cũ (nếu có)
    3. Tạo watermark mới với ngày giờ cần bổ sung
    4. Sửa EXIF metadata
    5. Xuất ảnh hoàn chỉnh
    """
    
    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.upload_dir = os.path.join(base_dir, "supplement_uploads")
        self.output_dir = os.path.join(base_dir, "supplement_output")
        self.temp_dir = os.path.join(base_dir, "supplement_temp")
        self.data_file = os.path.join(base_dir, "supplement_records.json")
        
        # Tạo thư mục
        for d in [self.upload_dir, self.output_dir, self.temp_dir]:
            os.makedirs(d, exist_ok=True)
        
        # Components
        self.scanner = WatermarkScanner()
        self.remover = WatermarkRemover()
        self.generator = WatermarkGenerator()
        self.exif_editor = ExifEditor()
        
        # Load records
        self.records: List[SupplementRecord] = []
        self._load_records()
    
    def _load_records(self):
        """Load danh sách records từ file"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.records = [SupplementRecord.from_dict(r) for r in data]
            except Exception as e:
                logger.error(f"Lỗi load records: {e}")
                self.records = []
    
    def _save_records(self):
        """Lưu danh sách records ra file"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump([r.to_dict() for r in self.records], f, 
                         ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Lỗi save records: {e}")
    
    def upload_photo(self, file_path: str, employee_name: str,
                     target_date: str, target_time: str,
                     **kwargs) -> SupplementRecord:
        """
        Bước 1: Upload ảnh chụp bù và tạo record.
        
        Args:
            file_path: Đường dẫn file ảnh đã upload
            employee_name: Tên nhân viên
            target_date: Ngày cần bổ sung (YYYY-MM-DD)
            target_time: Giờ cần bổ sung (HH:MM:SS)
            **kwargs: Tùy chọn bổ sung (watermark_style, location_name, etc.)
        """
        record = SupplementRecord()
        record.employee_name = employee_name
        record.target_date = target_date
        record.target_time = target_time
        
        # Copy ảnh vào upload directory
        ext = os.path.splitext(file_path)[1].lower() or '.jpg'
        upload_filename = f"{record.id}_{employee_name}{ext}"
        upload_path = os.path.join(self.upload_dir, upload_filename)
        
        shutil.copy2(file_path, upload_path)
        record.original_path = upload_path
        
        # Áp dụng tùy chọn
        record.watermark_style = kwargs.get('watermark_style', 'timestamp_camera')
        record.watermark_position = kwargs.get('watermark_position', 'bottom-left')
        color = kwargs.get('watermark_color')
        if color and isinstance(color, (list, tuple)) and len(color) == 3:
            record.watermark_color = tuple(color)
        record.location_name = kwargs.get('location_name', '')
        record.gps_coords = kwargs.get('gps_coords', '')
        record.remove_old_watermark = kwargs.get('remove_old_watermark', True)
        record.modify_exif = kwargs.get('modify_exif', True)
        
        self.records.append(record)
        self._save_records()
        
        logger.info(f"Uploaded supplement photo: {record.id} for {employee_name} on {target_date}")
        return record
    
    def process_record(self, record_id: str) -> SupplementRecord:
        """
        Xử lý một record: xóa watermark cũ → tạo watermark mới → sửa EXIF.
        """
        record = self.get_record(record_id)
        if not record:
            raise ValueError(f"Không tìm thấy record: {record_id}")
        
        if not os.path.exists(record.original_path):
            record.status = 'error'
            record.error_message = 'File ảnh gốc không tồn tại'
            self._save_records()
            return record
        
        record.status = 'processing'
        self._save_records()
        
        try:
            current_path = record.original_path
            ext = os.path.splitext(current_path)[1].lower() or '.jpg'
            
            # Bước 1: Xóa watermark cũ (nếu cần)
            if record.remove_old_watermark:
                temp_no_watermark = os.path.join(
                    self.temp_dir, f"{record.id}_no_watermark{ext}"
                )
                success = self.remover.remove_watermark_smart(
                    current_path, temp_no_watermark
                )
                if success:
                    current_path = temp_no_watermark
                    logger.info(f"[{record.id}] Đã xóa watermark cũ")
                else:
                    logger.warning(f"[{record.id}] Không xóa được watermark, tiếp tục...")
            
            # Bước 2: Tạo watermark mới
            output_filename = f"{record.employee_name}_{record.target_date}_{record.target_time.replace(':', '-')}{ext}"
            output_path = os.path.join(self.output_dir, output_filename)
            
            success = self.generator.generate_gps_camera_style(
                current_path, output_path,
                target_date=record.target_date,
                target_time=record.target_time,
                location_name=record.location_name,
                gps_coords=record.gps_coords,
                app_style=record.watermark_style
            )
            
            if not success:
                raise Exception("Lỗi tạo watermark mới")
            
            logger.info(f"[{record.id}] Đã tạo watermark mới")
            
            # Bước 3: Sửa EXIF metadata (nếu cần)
            if record.modify_exif:
                temp_exif = os.path.join(self.temp_dir, f"{record.id}_exif{ext}")
                self.exif_editor.modify_exif_date(
                    output_path, temp_exif,
                    target_date=record.target_date,
                    target_time=record.target_time
                )
                # Thay thế output bằng file đã sửa EXIF
                if os.path.exists(temp_exif):
                    shutil.move(temp_exif, output_path)
                
                logger.info(f"[{record.id}] Đã sửa EXIF metadata")
            
            record.processed_path = output_path
            record.status = 'done'
            record.processed_at = datetime.now().isoformat()
            
            # Dọn temp files
            self._cleanup_temp(record.id)
            
        except Exception as e:
            record.status = 'error'
            record.error_message = str(e)
            logger.error(f"[{record.id}] Lỗi xử lý: {e}")
        
        self._save_records()
        return record
    
    def batch_process(self, record_ids: Optional[List[str]] = None) -> List[SupplementRecord]:
        """Xử lý hàng loạt nhiều records"""
        if record_ids is None:
            # Xử lý tất cả pending records
            targets = [r for r in self.records if r.status == 'pending']
        else:
            targets = [r for r in self.records if r.id in record_ids]
        
        results = []
        for record in targets:
            result = self.process_record(record.id)
            results.append(result)
        
        return results
    
    def get_record(self, record_id: str) -> Optional[SupplementRecord]:
        """Lấy record theo ID"""
        for r in self.records:
            if r.id == record_id:
                return r
        return None
    
    def get_all_records(self) -> List[Dict]:
        """Lấy danh sách tất cả records"""
        return [r.to_dict() for r in self.records]
    
    def get_records_by_status(self, status: str) -> List[Dict]:
        """Lấy records theo status"""
        return [r.to_dict() for r in self.records if r.status == status]
    
    def delete_record(self, record_id: str) -> bool:
        """Xóa record và các file liên quan"""
        record = self.get_record(record_id)
        if not record:
            return False
        
        # Xóa files
        for path in [record.original_path, record.processed_path]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass
        
        self.records = [r for r in self.records if r.id != record_id]
        self._save_records()
        return True
    
    def validate_result(self, record_id: str) -> Dict:
        """
        Kiểm tra chất lượng ảnh đã xử lý.
        Trả về thông tin validation.
        """
        record = self.get_record(record_id)
        if not record or not record.processed_path:
            return {'valid': False, 'error': 'Record không tồn tại hoặc chưa xử lý'}
        
        if not os.path.exists(record.processed_path):
            return {'valid': False, 'error': 'File ảnh đã xử lý không tồn tại'}
        
        result = {
            'valid': True,
            'checks': {},
            'warnings': []
        }
        
        try:
            # Kiểm tra file size
            file_size = os.path.getsize(record.processed_path)
            result['checks']['file_size'] = {
                'value': file_size,
                'ok': file_size > 100000  # > 100KB
            }
            if file_size < 100000:
                result['warnings'].append('File quá nhỏ, có thể bị mất chất lượng')
            
            # Kiểm tra kích thước ảnh
            img = Image.open(record.processed_path)
            w, h = img.size
            result['checks']['dimensions'] = {
                'width': w, 'height': h,
                'ok': w >= 640 and h >= 480
            }
            if w < 640 or h < 480:
                result['warnings'].append('Ảnh có độ phân giải thấp')
            
            # Kiểm tra EXIF
            exif = ExifEditor.read_exif(record.processed_path)
            has_date_exif = any(k in exif for k in ['DateTime', 'DateTimeOriginal', 'DateTimeDigitized'])
            result['checks']['exif'] = {
                'has_date': has_date_exif,
                'ok': has_date_exif or not record.modify_exif
            }
            if record.modify_exif and not has_date_exif:
                result['warnings'].append('EXIF không chứa ngày giờ')
            
            # Kiểm tra watermark có visible không
            scanner = WatermarkScanner()
            watermark_analysis = scanner.scan_image(record.processed_path)
            result['checks']['watermark'] = {
                'detected': watermark_analysis is not None,
                'ok': watermark_analysis is not None
            }
            if not watermark_analysis:
                result['warnings'].append('Không phát hiện được watermark trên ảnh')
            
            result['valid'] = all(c['ok'] for c in result['checks'].values())
            
        except Exception as e:
            result['valid'] = False
            result['error'] = str(e)
        
        return result
    
    def get_preview(self, record_id: str) -> Optional[str]:
        """Lấy đường dẫn ảnh preview (ảnh đã xử lý hoặc ảnh gốc)"""
        record = self.get_record(record_id)
        if not record:
            return None
        
        if record.processed_path and os.path.exists(record.processed_path):
            return record.processed_path
        elif record.original_path and os.path.exists(record.original_path):
            return record.original_path
        return None
    
    def _cleanup_temp(self, record_id: str):
        """Dọn file tạm của record"""
        for f in os.listdir(self.temp_dir):
            if f.startswith(record_id):
                try:
                    os.remove(os.path.join(self.temp_dir, f))
                except Exception:
                    pass
