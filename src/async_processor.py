# -*- coding: utf-8 -*-
"""
Module xử lý bất đồng bộ
Xử lý nhiều ảnh song song với progress tracking
"""

import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

from src.config import MAX_WORKERS, SUPPORTED_IMAGE_EXTENSIONS, RESULTS_DIR
from src.face_detector import get_face_encoding, find_best_match, get_all_face_encodings
from src.text_extractor import extract_datetime_and_location, extract_datetime_simple
from src.database_manager import get_database_manager


class ProcessingTask:
    """Đại diện cho một task xử lý"""
    
    def __init__(self, task_id):
        self.task_id = task_id
        self.status = 'pending'  # pending, running, completed, failed
        self.progress = 0
        self.total = 0
        self.current_file = ''
        self.results = []
        self.errors = []
        self.start_time = None
        self.end_time = None
        self.output_file = None
    
    def to_dict(self):
        return {
            'task_id': self.task_id,
            'status': self.status,
            'progress': self.progress,
            'total': self.total,
            'current_file': self.current_file,
            'results_count': len(self.results),
            'errors_count': len(self.errors),
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
            'output_file': self.output_file,
            'elapsed_seconds': (datetime.now() - self.start_time).total_seconds() if self.start_time else 0
        }


class AsyncProcessor:
    """Xử lý ảnh bất đồng bộ"""
    
    def __init__(self):
        self.tasks = {}  # {task_id: ProcessingTask}
        self.executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
        self.db_manager = get_database_manager()
    
    def _get_image_files(self, folder_path):
        """Lấy danh sách file ảnh trong thư mục"""
        image_files = []
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_IMAGE_EXTENSIONS:
                    image_files.append(os.path.join(root, file))
        
        return image_files
    
    def _process_single_image(self, image_path):
        """
        Xử lý một ảnh
        
        Returns:
            dict: Kết quả xử lý
        """
        result = {
            'image_path': image_path,
            'filename': os.path.basename(image_path),
            'datetime': None,
            'location': None,
            'faces': [],
            'matched_person': None,
            'branch': None,
            'person_name': None,
            'confidence': None,
            'error': None
        }
        
        try:
            # 1. Trích xuất ngày tháng và địa điểm
            text_data = extract_datetime_simple(image_path)
            result['datetime'] = text_data.get('datetime')
            result['location'] = text_data.get('location')
            
            # 2. Phát hiện và nhận diện khuôn mặt
            faces_data = get_all_face_encodings(image_path)
            
            if faces_data:
                result['faces'] = [{'location': loc} for loc, enc in faces_data]
                
                # Tìm người phù hợp nhất cho khuôn mặt đầu tiên (lớn nhất)
                if faces_data:
                    _, first_encoding = faces_data[0]
                    known_faces = self.db_manager.get_all_faces()
                    
                    match = find_best_match(first_encoding, known_faces)
                    
                    if match:
                        result['matched_person'] = match['person_id']
                        result['branch'] = match['branch']
                        result['person_name'] = match['name']
                        result['confidence'] = match['confidence']
            
        except Exception as e:
            result['error'] = str(e)
        
        return result
    
    def start_processing(self, folder_path=None, image_files=None):
        """
        Bắt đầu xử lý ảnh
        
        Args:
            folder_path: Đường dẫn thư mục chứa ảnh
            image_files: Hoặc danh sách file ảnh cụ thể
        
        Returns:
            str: Task ID
        """
        # Tạo task mới
        task_id = f"task_{int(time.time() * 1000)}"
        task = ProcessingTask(task_id)
        self.tasks[task_id] = task
        
        # Lấy danh sách ảnh
        if image_files:
            files = image_files
        elif folder_path:
            files = self._get_image_files(folder_path)
        else:
            task.status = 'failed'
            task.errors.append('Không có ảnh để xử lý')
            return task_id
        
        if not files:
            task.status = 'failed'
            task.errors.append('Không tìm thấy file ảnh nào')
            return task_id
        
        task.total = len(files)
        task.status = 'running'
        task.start_time = datetime.now()
        
        # Chạy trong thread riêng
        thread = threading.Thread(target=self._run_processing, args=(task_id, files))
        thread.daemon = True
        thread.start()
        
        return task_id
    
    def _run_processing(self, task_id, files):
        """Xử lý trong background thread"""
        task = self.tasks[task_id]
        
        try:
            # Submit tất cả job lên executor
            futures = {}
            for file_path in files:
                future = self.executor.submit(self._process_single_image, file_path)
                futures[future] = file_path
            
            # Thu thập kết quả
            for future in as_completed(futures):
                file_path = futures[future]
                task.current_file = os.path.basename(file_path)
                
                try:
                    result = future.result(timeout=60)  # Timeout 60s per image
                    task.results.append(result)
                except Exception as e:
                    task.errors.append({
                        'file': file_path,
                        'error': str(e)
                    })
                
                task.progress += 1
            
            # Xuất kết quả ra Excel
            task.output_file = self._export_results(task)
            task.status = 'completed'
            
        except Exception as e:
            task.status = 'failed'
            task.errors.append(str(e))
        
        task.end_time = datetime.now()
    
    def _export_results(self, task):
        """Xuất kết quả ra file Excel"""
        try:
            import pandas as pd
            
            # Chuẩn bị data
            rows = []
            for i, result in enumerate(task.results, 1):
                rows.append({
                    'STT': i,
                    'Tên File': result['filename'],
                    'Ngày Giờ': result['datetime'] or '',
                    'Địa Điểm': result['location'] or '',
                    'Chi Nhánh': result['branch'] or '',
                    'Tên Người': result['person_name'] or 'Không xác định',
                    'Độ Tin Cậy (%)': result['confidence'] or 0,
                    'Số Khuôn Mặt': len(result['faces']),
                    'Lỗi': result['error'] or ''
                })
            
            df = pd.DataFrame(rows)
            
            # Tạo tên file output
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f'result_{timestamp}.xlsx'
            output_path = os.path.join(RESULTS_DIR, output_filename)
            
            # Xuất Excel
            df.to_excel(output_path, index=False, engine='openpyxl')
            
            return output_path
            
        except Exception as e:
            print(f"Lỗi xuất Excel: {e}")
            return None
    
    def get_task_status(self, task_id):
        """Lấy trạng thái task"""
        task = self.tasks.get(task_id)
        if task:
            return task.to_dict()
        return None
    
    def get_task_results(self, task_id):
        """Lấy kết quả chi tiết của task"""
        task = self.tasks.get(task_id)
        if task:
            return {
                **task.to_dict(),
                'results': task.results,
                'errors': task.errors
            }
        return None
    
    def get_all_tasks(self):
        """Lấy danh sách tất cả tasks"""
        return [task.to_dict() for task in self.tasks.values()]


# Singleton instance
_processor = None

def get_processor():
    global _processor
    if _processor is None:
        _processor = AsyncProcessor()
    return _processor
