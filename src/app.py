# -*- coding: utf-8 -*-
"""
Flask Web Server cho phần mềm nhận diện khuôn mặt
Phiên bản đơn giản để test UI
"""

import os
import sys
import time
import json
import queue
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import shutil
import re
import logging
from typing import List, Dict, Optional, Tuple

# Thêm thư mục gốc vào path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, Response
from werkzeug.utils import secure_filename

# Cấu hình - Phát hiện đúng thư mục khi chạy từ EXE
def get_base_dir():
    """Lấy thư mục gốc (chứa data: input_images, database, ...) - hỗ trợ cả khi chạy từ source và từ EXE"""
    if getattr(sys, 'frozen', False):
        # Chạy từ EXE (PyInstaller) - data nằm cạnh file exe
        return os.path.dirname(sys.executable)
    else:
        # Chạy từ source
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_resource_dir():
    """Lấy thư mục resources (templates, static) - hỗ trợ cả khi chạy từ EXE"""
    if getattr(sys, 'frozen', False):
        # Chạy từ EXE - resources được PyInstaller giải nén vào _MEIPASS
        return sys._MEIPASS
    else:
        # Chạy từ source
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = get_base_dir()
RESOURCE_DIR = get_resource_dir()
INPUT_IMAGES_DIR = os.path.join(BASE_DIR, "input_images")
DATABASE_DIR = os.path.join(BASE_DIR, "database")
RESULTS_DIR = os.path.join(BASE_DIR, "results")

def _normalize_folder_name(name: str) -> str:
    import unicodedata
    name = unicodedata.normalize('NFD', name)
    name = ''.join(c for c in name if unicodedata.category(c) != 'Mn')
    return ' '.join(name.lower().strip().split())

def _count_images_in_dir(path: str) -> int:
    exts = {'.jpg', '.jpeg', '.png', '.bmp'}
    count = 0
    for root, _, files in os.walk(path):
        for f in files:
            if os.path.splitext(f)[1].lower() in exts:
                count += 1
    return count

def resolve_portrait_dir(base_dir: str) -> str:
    """Find portrait folder even if accents/encoding differ."""
    candidates = ["Ảnh BV", "Anh BV", "ANH BV", "anh bv", "ẢnhBV"]
    for c in candidates:
        p = os.path.join(base_dir, c)
        if os.path.exists(p):
            return p
    try:
        target = _normalize_folder_name("Ảnh BV")
        for name in os.listdir(base_dir):
            if _normalize_folder_name(name) == target:
                return os.path.join(base_dir, name)
    except Exception:
        pass
    return os.path.join(base_dir, "Ảnh BV")

def resolve_portrait_dir_by_scan(base_dir: str) -> str:
    """Fallback: scan subfolders and pick one that looks like portrait dir."""
    best_dir = None
    best_count = 0
    keywords = ("anh", "chan dung", "portrait", "bv")
    try:
        for name in os.listdir(base_dir):
            path = os.path.join(base_dir, name)
            if not os.path.isdir(path):
                continue
            norm = _normalize_folder_name(name)
            if not any(k in norm for k in keywords):
                continue
            count = _count_images_in_dir(path)
            if count > best_count:
                best_count = count
                best_dir = path
    except Exception:
        pass
    return best_dir or resolve_portrait_dir(base_dir)
SUPPORTED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

FLASK_HOST = "0.0.0.0"
FLASK_PORT = 5000
FLASK_DEBUG = True
MAX_WORKERS = 4

# Thu muc du lieu duoc tao tu dong khi clone moi
CHAMCONG_DIR = os.path.join(BASE_DIR, "chamcong")
PORTRAIT_DIR = resolve_portrait_dir(BASE_DIR)
NGAY_RONG_DIR = os.path.join(BASE_DIR, "ngay_rong")

# Tao thu muc neu chua ton tai
for directory in [INPUT_IMAGES_DIR, RESULTS_DIR, CHAMCONG_DIR, PORTRAIT_DIR, NGAY_RONG_DIR]:
    os.makedirs(directory, exist_ok=True)

DEFAULT_PROJECT_NAME = "Chung cư Tân Thuận Đông"

def ensure_initial_project_migration():
    """Tự động gom các nhân viên ban đầu vào dự án mặc định nếu chưa gom"""
    default_portrait_project_dir = os.path.join(PORTRAIT_DIR, DEFAULT_PROJECT_NAME)
    default_input_project_dir = os.path.join(INPUT_IMAGES_DIR, DEFAULT_PROJECT_NAME)
    os.makedirs(default_portrait_project_dir, exist_ok=True)
    os.makedirs(default_input_project_dir, exist_ok=True)
    
    if os.path.exists(PORTRAIT_DIR):
        for item in os.listdir(PORTRAIT_DIR):
            item_path = os.path.join(PORTRAIT_DIR, item)
            if not os.path.isdir(item_path) or item == DEFAULT_PROJECT_NAME:
                continue
            try:
                sub_items = os.listdir(item_path)
            except Exception:
                continue
            sub_dirs = [s for s in sub_items if os.path.isdir(os.path.join(item_path, s))]
            if not sub_dirs:
                dest_path = os.path.join(default_portrait_project_dir, item)
                try:
                    if not os.path.exists(dest_path):
                        shutil.move(item_path, dest_path)
                    else:
                        for f in sub_items:
                            src_f = os.path.join(item_path, f)
                            dst_f = os.path.join(dest_path, f)
                            if not os.path.exists(dst_f):
                                shutil.move(src_f, dst_f)
                        shutil.rmtree(item_path, ignore_errors=True)
                except Exception as e:
                    logging.error(f"[MIGRATION] Lỗi di chuyển {item}: {e}")

ensure_initial_project_migration()

def ensure_project_structure(project_name: str, create_day_folders: bool = True) -> Tuple[str, str]:
    """Đảm bảo đầy đủ cấu trúc thư mục cho dự án:
    1. Ảnh BV/<project_name> (Thư mục chân dung nhân viên)
    2. input_images/<project_name> (Thư mục ảnh camera theo ngày)
    3. input_images/<project_name>/01..31 (31 thư mục ngày trong tháng)
    """
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', project_name.strip()) if project_name else DEFAULT_PROJECT_NAME
    p_dir = os.path.join(PORTRAIT_DIR, safe_name)
    i_dir = os.path.join(INPUT_IMAGES_DIR, safe_name)
    os.makedirs(p_dir, exist_ok=True)
    os.makedirs(i_dir, exist_ok=True)
    if create_day_folders:
        for d in range(1, 32):
            os.makedirs(os.path.join(i_dir, f"{d:02d}"), exist_ok=True)
    return p_dir, i_dir

def get_all_projects() -> List[str]:
    """Lấy danh sách tên tất cả các dự án từ PORTRAIT_DIR và INPUT_IMAGES_DIR"""
    projects = set()
    if os.path.exists(PORTRAIT_DIR):
        for item in os.listdir(PORTRAIT_DIR):
            p_path = os.path.join(PORTRAIT_DIR, item)
            if os.path.isdir(p_path):
                projects.add(item)
    if os.path.exists(INPUT_IMAGES_DIR):
        for item in os.listdir(INPUT_IMAGES_DIR):
            p_path = os.path.join(INPUT_IMAGES_DIR, item)
            if os.path.isdir(p_path) and not (item.isdigit() and len(item) <= 2):
                projects.add(item)
    if not projects:
        projects.add(DEFAULT_PROJECT_NAME)
    # Tự động đồng bộ và tạo đầy đủ cấu trúc thư mục cho tất cả các dự án
    for p in projects:
        ensure_project_structure(p)
    return sorted(list(projects))

def resolve_project_dirs(project_name: str) -> Tuple[str, str]:
    """Trả về (portrait_dir, input_images_dir) an toàn cho dự án"""
    return ensure_project_structure(project_name)

# Flask app - dùng RESOURCE_DIR cho templates/static (đúng cả khi chạy từ EXE)
app = Flask(__name__, 
            template_folder=os.path.join(RESOURCE_DIR, 'templates'),
            static_folder=os.path.join(RESOURCE_DIR, 'static'))

app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max upload

# ==================== TASK MANAGER ====================

class ProcessingTask:
    def __init__(self, task_id):
        self.task_id = task_id
        self.status = 'pending'
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
            'output_file': self.output_file
        }

# Global state
tasks = {}
database = {}

# ==================== LOG STREAMING ====================
# Global log queue for SSE streaming
log_queues = []
log_lock = threading.Lock()

# Setup file logging
LOG_FILE = os.path.join(BASE_DIR, 'app.log')

def setup_file_logging():
    """Configure logging to file"""
    import logging
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8'
    )
    # Log startup
    logging.info("="*50)
    logging.info("APP STARTED")
    logging.info(f"Frozen: {getattr(sys, 'frozen', False)}")
    logging.info(f"Base Dir (data): {BASE_DIR}")
    logging.info(f"Resource Dir (templates/static): {RESOURCE_DIR}")
    logging.info(f"Input Images Dir: {INPUT_IMAGES_DIR}")
    logging.info(f"Template folder: {app.template_folder}")
    logging.info(f"Static folder: {app.static_folder}")
    logging.info(f"Template folder exists: {os.path.exists(app.template_folder)}")
    logging.info(f"Static folder exists: {os.path.exists(app.static_folder)}")
    logging.info("="*50)

# Initialize logging
setup_file_logging()

def send_log(message, log_type='default'):
    """Send log message to all connected SSE clients and write to file"""
    import logging
    
    # Write to file
    if log_type == 'error':
        logging.error(message)
    elif log_type == 'warning':
        logging.warning(message)
    elif log_type == 'success':
        logging.info(f"[SUCCESS] {message}")
    else:
        logging.info(message)
        
    # Also print to console
    print(message)
    
    # Send to SSE clients
    with log_lock:
        for q in log_queues:
            try:
                msg_data = json.dumps({
                    'message': message,
                    'type': log_type,
                    'time': datetime.now().strftime('%H:%M:%S')
                })
                q.put(msg_data)
            except Exception:
                pass


def get_image_files(folder_path):
    """Lấy danh sách file ảnh trong thư mục"""
    image_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in SUPPORTED_IMAGE_EXTENSIONS:
                image_files.append(os.path.join(root, file))
    return image_files

def process_image(image_path):
    """Xử lý một ảnh (demo version - trả về dữ liệu mẫu)"""
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
    
    # Simulate processing time
    time.sleep(0.5)
    
    # Demo data
    result['datetime'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    result['location'] = 'Demo Location'
    result['faces'] = [{'location': (0, 100, 100, 0)}]
    
    # Tìm trong database
    if database:
        for person_id, data in database.items():
            result['matched_person'] = person_id
            result['branch'] = data.get('branch', 'Unknown')
            result['person_name'] = data.get('name', 'Unknown')
            result['confidence'] = 85.5
            break
    
    return result

def run_processing(task_id, files):
    """Xử lý trong background thread"""
    task = tasks[task_id]
    task.status = 'running'
    task.start_time = datetime.now()
    task.total = len(files)
    
    try:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_image, f): f for f in files}
            
            for future in as_completed(futures):
                file_path = futures[future]
                task.current_file = os.path.basename(file_path)
                
                try:
                    result = future.result(timeout=60)
                    task.results.append(result)
                except Exception as e:
                    task.errors.append({'file': file_path, 'error': str(e)})
                
                task.progress += 1
        
        # Export to Excel
        task.output_file = export_results(task)
        task.status = 'completed'
        
    except Exception as e:
        task.status = 'failed'
        task.errors.append(str(e))
    
    task.end_time = datetime.now()

def export_results(task):
    """Xuất kết quả ra file Excel hoặc CSV"""
    try:
        import pandas as pd
        
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
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'result_{timestamp}.xlsx'
        output_path = os.path.join(RESULTS_DIR, output_filename)
        
        df.to_excel(output_path, index=False, engine='openpyxl')
        return output_path
        
    except ImportError:
        # Fallback to CSV if pandas not available
        import csv
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'result_{timestamp}.csv'
        output_path = os.path.join(RESULTS_DIR, output_filename)
        
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['STT', 'Tên File', 'Ngày Giờ', 'Địa Điểm', 'Chi Nhánh', 'Tên Người', 'Độ Tin Cậy (%)', 'Số Khuôn Mặt', 'Lỗi'])
            
            for i, result in enumerate(task.results, 1):
                writer.writerow([
                    i,
                    result['filename'],
                    result['datetime'] or '',
                    result['location'] or '',
                    result['branch'] or '',
                    result['person_name'] or 'Không xác định',
                    result['confidence'] or 0,
                    len(result['faces']),
                    result['error'] or ''
                ])
        
        return output_path
    except Exception as e:
        print(f"Lỗi xuất file: {e}")
        return None

def scan_database():
    """Quét database ảnh chân dung"""
    global database
    database = {}
    
    if not os.path.exists(DATABASE_DIR):
        return
    
    for branch in os.listdir(DATABASE_DIR):
        branch_path = os.path.join(DATABASE_DIR, branch)
        if not os.path.isdir(branch_path) or branch.startswith('.'):
            continue
        
        for person_name in os.listdir(branch_path):
            person_path = os.path.join(branch_path, person_name)
            if not os.path.isdir(person_path):
                continue
            
            # Tìm ảnh
            image_files = []
            for file in os.listdir(person_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_IMAGE_EXTENSIONS:
                    image_files.append(os.path.join(person_path, file))
            
            if image_files:
                person_id = f"{branch}/{person_name}"
                database[person_id] = {
                    'encoding': None,  # Placeholder
                    'branch': branch,
                    'name': person_name,
                    'image_path': image_files[0]
                }

# ==================== PAGES ====================

@app.route('/')
def index():
    return render_template('index.html')

# ==================== SSE LOG STREAM ====================

@app.route('/api/log-stream')
def log_stream():
    """SSE endpoint for real-time log streaming"""
    def generate():
        q = queue.Queue()
        with log_lock:
            log_queues.append(q)
        try:
            while True:
                try:
                    # Wait for log message with timeout
                    data = q.get(timeout=30)
                    yield f"data: {data}\n\n"
                except queue.Empty:
                    # Send heartbeat to keep connection alive
                    yield f"data: {json.dumps({'message': '', 'type': 'heartbeat'})}\n\n"
        except GeneratorExit:
            pass
        finally:
            with log_lock:
                if q in log_queues:
                    log_queues.remove(q)
    
    return Response(generate(), mimetype='text/event-stream', 
                    headers={'Cache-Control': 'no-cache', 'Connection': 'keep-alive'})

# ==================== API: SCAN ====================

@app.route('/api/scan/start', methods=['POST'])
def start_scan():
    data = request.json or {}
    date_folder = data.get('date', None)  # Có thể chỉ định ngày cụ thể
    folder_path = data.get('folder_path', INPUT_IMAGES_DIR)
    
    # Nếu có chỉ định ngày, quét trong thư mục ngày đó
    if date_folder:
        folder_path = os.path.join(INPUT_IMAGES_DIR, date_folder)
    
    if not os.path.exists(folder_path):
        return jsonify({'error': f'Thư mục không tồn tại: {folder_path}'}), 400
    
    files = get_image_files(folder_path)
    if not files:
        return jsonify({'error': 'Không tìm thấy file ảnh nào'}), 400
    
    task_id = f"task_{int(time.time() * 1000)}"
    task = ProcessingTask(task_id)
    tasks[task_id] = task
    
    thread = threading.Thread(target=run_processing, args=(task_id, files))
    thread.daemon = True
    thread.start()
    
    return jsonify({
        'success': True,
        'task_id': task_id,
        'message': f'Đã bắt đầu quét {len(files)} ảnh',
        'folder': folder_path,
        'total_files': len(files)
    })

@app.route('/api/scan/dates')
def list_date_folders():
    """Liệt kê các thư mục ngày có sẵn trong input_images"""
    date_folders = []
    if os.path.exists(INPUT_IMAGES_DIR):
        for item in os.listdir(INPUT_IMAGES_DIR):
            item_path = os.path.join(INPUT_IMAGES_DIR, item)
            if os.path.isdir(item_path):
                # Đếm số ảnh trong thư mục
                image_count = len(get_image_files(item_path))
                if image_count > 0:
                    date_folders.append({
                        'name': item,
                        'path': item_path,
                        'image_count': image_count
                    })
    
    # Sắp xếp theo tên (ngày)
    date_folders.sort(key=lambda x: x['name'])
    
    return jsonify({
        'success': True,
        'folders': date_folders,
        'total': len(date_folders)
    })

@app.route('/api/scan/all-dates', methods=['POST'])
def scan_all_dates():
    """Quét tất cả các thư mục ngày có ảnh trong input_images"""
    date_folders = []
    total_images = 0
    
    if not os.path.exists(INPUT_IMAGES_DIR):
        return jsonify({'error': 'Thư mục input_images không tồn tại'}), 400
    
    # Tìm tất cả thư mục con có ảnh
    for item in os.listdir(INPUT_IMAGES_DIR):
        item_path = os.path.join(INPUT_IMAGES_DIR, item)
        if os.path.isdir(item_path):
            images = get_image_files(item_path)
            if images:
                date_folders.append({
                    'name': item,
                    'path': item_path,
                    'images': images
                })
                total_images += len(images)
    
    if not date_folders:
        return jsonify({'error': 'Không tìm thấy thư mục ngày nào có ảnh'}), 400
    
    # Sắp xếp theo tên ngày
    date_folders.sort(key=lambda x: x['name'])
    
    # Tạo task và bắt đầu quét từng thư mục
    task_id = f"task_{int(time.time() * 1000)}"
    task = ProcessingTask(task_id)
    task.total = total_images
    tasks[task_id] = task
    
    def process_all_dates():
        task.status = 'running'
        task.start_time = datetime.now()
        
        for folder_info in date_folders:
            folder_name = folder_info['name']
            images = folder_info['images']
            
            for image_path in images:
                task.current_file = f"[{folder_name}] {os.path.basename(image_path)}"
                
                try:
                    result = process_image(image_path)
                    result['date_folder'] = folder_name  # Thêm thông tin thư mục ngày
                    task.results.append(result)
                except Exception as e:
                    task.errors.append({'file': image_path, 'error': str(e)})
                
                task.progress += 1
        
        # Export kết quả
        task.output_file = export_results(task)
        task.status = 'completed'
        task.end_time = datetime.now()
    
    thread = threading.Thread(target=process_all_dates, daemon=True)
    thread.start()
    
    return jsonify({
        'success': True,
        'task_id': task_id,
        'message': f'Đang quét {len(date_folders)} thư mục, tổng {total_images} ảnh',
        'folders': [{'name': f['name'], 'count': len(f['images'])} for f in date_folders],
        'total_images': total_images
    })

@app.route('/api/scan/status/<task_id>')
def get_scan_status(task_id):
    task = tasks.get(task_id)
    if task:
        return jsonify(task.to_dict())
    return jsonify({'error': 'Task không tồn tại'}), 404

@app.route('/api/scan/results/<task_id>')
def get_scan_results(task_id):
    task = tasks.get(task_id)
    if task:
        return jsonify({
            **task.to_dict(),
            'results': task.results,
            'errors': task.errors
        })
    return jsonify({'error': 'Task không tồn tại'}), 404

@app.route('/api/scan/tasks')
def get_all_tasks():
    return jsonify({'tasks': [t.to_dict() for t in tasks.values()]})

# ==================== API: DATABASE ====================

@app.route('/api/database/stats')
def get_database_stats():
    stats = {
        'total_persons': len(database),
        'total_branches': len(set(d['branch'] for d in database.values())),
        'branches': {}
    }
    
    for person_id, data in database.items():
        branch = data.get('branch', 'Unknown')
        if branch not in stats['branches']:
            stats['branches'][branch] = 0
        stats['branches'][branch] += 1
    
    return jsonify(stats)

@app.route('/api/database/scan', methods=['POST'])
def rescan_database():
    scan_database()
    stats = get_database_stats().get_json()
    return jsonify({
        'success': True,
        'message': 'Đã quét database',
        'stats': stats
    })

@app.route('/api/database/branches')
def get_branches():
    branches = []
    if os.path.exists(DATABASE_DIR):
        for branch in os.listdir(DATABASE_DIR):
            branch_path = os.path.join(DATABASE_DIR, branch)
            if os.path.isdir(branch_path) and not branch.startswith('.'):
                branches.append(branch)
    return jsonify({'branches': branches})

@app.route('/api/database/branches', methods=['POST'])
def add_branch():
    data = request.json or {}
    branch_name = data.get('name', '').strip()
    
    if not branch_name:
        return jsonify({'error': 'Tên chi nhánh không hợp lệ'}), 400
    
    branch_path = os.path.join(DATABASE_DIR, branch_name)
    os.makedirs(branch_path, exist_ok=True)
    
    return jsonify({
        'success': True,
        'message': f'Đã thêm chi nhánh: {branch_name}'
    })

@app.route('/api/database/persons/<branch>')
def get_persons(branch):
    persons = []
    branch_path = os.path.join(DATABASE_DIR, branch)
    
    if os.path.exists(branch_path):
        for person in os.listdir(branch_path):
            person_path = os.path.join(branch_path, person)
            if os.path.isdir(person_path):
                persons.append({
                    'name': person,
                    'branch': branch,
                    'person_id': f"{branch}/{person}"
                })
    
    return jsonify({'persons': persons})

# ==================== API: FILES ====================

@app.route('/api/files/input')
def list_input_files():
    files = []
    if os.path.exists(INPUT_IMAGES_DIR):
        for file in os.listdir(INPUT_IMAGES_DIR):
            file_path = os.path.join(INPUT_IMAGES_DIR, file)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_IMAGE_EXTENSIONS:
                    files.append({
                        'name': file,
                        'size': os.path.getsize(file_path),
                        'path': file_path
                    })
    
    return jsonify({
        'folder': INPUT_IMAGES_DIR,
        'files': files,
        'count': len(files)
    })

@app.route('/api/files/results')
def list_result_files():
    files = []
    if os.path.exists(RESULTS_DIR):
        for file in os.listdir(RESULTS_DIR):
            if file.endswith('.xlsx') or file.endswith('.csv'):
                file_path = os.path.join(RESULTS_DIR, file)
                files.append({
                    'name': file,
                    'size': os.path.getsize(file_path),
                    'path': file_path
                })
    
    files.sort(key=lambda x: x['name'], reverse=True)
    
    return jsonify({
        'folder': RESULTS_DIR,
        'files': files,
        'count': len(files)
    })

@app.route('/api/files/download/<filename>')
def download_file(filename):
    file_path = os.path.join(RESULTS_DIR, secure_filename(filename))
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File không tồn tại'}), 404

@app.route('/api/files/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'Không có file được upload'}), 400
    
    files = request.files.getlist('files')
    uploaded = []
    
    for file in files:
        if file.filename:
            filename = secure_filename(file.filename)
            ext = os.path.splitext(filename)[1].lower()
            
            if ext in SUPPORTED_IMAGE_EXTENSIONS:
                file_path = os.path.join(INPUT_IMAGES_DIR, filename)
                file.save(file_path)
                uploaded.append(filename)
    
    return jsonify({
        'success': True,
        'uploaded': uploaded,
        'count': len(uploaded)
    })

# ==================== API: CONFIG ====================

@app.route('/api/config')
def get_config():
    return jsonify({
        'input_dir': INPUT_IMAGES_DIR,
        'database_dir': DATABASE_DIR,
        'results_dir': RESULTS_DIR
    })

# ==================== STATIC FILES ====================

@app.route('/images/<path:filename>')
def serve_image(filename):
    file_path = os.path.join(INPUT_IMAGES_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path)
    return jsonify({'error': 'File không tồn tại'}), 404

# ==================== API: DỰ ÁN & QUẢN LÝ ẢNH ====================

@app.route('/api/projects', methods=['GET'])
def list_projects():
    """Lấy danh sách dự án kèm thống kê số lượng nhân viên và ảnh camera"""
    projects_list = []
    names = get_all_projects()
    for name in names:
        p_dir = os.path.join(PORTRAIT_DIR, name)
        i_dir = os.path.join(INPUT_IMAGES_DIR, name)
        emp_count = 0
        if os.path.exists(p_dir):
            for item in os.listdir(p_dir):
                sub_p = os.path.join(p_dir, item)
                if os.path.isdir(sub_p):
                    emp_count += 1
                elif os.path.splitext(item)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS:
                    emp_count += 1
        cam_days = 0
        cam_imgs = 0
        if os.path.exists(i_dir):
            for item in os.listdir(i_dir):
                sub_p = os.path.join(i_dir, item)
                if os.path.isdir(sub_p):
                    imgs = [f for f in os.listdir(sub_p) if os.path.splitext(f)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS]
                    if imgs:
                        cam_days += 1
                        cam_imgs += len(imgs)
        projects_list.append({
            'name': name,
            'employee_count': emp_count,
            'camera_days_count': cam_days,
            'camera_total_images': cam_imgs
        })
    return jsonify({
        'success': True,
        'projects': projects_list,
        'default_project': DEFAULT_PROJECT_NAME
    })

@app.route('/api/projects/create', methods=['POST'])
def create_project():
    """Tạo dự án mới đầy đủ cấu trúc thư mục"""
    data = request.json or {}
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Tên dự án không được để trống'}), 400
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    p_dir, i_dir = ensure_project_structure(safe_name)
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True, 'project': safe_name})

@app.route('/api/projects/rename', methods=['POST'])
def rename_project():
    """Đổi tên dự án"""
    data = request.json or {}
    old_name = data.get('old_name', '').strip()
    new_name = data.get('new_name', '').strip()
    if not old_name or not new_name:
        return jsonify({'error': 'Tên dự án không hợp lệ'}), 400
    safe_old = re.sub(r'[<>:"/\\|?*]', '_', old_name)
    safe_new = re.sub(r'[<>:"/\\|?*]', '_', new_name)
    
    old_p = os.path.join(PORTRAIT_DIR, safe_old)
    new_p = os.path.join(PORTRAIT_DIR, safe_new)
    if os.path.exists(old_p):
        os.rename(old_p, new_p)
        
    old_i = os.path.join(INPUT_IMAGES_DIR, safe_old)
    new_i = os.path.join(INPUT_IMAGES_DIR, safe_new)
    if os.path.exists(old_i):
        os.rename(old_i, new_i)
        
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True, 'name': safe_new})

@app.route('/api/projects/delete', methods=['POST'])
def delete_project():
    """Xóa dự án"""
    data = request.json or {}
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Tên dự án không hợp lệ'}), 400
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    p_dir = os.path.join(PORTRAIT_DIR, safe_name)
    i_dir = os.path.join(INPUT_IMAGES_DIR, safe_name)
    
    force = data.get('force', False)
    has_content = False
    if os.path.exists(p_dir) and os.listdir(p_dir):
        has_content = True
    if os.path.exists(i_dir) and os.listdir(i_dir):
        has_content = True
    if has_content and not force:
        return jsonify({'error': 'Dự án này đang có dữ liệu nhân viên/ảnh camera. Vui lòng chuyển nhân viên hoặc xóa ảnh trước khi xóa dự án'}), 400
        
    if os.path.exists(p_dir):
        shutil.rmtree(p_dir, ignore_errors=True)
    if os.path.exists(i_dir):
        shutil.rmtree(i_dir, ignore_errors=True)
        
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True})

# ---------- API: ẢNH CAMERA THEO NGÀY ----------

@app.route('/api/photos/daily', methods=['GET'])
def get_daily_photo_stats():
    """Lấy danh sách 31 ngày kèm số ảnh camera của dự án (hỗ trợ cả dạng '01' và 'YYYY-MM-01')"""
    project = request.args.get('project', DEFAULT_PROJECT_NAME).strip()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    proj_dir = os.path.join(INPUT_IMAGES_DIR, safe_proj)
    days_data = []

    subdirs = []
    if os.path.exists(proj_dir):
        subdirs = [d for d in os.listdir(proj_dir) if os.path.isdir(os.path.join(proj_dir, d))]

    for d in range(1, 32):
        d_str = f"{d:02d}"
        matching_dirs = [os.path.join(proj_dir, sd) for sd in subdirs if sd == d_str or sd.endswith(f"-{d_str}") or sd.startswith(f"{d_str}-")]
        if not matching_dirs:
            if os.path.exists(os.path.join(INPUT_IMAGES_DIR, d_str)):
                matching_dirs.append(os.path.join(INPUT_IMAGES_DIR, d_str))

        counted_files = set()
        for md in matching_dirs:
            if os.path.exists(md):
                for f in os.listdir(md):
                    if os.path.splitext(f)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS:
                        counted_files.add(f)

        days_data.append({
            'day': d_str,
            'image_count': len(counted_files),
            'has_images': len(counted_files) > 0
        })
    return jsonify({'success': True, 'project': project, 'days': days_data})

@app.route('/api/photos/daily/<day>', methods=['GET'])
def get_daily_photos(day):
    """Lấy danh sách ảnh camera của 1 ngày trong dự án (hỗ trợ cả dạng '01' và 'YYYY-MM-01')"""
    project = request.args.get('project', DEFAULT_PROJECT_NAME).strip()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    day_str = str(day).zfill(2)
    proj_dir = os.path.join(INPUT_IMAGES_DIR, safe_proj)

    matching_dirs = []
    if os.path.exists(proj_dir):
        for sd in os.listdir(proj_dir):
            if os.path.isdir(os.path.join(proj_dir, sd)):
                if sd == day_str or sd.endswith(f"-{day_str}") or sd.startswith(f"{day_str}-"):
                    matching_dirs.append(os.path.join(proj_dir, sd))

    if not matching_dirs and os.path.exists(os.path.join(INPUT_IMAGES_DIR, day_str)):
        matching_dirs.append(os.path.join(INPUT_IMAGES_DIR, day_str))

    photos = []
    seen_filenames = set()
    for md in matching_dirs:
        for f in sorted(os.listdir(md)):
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED_IMAGE_EXTENSIONS and f not in seen_filenames:
                seen_filenames.add(f)
                f_path = os.path.join(md, f)
                size_kb = round(os.path.getsize(f_path) / 1024, 1)
                photos.append({
                    'filename': f,
                    'size_kb': size_kb,
                    'url': f"/api/photos/view/daily?project={safe_proj}&day={day_str}&filename={f}"
                })
    return jsonify({'success': True, 'project': project, 'day': day_str, 'photos': photos, 'total': len(photos)})

@app.route('/api/photos/daily/upload', methods=['POST'])
def upload_daily_photos():
    """Tải lên nhiều ảnh camera vào 1 ngày của dự án"""
    project = request.form.get('project', DEFAULT_PROJECT_NAME).strip()
    day = request.form.get('day', '01').strip()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    day_str = str(day).zfill(2)
    
    target_dir = os.path.join(INPUT_IMAGES_DIR, safe_proj, day_str)
    os.makedirs(target_dir, exist_ok=True)
    
    files = request.files.getlist('files') or request.files.getlist('photos')
    if not files and 'file' in request.files:
        files = [request.files['file']]
            
    if not files:
        return jsonify({'error': 'Không có file ảnh nào được gửi lên'}), 400
        
    saved_count = 0
    for f in files:
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext in SUPPORTED_IMAGE_EXTENSIONS:
            safe_filename = re.sub(r'[<>:"/\\|?*]', '_', f.filename)
            save_path = os.path.join(target_dir, safe_filename)
            f.save(save_path)
            saved_count += 1
            
    return jsonify({'success': True, 'saved_count': saved_count, 'day': day_str, 'project': project})

@app.route('/api/photos/daily/delete', methods=['POST'])
def delete_daily_photo():
    """Xóa 1 ảnh hoặc xóa toàn bộ ảnh của ngày (hỗ trợ cả dạng '01' và 'YYYY-MM-01')"""
    data = request.json or {}
    project = data.get('project', DEFAULT_PROJECT_NAME).strip()
    day = str(data.get('day', '01')).zfill(2)
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    filename = data.get('filename')
    delete_all = data.get('delete_all', False)
    
    # Tìm tất cả thư mục khớp với ngày (giống get_daily_photos)
    proj_dir = os.path.join(INPUT_IMAGES_DIR, safe_proj)
    matching_dirs = []
    if os.path.exists(proj_dir):
        for sd in os.listdir(proj_dir):
            if os.path.isdir(os.path.join(proj_dir, sd)):
                if sd == day or sd.endswith(f"-{day}") or sd.startswith(f"{day}-"):
                    matching_dirs.append(os.path.join(proj_dir, sd))
    
    if not matching_dirs and os.path.exists(os.path.join(INPUT_IMAGES_DIR, day)):
        matching_dirs.append(os.path.join(INPUT_IMAGES_DIR, day))
        
    if not matching_dirs:
        return jsonify({'success': True})
        
    if delete_all:
        deleted_count = 0
        for d_path in matching_dirs:
            for f in os.listdir(d_path):
                if os.path.splitext(f)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS:
                    try:
                        os.remove(os.path.join(d_path, f))
                        deleted_count += 1
                    except Exception:
                        pass
        return jsonify({'success': True, 'message': f'Đã xóa toàn bộ {deleted_count} ảnh ngày {day}'})
    elif filename:
        for d_path in matching_dirs:
            file_path = os.path.join(d_path, filename)
            if os.path.exists(file_path):
                os.remove(file_path)
                return jsonify({'success': True, 'message': f'Đã xóa ảnh {filename}'})
        return jsonify({'success': True, 'message': f'Ảnh {filename} không tồn tại'})
    return jsonify({'error': 'Yêu cầu không hợp lệ'}), 400

# ---------- API: ẢNH CHÂN DUNG NHÂN VIÊN & THUYÊN CHUYỂN ----------

@app.route('/api/portraits', methods=['GET'])
def get_employees_portraits():
    """Lấy danh sách nhân viên và ảnh chân dung theo dự án"""
    project = request.args.get('project', DEFAULT_PROJECT_NAME).strip()
    search = request.args.get('search', '').strip().lower()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    proj_dir = os.path.join(PORTRAIT_DIR, safe_proj)
    
    employees = []
    if os.path.exists(proj_dir):
        for item in os.listdir(proj_dir):
            item_path = os.path.join(proj_dir, item)
            if os.path.isdir(item_path):
                emp_name = item
                if search and search not in emp_name.lower():
                    continue
                imgs = [
                    f for f in sorted(os.listdir(item_path))
                    if os.path.splitext(f)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS
                ]
                avatar_url = ""
                if imgs:
                    avatar_url = f"/api/photos/view/portrait?project={safe_proj}&person={emp_name}&filename={imgs[0]}"
                employees.append({
                    'name': emp_name,
                    'project': project,
                    'image_count': len(imgs),
                    'images': imgs,
                    'avatar_url': avatar_url
                })
            elif os.path.splitext(item)[1].lower() in SUPPORTED_IMAGE_EXTENSIONS:
                emp_name = os.path.splitext(item)[0]
                if search and search not in emp_name.lower():
                    continue
                avatar_url = f"/api/photos/view/portrait?project={safe_proj}&person=&filename={item}"
                employees.append({
                    'name': emp_name,
                    'project': project,
                    'image_count': 1,
                    'images': [item],
                    'avatar_url': avatar_url
                })
    employees.sort(key=lambda x: x['name'])
    return jsonify({'success': True, 'project': project, 'employees': employees, 'total': len(employees)})

@app.route('/api/portraits/employee/create', methods=['POST'])
def create_employee():
    """Thêm nhân viên mới vào dự án"""
    data = request.json or {}
    project = data.get('project', DEFAULT_PROJECT_NAME).strip()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Tên nhân viên không được để trống'}), 400
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    emp_dir = os.path.join(PORTRAIT_DIR, safe_proj, safe_name)
    os.makedirs(emp_dir, exist_ok=True)
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True, 'name': safe_name, 'project': safe_proj})

@app.route('/api/portraits/employee/upload', methods=['POST'])
def upload_employee_photos():
    """Tải lên ảnh chân dung cho nhân viên"""
    project = request.form.get('project', DEFAULT_PROJECT_NAME).strip()
    name = request.form.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Tên nhân viên không hợp lệ'}), 400
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    emp_dir = os.path.join(PORTRAIT_DIR, safe_proj, safe_name)
    os.makedirs(emp_dir, exist_ok=True)
    
    files = request.files.getlist('files') or request.files.getlist('photos')
    if not files and 'file' in request.files:
        files = [request.files['file']]
    if not files:
        return jsonify({'error': 'Không có file ảnh nào được gửi lên'}), 400
        
    saved = []
    for f in files:
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext in SUPPORTED_IMAGE_EXTENSIONS:
            safe_fname = re.sub(r'[<>:"/\\|?*]', '_', f.filename)
            save_path = os.path.join(emp_dir, safe_fname)
            f.save(save_path)
            saved.append(safe_fname)
            
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True, 'saved_count': len(saved), 'files': saved})

@app.route('/api/portraits/employee/delete-photo', methods=['POST'])
def delete_employee_photo():
    """Xóa 1 ảnh chân dung của nhân viên"""
    data = request.json or {}
    project = data.get('project', DEFAULT_PROJECT_NAME).strip()
    name = data.get('name', '').strip()
    filename = data.get('filename', '').strip()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    
    p_path = os.path.join(PORTRAIT_DIR, safe_proj, safe_name, filename)
    if os.path.exists(p_path):
        os.remove(p_path)
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True})

@app.route('/api/portraits/employee/delete', methods=['POST'])
def delete_employee():
    """Xóa nhân viên và toàn bộ ảnh chân dung"""
    data = request.json or {}
    project = data.get('project', DEFAULT_PROJECT_NAME).strip()
    name = data.get('name', '').strip()
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project)
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    
    emp_dir = os.path.join(PORTRAIT_DIR, safe_proj, safe_name)
    if os.path.exists(emp_dir):
        shutil.rmtree(emp_dir, ignore_errors=True)
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
    return jsonify({'success': True})

@app.route('/api/portraits/employee/transfer', methods=['POST'])
def transfer_employee():
    """Thuyên chuyển nhân viên từ dự án này sang dự án khác"""
    data = request.json or {}
    source_proj = data.get('source_project', '').strip()
    target_proj = data.get('target_project', '').strip()
    name = data.get('name', '').strip()
    
    if not source_proj or not target_proj or not name:
        return jsonify({'error': 'Thông tin thuyên chuyển không đầy đủ'}), 400
        
    if source_proj == target_proj:
        return jsonify({'error': 'Dự án đích phải khác dự án nguồn'}), 400
        
    safe_source = re.sub(r'[<>:"/\\|?*]', '_', source_proj)
    safe_target = re.sub(r'[<>:"/\\|?*]', '_', target_proj)
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
    
    src_dir = os.path.join(PORTRAIT_DIR, safe_source, safe_name)
    dst_proj_dir = os.path.join(PORTRAIT_DIR, safe_target)
    dst_dir = os.path.join(dst_proj_dir, safe_name)
    
    if not os.path.exists(src_dir):
        return jsonify({'error': f'Không tìm thấy nhân viên {name} trong dự án {source_proj}'}), 404
        
    os.makedirs(dst_proj_dir, exist_ok=True)
    
    if not os.path.exists(dst_dir):
        shutil.move(src_dir, dst_dir)
    else:
        for f in os.listdir(src_dir):
            s_file = os.path.join(src_dir, f)
            d_file = os.path.join(dst_dir, f)
            if not os.path.exists(d_file):
                shutil.move(s_file, d_file)
        shutil.rmtree(src_dir, ignore_errors=True)
        
    matcher = get_face_matcher()
    if matcher:
        matcher.reload_portraits()
        
    return jsonify({
        'success': True,
        'message': f'Đã thuyên chuyển nhân viên {name} từ {source_proj} sang {target_proj}'
    })

# ---------- API: PHỤC VỤ XEM ẢNH AN TOÀN UTF-8 ----------

@app.route('/api/photos/view/portrait')
def view_portrait_photo():
    """Xem ảnh chân dung an toàn với tên tiếng Việt"""
    project = request.args.get('project', '').strip()
    person = request.args.get('person', '').strip()
    filename = request.args.get('filename', '').strip()
    if not filename:
        return "File not found", 404
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project) if project else ''
    safe_person = re.sub(r'[<>:"/\\|?*]', '_', person) if person else ''
    
    if safe_person:
        file_path = os.path.join(PORTRAIT_DIR, safe_proj, safe_person, filename)
    elif safe_proj:
        file_path = os.path.join(PORTRAIT_DIR, safe_proj, filename)
    else:
        file_path = os.path.join(PORTRAIT_DIR, filename)
        
    if not os.path.exists(file_path):
        file_path = os.path.join(PORTRAIT_DIR, safe_person, filename)
        
    if os.path.exists(file_path):
        return send_file(os.path.abspath(file_path))
    return "Not found", 404

@app.route('/api/photos/view/daily')
def view_daily_photo():
    """Xem ảnh camera theo ngày an toàn với tên tiếng Việt"""
    project = request.args.get('project', '').strip()
    day = str(request.args.get('day', '')).zfill(2)
    filename = request.args.get('filename', '').strip()
    if not filename:
        return "File not found", 404
    safe_proj = re.sub(r'[<>:"/\\|?*]', '_', project) if project else ''
    
    file_path = os.path.join(INPUT_IMAGES_DIR, safe_proj, day, filename)
    if not os.path.exists(file_path):
        proj_dir = os.path.join(INPUT_IMAGES_DIR, safe_proj)
        if os.path.exists(proj_dir):
            for sd in os.listdir(proj_dir):
                if sd == day or sd.endswith(f"-{day}") or sd.startswith(f"{day}-"):
                    candidate = os.path.join(proj_dir, sd, filename)
                    if os.path.exists(candidate):
                        file_path = candidate
                        break
    if not os.path.exists(file_path):
        file_path = os.path.join(INPUT_IMAGES_DIR, day, filename)
        
    if os.path.exists(file_path):
        return send_file(os.path.abspath(file_path))
    return "Not found", 404

# ==================== API: ATTENDANCE ====================

# Đường dẫn cho attendance

@app.route('/api/attendance/analyze', methods=['POST'])
def analyze_attendance():
    """Phân tích file chấm công và tìm các bản ghi thiếu"""
    try:
        from src.attendance_processor import AttendanceProcessor
        
        processor = AttendanceProcessor(CHAMCONG_DIR)
        processor.scan_all_files()
        missing = processor.get_missing_records()
        summary = processor.get_summary()
        
        return jsonify({
            'success': True,
            'summary': summary,
            'missing_records': missing
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/attendance/export', methods=['POST'])
def export_attendance():
    """Xuất file Word giải trình với ảnh"""
    try:
        from src.attendance_processor import AttendanceProcessor
        from src.word_exporter import WordExporter
        
        data = request.json or {}
        project_name = data.get('project_name', 'Chung cư Tân Thuận Đông')
        month = data.get('month', None)
        
        # Xử lý chấm công
        processor = AttendanceProcessor(CHAMCONG_DIR)
        processor.scan_all_files()
        missing = processor.get_missing_records()
        
        if not missing:
            return jsonify({
                'success': True,
                'message': 'Không có bản ghi thiếu cần giải trình',
                'output_file': None
            })
        
        # Xuất Word
        exporter = WordExporter(PORTRAIT_DIR, RESULTS_DIR)
        output_file = exporter.create_summary_document(missing, project_name, month)
        
        return jsonify({
            'success': True,
            'message': f'Đã xuất {len(missing)} bản ghi thiếu',
            'output_file': os.path.basename(output_file),
            'total_missing': len(missing)
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/attendance/portraits')
def get_portrait_stats():
    """Thống kê ảnh chân dung"""
    try:
        from src.word_exporter import WordExporter
        
        exporter = WordExporter(PORTRAIT_DIR, RESULTS_DIR)
        stats = exporter.get_portrait_stats()
        return jsonify({'success': True, **stats})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ==================== API: FULL ANALYSIS (NEW) ====================

# Face matcher instance (lazy loaded)
_face_matcher = None

def get_face_matcher():
    """Get or create face matcher instance"""
    global _face_matcher
    try:
        from src.face_matcher import FaceMatcher
        if _face_matcher is None:
            send_log("⏳ Dang khoi tao Face Matcher (DeepFace)...", "info")
            portrait_dir = resolve_portrait_dir(BASE_DIR)
            _face_matcher = FaceMatcher(portrait_dir, log_callback=send_log)
            send_log(
                f"✅ Face Matcher san sang. PortraitDir={portrait_dir} "
                f"(n={len(_face_matcher.portrait_cache)}, imgs={_count_images_in_dir(portrait_dir)})",
                "success"
            )
        # If cache empty, retry with scan-based directory
        if _face_matcher and len(_face_matcher.portrait_cache) == 0:
            alt_dir = resolve_portrait_dir_by_scan(BASE_DIR)
            if alt_dir:
                send_log(
                    f"🔁 Cache rong, thu lai PortraitDir={alt_dir} (imgs={_count_images_in_dir(alt_dir)})",
                    "warning"
                )
                _face_matcher = FaceMatcher(alt_dir, log_callback=send_log)
                send_log(
                    f"✅ Face Matcher san sang. PortraitDir={alt_dir} "
                    f"(n={len(_face_matcher.portrait_cache)})",
                    "success"
                )
    except Exception as e:
        send_log(f"❌ Loi khoi tao FaceMatcher: {e}", "error")
        import traceback
        traceback.print_exc()
        return None
    return _face_matcher

@app.route('/api/analyze-full', methods=['POST'])
def analyze_full():
    """Phân tích tổng hợp: tìm ngày thiếu + match ảnh camera bằng nhận diện khuôn mặt"""
    try:
        from src.attendance_processor import AttendanceProcessor
        
        # Step 1: Phân tích chấm công
        send_log(" Step 1: Đang phân tích file chấm công...", "info")
        processor = AttendanceProcessor(CHAMCONG_DIR)
        processor.scan_all_files()
        missing_records = processor.get_missing_records()
        summary = processor.get_summary()
        send_log(f" Tìm thấy {len(missing_records)} bản ghi thiếu từ {summary.get('total_persons', 0)} người", "info")
        
        # Step 2: Khởi tạo face matcher
        send_log(" Step 2: Đang khởi tạo Face Matcher...", "info")
        matcher = get_face_matcher()
        if matcher:
            send_log(" Face Matcher đã sẵn sàng", "success")
        else:
            send_log(" Face Matcher không khả dụng, sẽ dùng fallback", "warning")
        
        # Step 3: Match ảnh camera cho mỗi bản ghi thiếu
        send_log(f" Step 3: Bắt đầu matching ảnh cho {len(missing_records)} bản ghi...", "info")
        matched_count = 0
        
        for i, record in enumerate(missing_records):
            date_str = record['date']  # format: dd/mm/yyyy
            day = date_str.split('/')[0].zfill(2)  # extract dd
            person_name = record['person_name']
            
            # Tìm thư mục ngày tương ứng
            day_folder = os.path.join(INPUT_IMAGES_DIR, day)
            
            record['matched_image'] = None
            
            if os.path.exists(day_folder):
                images = get_image_files(day_folder)
                send_log(f"  [{i+1}/{len(missing_records)}] {person_name} (ngày {day}): Tìm thấy {len(images)} ảnh trong thư mục", "default")
                
                if images and matcher:
                    # Dùng face recognition để tìm ảnh match
                    try:
                        matched_image = matcher.match_face_in_images(person_name, images)
                        if matched_image:
                            record['matched_image'] = matched_image
                            matched_count += 1
                            send_log(f"  [{i+1}/{len(missing_records)}] ✓ {person_name} -> {os.path.basename(matched_image)}", "success")
                        else:
                            send_log(f"  [{i+1}/{len(missing_records)}]  {person_name}: Không tìm thấy ảnh match", "warning")
                    except Exception as match_err:
                        send_log(f"  [{i+1}/{len(missing_records)}]  {person_name}: Lỗi matcher ({match_err})", "error")
                elif images:
                    send_log(f"  [{i+1}/{len(missing_records)}]  {person_name}: FaceMatcher chưa sẵn sàng, bỏ qua", "warning")
                else:
                    send_log(f"  [{i+1}/{len(missing_records)}]  Thư mục {day} rỗng", "warning")
            else:
                send_log(f"  [{i+1}/{len(missing_records)}]  Không tìm thấy thư mục: {day_folder}", "error")
                record['matched_image'] = None
        
        summary['total_matched'] = matched_count
        send_log(f" Hoàn thành! Matched {matched_count}/{len(missing_records)} bản ghi", "success")
        
        return jsonify({
            'success': True,
            'summary': summary,
            'records': missing_records
        })
    except Exception as e:
        import traceback
        send_log(f" Lỗi: {e}", "error")
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/matched-image/<path:filepath>')
def serve_matched_image(filepath):
    """Serve ảnh đã match"""
    # Decode URL path nếu cần
    import urllib.parse
    filepath = urllib.parse.unquote(filepath)
    
    # Thử với đường dẫn nguyên gốc
    if os.path.exists(filepath):
        return send_file(filepath)
    
    # Thử với đường dẫn tuyệt đối từ BASE_DIR
    abs_path = os.path.join(BASE_DIR, filepath)
    if os.path.exists(abs_path):
        return send_file(abs_path)
    
    # Thử thay thế backslash/forward slash
    filepath_fixed = filepath.replace('/', os.sep).replace('\\', os.sep)
    abs_path_fixed = os.path.join(BASE_DIR, filepath_fixed)
    if os.path.exists(abs_path_fixed):
        return send_file(abs_path_fixed)
    
    print(f"[serve_matched_image] File không tồn tại:")
    print(f"  filepath: {filepath}")
    print(f"  abs_path: {abs_path}")
    print(f"  abs_path_fixed: {abs_path_fixed}")
    print(f"  BASE_DIR: {BASE_DIR}")
    
    return jsonify({'error': 'File không tồn tại', 'filepath': filepath}), 404

@app.route('/api/export-word', methods=['POST'])
def export_word():
    """Xuất file Word với ảnh camera đã match"""
    try:
        from docx import Document
        from docx.shared import Inches, Cm
        
        data = request.json or {}
        project_name = data.get('project_name', 'Chung cư Tân Thuận Đông')
        month = data.get('month', '')
        records = data.get('records', [])
        
        if not records:
            return jsonify({'success': False, 'error': 'Không có dữ liệu để xuất'})
        
        # Tạo document
        doc = Document()
        
        # Tiêu đề
        title = doc.add_paragraph()
        title.add_run(f'GIẢI TRÌNH CHẤM CÔNG - {project_name}').bold = True
        title.alignment = 1  # Center
        
        doc.add_paragraph(f'Tháng: {month}')
        doc.add_paragraph()
        
        # Tạo bảng
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        # Header
        headers = ['TÊN', 'NGÀY', 'GIẢI TRÌNH', 'HÌNH ẢNH', 'GHI CHÚ']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        
        # Thêm dữ liệu
        for record in records:
            row = table.add_row()
            row.cells[0].text = record.get('person_name', '')
            row.cells[1].text = record.get('date', '')
            row.cells[2].text = record.get('issue_description', 'Nhân viên có trực, bổ sung')
            
            # Thêm ảnh nếu có
            matched_image = record.get('matched_image')
            if matched_image and os.path.exists(matched_image):
                try:
                    run = row.cells[3].paragraphs[0].add_run()
                    run.add_picture(matched_image, width=Cm(3))
                except Exception:
                    row.cells[3].text = '[Lỗi ảnh]'
            else:
                row.cells[3].text = '[Không có ảnh]'
            
            row.cells[4].text = ''
        
        # Lưu file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"GIAI_TRINH_{project_name.replace(' ', '_')}_{timestamp}.docx"
        output_path = os.path.join(RESULTS_DIR, filename)
        doc.save(output_path)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'path': output_path
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

# ==================== API: PDF EXTRACTION ====================

# Import PDF extractor
try:
    from src import pdf_extractor
    PDF_EXTRACTOR_AVAILABLE = True
except ImportError:
    PDF_EXTRACTOR_AVAILABLE = False

# Thư mục cho PDF
PDF_UPLOAD_DIR = os.path.join(BASE_DIR, "pdf_uploads")
PDF_OUTPUT_DIR = os.path.join(BASE_DIR, "pdf_extracted")
PDF_FACE_OUTPUT_DIR = os.path.join(BASE_DIR, "pdf_face_output")
os.makedirs(PDF_UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
os.makedirs(PDF_FACE_OUTPUT_DIR, exist_ok=True)

pdf_face_tasks = {}


class PDFFaceTask:
    def __init__(self, task_id):
        self.task_id = task_id
        self.status = 'pending'
        self.progress = 0
        self.total = 0
        self.current = ''
        self.files = []
        self.errors = []
        self.start_time = None
        self.end_time = None

    def to_dict(self):
        return {
            'task_id': self.task_id,
            'status': self.status,
            'progress': self.progress,
            'total': self.total,
            'current': self.current,
            'files': self.files,
            'errors': self.errors,
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
        }

@app.route('/api/pdf/check')
def pdf_check_available():
    """Kiểm tra xem tính năng PDF có sẵn không"""
    available = PDF_EXTRACTOR_AVAILABLE and pdf_extractor.is_available()
    return jsonify({
        'available': available,
        'message': 'Sẵn sàng' if available else 'Cần cài đặt: pip install pdf2docx PyMuPDF python-docx'
    })

@app.route('/api/pdf/upload', methods=['POST'])
def pdf_upload():
    """Upload file PDF"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Không có file được upload'}), 400
    
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'error': 'Tên file không hợp lệ'}), 400
    
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'success': False, 'error': 'Chỉ chấp nhận file PDF'}), 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(PDF_UPLOAD_DIR, filename)
    file.save(filepath)
    
    return jsonify({
        'success': True,
        'filename': filename,
        'filepath': filepath,
        'size': os.path.getsize(filepath)
    })

@app.route('/api/pdf/extract', methods=['POST'])
def pdf_extract():
    """Bắt đầu tách PDF thành các file Word"""
    if not PDF_EXTRACTOR_AVAILABLE or not pdf_extractor.is_available():
        return jsonify({
            'success': False, 
            'error': 'PDF Extractor không khả dụng. Cần cài đặt: pip install pdf2docx PyMuPDF python-docx'
        }), 400
    
    data = request.json or {}
    filename = data.get('filename')
    
    if not filename:
        return jsonify({'success': False, 'error': 'Thiếu tên file'}), 400
    
    filepath = os.path.join(PDF_UPLOAD_DIR, secure_filename(filename))
    
    if not os.path.exists(filepath):
        return jsonify({'success': False, 'error': 'File PDF không tồn tại'}), 404
    
    # Tạo thư mục output riêng cho file này
    base_name = os.path.splitext(filename)[0]
    output_dir = os.path.join(PDF_OUTPUT_DIR, base_name)
    
    # Bắt đầu task trong background
    task_id = pdf_extractor.start_extraction_task(filepath, output_dir)
    
    return jsonify({
        'success': True,
        'task_id': task_id,
        'message': 'Đã bắt đầu tách PDF',
        'output_dir': output_dir
    })

@app.route('/api/pdf/status/<task_id>')
def pdf_status(task_id):
    """Kiểm tra tiến độ tách PDF"""
    if not PDF_EXTRACTOR_AVAILABLE:
        return jsonify({'error': 'PDF Extractor không khả dụng'}), 400
    
    task = pdf_extractor.get_task(task_id)
    if not task:
        return jsonify({'error': 'Task không tồn tại'}), 404
    
    return jsonify(task.to_dict())

@app.route('/api/pdf/files')
def pdf_list_files():
    """Liệt kê các file Word đã tách"""
    files = []
    
    if os.path.exists(PDF_OUTPUT_DIR):
        for folder in os.listdir(PDF_OUTPUT_DIR):
            folder_path = os.path.join(PDF_OUTPUT_DIR, folder)
            if os.path.isdir(folder_path):
                folder_files = pdf_extractor.list_extracted_files(folder_path) if PDF_EXTRACTOR_AVAILABLE else []
                files.append({
                    'folder': folder,
                    'path': folder_path,
                    'files': folder_files,
                    'count': len(folder_files)
                })
    
    return jsonify({
        'success': True,
        'folders': files,
        'output_dir': PDF_OUTPUT_DIR
    })

@app.route('/api/pdf/files/<folder>')
def pdf_list_folder_files(folder):
    """Liệt kê các file Word trong một thư mục"""
    folder_path = os.path.join(PDF_OUTPUT_DIR, secure_filename(folder))
    
    if not os.path.exists(folder_path):
        return jsonify({'error': 'Thư mục không tồn tại'}), 404
    
    files = pdf_extractor.list_extracted_files(folder_path) if PDF_EXTRACTOR_AVAILABLE else []
    
    return jsonify({
        'success': True,
        'folder': folder,
        'files': files,
        'count': len(files)
    })

@app.route('/api/pdf/download/<folder>/<filename>')
def pdf_download(folder, filename):
    """Tải file Word đã tách"""
    # Decode URL-encoded names (không dùng secure_filename vì nó xóa tiếng Việt)
    from urllib.parse import unquote
    folder = unquote(folder)
    filename = unquote(filename)
    
    # Bảo vệ path traversal
    if '..' in folder or '..' in filename or '/' in folder or '\\' in folder:
        return jsonify({'error': 'Invalid path'}), 400
    
    file_path = os.path.join(PDF_OUTPUT_DIR, folder, filename)
    
    # Kiểm tra file nằm trong thư mục cho phép
    if not os.path.abspath(file_path).startswith(os.path.abspath(PDF_OUTPUT_DIR)):
        return jsonify({'error': 'Invalid path'}), 400
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File không tồn tại', 'path': file_path}), 404

@app.route('/api/pdf/uploads')
def pdf_list_uploads():
    """Liệt kê các file PDF đã upload"""
    files = []
    
    if os.path.exists(PDF_UPLOAD_DIR):
        for file in os.listdir(PDF_UPLOAD_DIR):
            if file.lower().endswith('.pdf'):
                filepath = os.path.join(PDF_UPLOAD_DIR, file)
                files.append({
                    'name': file,
                    'size': os.path.getsize(filepath),
                    'modified': datetime.fromtimestamp(os.path.getmtime(filepath)).isoformat()
                })
    
    files.sort(key=lambda x: x['modified'], reverse=True)
    
    return jsonify({
        'success': True,
        'files': files,
        'upload_dir': PDF_UPLOAD_DIR
    })


@app.route('/api/pdf/face/analyze', methods=['POST'])
def pdf_face_analyze():
    """Phân tích khuôn mặt từ các file Word đã tách ra từ PDF"""
    try:
        data = request.json or {}
        folder = data.get('folder')
        threshold = data.get('distance_threshold')
        try:
            threshold = float(threshold) if threshold not in (None, '') else None
        except Exception:
            threshold = None

        if not folder:
            return jsonify({'success': False, 'error': 'Thiếu tên thư mục'}), 400
        if '..' in folder or '/' in folder or '\\' in folder:
            return jsonify({'success': False, 'error': 'Tên thư mục không hợp lệ'}), 400

        input_dir = os.path.join(PDF_OUTPUT_DIR, folder)
        if not os.path.isdir(input_dir):
            return jsonify({'success': False, 'error': 'Thư mục PDF đã tách không tồn tại'}), 404

        output_dir = os.path.join(PDF_FACE_OUTPUT_DIR, folder)
        os.makedirs(output_dir, exist_ok=True)

        task_id = f"pdf_face_{int(time.time() * 1000)}"
        task = PDFFaceTask(task_id)
        pdf_face_tasks[task_id] = task

        project_name = data.get('project') or DEFAULT_PROJECT_NAME
        p_dir = os.path.join(PORTRAIT_DIR, project_name) if os.path.exists(os.path.join(PORTRAIT_DIR, project_name)) else PORTRAIT_DIR
        i_dir = os.path.join(INPUT_IMAGES_DIR, project_name) if os.path.exists(os.path.join(INPUT_IMAGES_DIR, project_name)) else INPUT_IMAGES_DIR

        def _run():
            task.status = 'running'
            task.start_time = datetime.now()
            try:
                send_log(f" Bắt đầu phân tích khuôn mặt cho thư mục PDF: {folder} (Dự án: {project_name})", "info")
                matcher = get_face_matcher()
                if matcher:
                    send_log("✅ Face Matcher đã sẵn sàng", "success")
                else:
                    send_log("⚠️ Face Matcher không khả dụng, sẽ bỏ qua tìm ảnh camera", "warning")

                from src.pdf_face_analyzer import PDFFaceAnalyzer
                analyzer = PDFFaceAnalyzer(
                    p_dir,
                    i_dir,
                    matcher,
                    accuracy_mode=True,
                    match_distance_threshold=threshold,
                    log_detail=True
                )

                def _log(msg, t='default'):
                    send_log(msg, t)

                def _progress(completed, total, name, file_path):
                    task.total = total
                    task.progress = completed
                    task.current = name
                    if file_path:
                        task.files.append({
                            'name': os.path.basename(file_path),
                            'folder': folder,
                        })

                files = analyzer.analyze_folder(input_dir, output_dir, log_callback=_log, progress_callback=_progress)
                task.total = len(files)
                task.progress = len(files)
                task.status = 'completed'
                send_log(f"🎉 Hoàn tất! Đã xuất {len(files)} file Word từ PDF", "success")
            except Exception as e:
                import traceback
                task.status = 'failed'
                task.errors.append(str(e))
                send_log(f"❌ Lỗi phân tích PDF: {e}", "error")
                traceback.print_exc()
            task.end_time = datetime.now()

        t = threading.Thread(target=_run, daemon=True)
        t.start()

        return jsonify({
            'success': True,
            'task_id': task_id,
            'message': f'Đã bắt đầu phân tích {folder}',
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/pdf/face/status/<task_id>')
def pdf_face_status(task_id):
    """Kiểm tra tiến độ phân tích khuôn mặt từ PDF"""
    task = pdf_face_tasks.get(task_id)
    if not task:
        return jsonify({'error': 'Task không tồn tại'}), 404
    return jsonify(task.to_dict())


@app.route('/api/pdf/face/files')
def pdf_face_files():
    """Liệt kê các file Word đã phân tích từ PDF"""
    folders = []
    if os.path.exists(PDF_FACE_OUTPUT_DIR):
        for folder in os.listdir(PDF_FACE_OUTPUT_DIR):
            folder_path = os.path.join(PDF_FACE_OUTPUT_DIR, folder)
            if os.path.isdir(folder_path):
                word_files = [
                    {
                        'name': f,
                        'size': os.path.getsize(os.path.join(folder_path, f)),
                        'modified': datetime.fromtimestamp(
                            os.path.getmtime(os.path.join(folder_path, f))
                        ).isoformat()
                    }
                    for f in os.listdir(folder_path) if f.lower().endswith('.docx')
                ]
                word_files.sort(key=lambda x: x['name'])
                folders.append({
                    'folder': folder,
                    'files': word_files,
                    'count': len(word_files)
                })
    return jsonify({'success': True, 'folders': folders, 'output_dir': PDF_FACE_OUTPUT_DIR})


@app.route('/api/pdf/face/download/<folder>/<filename>')
def pdf_face_download(folder, filename):
    """Tải file Word đã phân tích từ PDF"""
    from urllib.parse import unquote
    folder = unquote(folder)
    filename = unquote(filename)

    if '..' in folder or '..' in filename or '/' in folder or '\\' in folder:
        return jsonify({'error': 'Invalid path'}), 400

    file_path = os.path.join(PDF_FACE_OUTPUT_DIR, folder, filename)

    if not os.path.abspath(file_path).startswith(os.path.abspath(PDF_FACE_OUTPUT_DIR)):
        return jsonify({'error': 'Invalid path'}), 400

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File không tồn tại', 'path': file_path}), 404


# ==================== API: EXCEL EXTRACTION ====================

# Thư mục cho Excel
EXCEL_UPLOAD_DIR = os.path.join(BASE_DIR, "excel_uploads")
EXCEL_PERSON_DIR = os.path.join(BASE_DIR, "excel_persons")
EXCEL_OUTPUT_DIR = os.path.join(BASE_DIR, "excel_extracted")  # Word chi tiet cham cong theo nguoi
EXCEL_FACE_OUTPUT_DIR = os.path.join(BASE_DIR, "excel_face_output")
os.makedirs(EXCEL_UPLOAD_DIR, exist_ok=True)
os.makedirs(EXCEL_PERSON_DIR, exist_ok=True)
os.makedirs(EXCEL_OUTPUT_DIR, exist_ok=True)
os.makedirs(EXCEL_FACE_OUTPUT_DIR, exist_ok=True)

# Task storage cho excel
excel_tasks = {}
excel_face_tasks = {}

class ExcelTask:
    def __init__(self, task_id):
        self.task_id = task_id
        self.status = 'pending'   # pending | running | completed | failed
        self.progress = 0
        self.total = 0
        self.current = ''
        self.files = []
        self.errors = []
        self.start_time = None
        self.end_time = None

    def to_dict(self):
        return {
            'task_id': self.task_id,
            'status': self.status,
            'progress': self.progress,
            'total': self.total,
            'current': self.current,
            'files': self.files,
            'errors': self.errors,
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
        }

class ExcelFaceTask:
    def __init__(self, task_id):
        self.task_id = task_id
        self.status = 'pending'
        self.progress = 0
        self.total = 0
        self.current = ''
        self.files = []
        self.errors = []
        self.start_time = None
        self.end_time = None

    def to_dict(self):
        return {
            'task_id': self.task_id,
            'status': self.status,
            'progress': self.progress,
            'total': self.total,
            'current': self.current,
            'files': self.files,
            'errors': self.errors,
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
        }

@app.route('/api/excel/upload', methods=['POST'])
def excel_upload():
    """Upload file Excel chấm công (.xls/.xlsx)"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Không có file được upload'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'error': 'Tên file không hợp lệ'}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.xls', '.xlsx'):
        return jsonify({'success': False, 'error': 'Chỉ chấp nhận file .xls hoặc .xlsx'}), 400

    # Giữ tên gốc (có tiếng Việt)
    filename = file.filename
    filepath = os.path.join(EXCEL_UPLOAD_DIR, filename)
    file.save(filepath)

    return jsonify({
        'success': True,
        'filename': filename,
        'filepath': filepath,
        'size': os.path.getsize(filepath)
    })

@app.route('/api/excel/extract', methods=['POST'])
def excel_extract():
    """Bat dau tach Excel -> file Excel theo nguoi + file Word de in"""
    try:
        data = request.json or {}
        filename = data.get('filename')
        if not filename:
            return jsonify({'success': False, 'error': 'Thiếu tên file'}), 400

        filepath = os.path.join(EXCEL_UPLOAD_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': 'File Excel không tồn tại'}), 404

        # Tạo thư mục output riêng
        base_name = os.path.splitext(filename)[0]
        person_dir = os.path.join(EXCEL_PERSON_DIR, base_name)
        output_dir = os.path.join(EXCEL_OUTPUT_DIR, base_name)
        os.makedirs(person_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        task_id = f"excel_{int(time.time() * 1000)}"
        task = ExcelTask(task_id)
        excel_tasks[task_id] = task

        def _run():
            task.status = 'running'
            task.start_time = datetime.now()
            try:
                send_log(f"📂 Đang đọc file Excel: {filename}", "info")
                from src.excel_splitter import ExcelAttendanceSplitter
                from src.excel_list_word_exporter import ExcelListWordExporter

                splitter = ExcelAttendanceSplitter(filepath)
                person_files, summaries = splitter.split(person_dir)
                task.total = len(summaries)
                send_log(f"✅ Tìm thấy {len(summaries)} nhân viên trong file", "success")

                exporter = ExcelListWordExporter(output_dir)

                for i, s in enumerate(summaries, 1):
                    task.current = s['name']
                    word_path = exporter.export_from_excel(person_files[i - 1])
                    if not word_path:
                        task.errors.append(f"Khong tao duoc Word cho {s['name']}")
                        send_log(f"⚠️ Không tạo được Word cho {s['name']}", "warning")
                    else:
                        task.files.append({
                            'name': os.path.basename(word_path),
                            'person': s['name'],
                            'days': s['rows'],
                            'present_rows': s['present_rows'],
                            'folder': base_name,
                        })
                    task.progress = i

                task.status = 'completed'
                send_log(
                    f"🎉 Hoàn tất! Đã tạo {len(task.files)} file Word in trong excel_extracted\\{base_name}",
                    "success"
                )
            except Exception as e:
                import traceback
                task.status = 'failed'
                task.errors.append(str(e))
                send_log(f"❌ Lỗi xử lý Excel: {e}", "error")
                traceback.print_exc()
            task.end_time = datetime.now()

        t = threading.Thread(target=_run, daemon=True)
        t.start()

        return jsonify({
            'success': True,
            'task_id': task_id,
            'message': f'Đã bắt đầu xử lý {filename}',
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/excel/status/<task_id>')
def excel_status(task_id):
    """Kiểm tra tiến độ tách Excel"""
    task = excel_tasks.get(task_id)
    if not task:
        return jsonify({'error': 'Task không tồn tại'}), 404
    return jsonify(task.to_dict())

@app.route('/api/excel/files')
def excel_list_files():
    """Liệt kê các file Word chi tiết chấm công đã tạo từ Excel"""
    folders = []
    if os.path.exists(EXCEL_OUTPUT_DIR):
        for folder in os.listdir(EXCEL_OUTPUT_DIR):
            folder_path = os.path.join(EXCEL_OUTPUT_DIR, folder)
            if os.path.isdir(folder_path):
                word_files = [
                    {
                        'name': f,
                        'size': os.path.getsize(os.path.join(folder_path, f)),
                        'modified': datetime.fromtimestamp(
                            os.path.getmtime(os.path.join(folder_path, f))
                        ).isoformat()
                    }
                    for f in os.listdir(folder_path) if f.lower().endswith('.docx')
                ]
                word_files.sort(key=lambda x: x['name'])
                folders.append({
                    'folder': folder,
                    'files': word_files,
                    'count': len(word_files)
                })
    return jsonify({'success': True, 'folders': folders, 'output_dir': EXCEL_OUTPUT_DIR})

@app.route('/api/excel/download/<folder>/<filename>')
def excel_download(folder, filename):
    """Tải file Word chi tiết chấm công theo từng người"""
    from urllib.parse import unquote
    folder = unquote(folder)
    filename = unquote(filename)

    if '..' in folder or '..' in filename or '/' in folder or '\\' in folder:
        return jsonify({'error': 'Invalid path'}), 400

    file_path = os.path.join(EXCEL_OUTPUT_DIR, folder, filename)

    if not os.path.abspath(file_path).startswith(os.path.abspath(EXCEL_OUTPUT_DIR)):
        return jsonify({'error': 'Invalid path'}), 400

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File không tồn tại', 'path': file_path}), 404


@app.route('/api/excel/face/analyze', methods=['POST'])
def excel_face_analyze():
    """Phân tích khuôn mặt từ các file Excel đã tách"""
    try:
        data = request.json or {}
        folder = data.get('folder')
        threshold = data.get('distance_threshold')
        try:
            threshold = float(threshold) if threshold not in (None, '') else None
        except Exception:
            threshold = None
        if not folder:
            return jsonify({'success': False, 'error': 'Thiếu tên thư mục'}), 400

        input_dir = os.path.join(EXCEL_PERSON_DIR, folder)
        if not os.path.exists(input_dir):
            return jsonify({'success': False, 'error': 'Thư mục Excel đã tách không tồn tại'}), 404

        output_dir = os.path.join(EXCEL_FACE_OUTPUT_DIR, folder)
        os.makedirs(output_dir, exist_ok=True)

        task_id = f"excel_face_{int(time.time() * 1000)}"
        task = ExcelFaceTask(task_id)
        excel_face_tasks[task_id] = task

        project_name = data.get('project') or DEFAULT_PROJECT_NAME
        p_dir = os.path.join(PORTRAIT_DIR, project_name) if os.path.exists(os.path.join(PORTRAIT_DIR, project_name)) else PORTRAIT_DIR
        i_dir = os.path.join(INPUT_IMAGES_DIR, project_name) if os.path.exists(os.path.join(INPUT_IMAGES_DIR, project_name)) else INPUT_IMAGES_DIR

        def _run():
            task.status = 'running'
            task.start_time = datetime.now()
            try:
                send_log(f" Bắt đầu phân tích khuôn mặt cho thư mục: {folder} (Dự án: {project_name})", "info")
                matcher = get_face_matcher()
                if matcher:
                    send_log("✅ Face Matcher đã sẵn sàng", "success")
                else:
                    send_log("⚠️ Face Matcher không khả dụng, sẽ bỏ qua tìm ảnh camera", "warning")

                from src.excel_face_analyzer import ExcelFaceAnalyzer
                analyzer = ExcelFaceAnalyzer(
                    p_dir,
                    i_dir,
                    matcher,
                    accuracy_mode=True,
                    match_distance_threshold=threshold,
                    log_detail=True
                )

                def _log(msg, t='default'):
                    send_log(msg, t)

                def _progress(completed, total, name, file_path):
                    task.total = total
                    task.progress = completed
                    task.current = name
                    if file_path:
                        task.files.append({
                            'name': os.path.basename(file_path),
                            'folder': folder,
                        })

                files = analyzer.analyze_folder(input_dir, output_dir, log_callback=_log, progress_callback=_progress)
                task.total = len(files)
                task.progress = len(files)
                task.status = 'completed'
                send_log(f"🎉 Hoàn tất! Đã xuất {len(files)} file Word", "success")
            except Exception as e:
                import traceback
                task.status = 'failed'
                task.errors.append(str(e))
                send_log(f"❌ Lỗi phân tích Excel: {e}", "error")
                traceback.print_exc()
            task.end_time = datetime.now()

        t = threading.Thread(target=_run, daemon=True)
        t.start()

        return jsonify({
            'success': True,
            'task_id': task_id,
            'message': f'Đã bắt đầu phân tích {folder}',
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/excel/face/status/<task_id>')
def excel_face_status(task_id):
    """Kiểm tra tiến độ phân tích khuôn mặt từ Excel"""
    task = excel_face_tasks.get(task_id)
    if not task:
        return jsonify({'error': 'Task không tồn tại'}), 404
    return jsonify(task.to_dict())


@app.route('/api/excel/face/files')
def excel_face_files():
    """Liệt kê các file Word đã phân tích từ Excel"""
    folders = []
    if os.path.exists(EXCEL_FACE_OUTPUT_DIR):
        for folder in os.listdir(EXCEL_FACE_OUTPUT_DIR):
            folder_path = os.path.join(EXCEL_FACE_OUTPUT_DIR, folder)
            if os.path.isdir(folder_path):
                word_files = [
                    {
                        'name': f,
                        'size': os.path.getsize(os.path.join(folder_path, f)),
                        'modified': datetime.fromtimestamp(
                            os.path.getmtime(os.path.join(folder_path, f))
                        ).isoformat()
                    }
                    for f in os.listdir(folder_path) if f.lower().endswith('.docx')
                ]
                word_files.sort(key=lambda x: x['name'])
                folders.append({
                    'folder': folder,
                    'files': word_files,
                    'count': len(word_files)
                })
    return jsonify({'success': True, 'folders': folders, 'output_dir': EXCEL_FACE_OUTPUT_DIR})


@app.route('/api/excel/face/download/<folder>/<filename>')
def excel_face_download(folder, filename):
    """Tải file Word đã phân tích từ Excel"""
    from urllib.parse import unquote
    folder = unquote(folder)
    filename = unquote(filename)

    if '..' in folder or '..' in filename or '/' in folder or '\\' in folder:
        return jsonify({'error': 'Invalid path'}), 400

    file_path = os.path.join(EXCEL_FACE_OUTPUT_DIR, folder, filename)

    if not os.path.abspath(file_path).startswith(os.path.abspath(EXCEL_FACE_OUTPUT_DIR)):
        return jsonify({'error': 'Invalid path'}), 400

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File không tồn tại', 'path': file_path}), 404

@app.route('/api/excel/uploads')
def excel_list_uploads():
    """Liệt kê các file Excel đã upload"""
    files = []
    if os.path.exists(EXCEL_UPLOAD_DIR):
        for f in os.listdir(EXCEL_UPLOAD_DIR):
            ext = os.path.splitext(f)[1].lower()
            if ext in ('.xls', '.xlsx'):
                fp = os.path.join(EXCEL_UPLOAD_DIR, f)
                files.append({
                    'name': f,
                    'size': os.path.getsize(fp),
                    'modified': datetime.fromtimestamp(os.path.getmtime(fp)).isoformat()
                })
    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify({'success': True, 'files': files, 'upload_dir': EXCEL_UPLOAD_DIR})


# ==================== DELETE APIS ====================

@app.route('/api/excel/delete-upload', methods=['POST'])
def excel_delete_upload():
    """Xóa file Excel đã tải lên"""
    try:
        data = request.get_json(silent=True) or {}
        filename = data.get('filename', '')
        if not filename:
            return jsonify({'success': False, 'error': 'Tên file không hợp lệ'}), 400
        safe_name = os.path.basename(filename)
        file_path = os.path.join(EXCEL_UPLOAD_DIR, safe_name)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_name}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/excel/delete-folder', methods=['POST'])
def excel_delete_folder():
    """Xóa cả thư mục đợt tách Excel"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        if not folder:
            return jsonify({'success': False, 'error': 'Tên thư mục không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        folder_path = os.path.join(EXCEL_OUTPUT_DIR, safe_folder)
        deleted = False
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            shutil.rmtree(folder_path)
            deleted = True
        person_path = os.path.join(EXCEL_PERSON_DIR, safe_folder)
        if os.path.exists(person_path) and os.path.isdir(person_path):
            shutil.rmtree(person_path)
            deleted = True
        if deleted:
            return jsonify({'success': True, 'message': f'Đã xóa đợt {safe_folder}'})
        return jsonify({'success': False, 'error': 'Thư mục không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/excel/delete-file', methods=['POST'])
def excel_delete_file():
    """Xóa 1 file Word chi tiết trong đợt tách Excel"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        filename = data.get('filename', '')
        if not folder or not filename:
            return jsonify({'success': False, 'error': 'Thông tin file không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        safe_file = os.path.basename(filename)
        file_path = os.path.join(EXCEL_OUTPUT_DIR, safe_folder, safe_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_file}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/excel/face/delete-folder', methods=['POST'])
def excel_face_delete_folder():
    """Xóa thư mục kết quả quét mặt Excel"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        if not folder:
            return jsonify({'success': False, 'error': 'Tên thư mục không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        folder_path = os.path.join(EXCEL_FACE_OUTPUT_DIR, safe_folder)
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            shutil.rmtree(folder_path)
            return jsonify({'success': True, 'message': f'Đã xóa kết quả {safe_folder}'})
        return jsonify({'success': False, 'error': 'Thư mục không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/excel/face/delete-file', methods=['POST'])
def excel_face_delete_file():
    """Xóa 1 file kết quả quét mặt Excel"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        filename = data.get('filename', '')
        if not folder or not filename:
            return jsonify({'success': False, 'error': 'Thông tin không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        safe_file = os.path.basename(filename)
        file_path = os.path.join(EXCEL_FACE_OUTPUT_DIR, safe_folder, safe_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_file}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pdf/delete-upload', methods=['POST'])
def pdf_delete_upload():
    """Xóa file PDF đã tải lên"""
    try:
        data = request.get_json(silent=True) or {}
        filename = data.get('filename', '')
        if not filename:
            return jsonify({'success': False, 'error': 'Tên file không hợp lệ'}), 400
        safe_name = os.path.basename(filename)
        file_path = os.path.join(PDF_UPLOAD_DIR, safe_name)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_name}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pdf/delete-folder', methods=['POST'])
def pdf_delete_folder():
    """Xóa cả thư mục đợt tách PDF"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        if not folder:
            return jsonify({'success': False, 'error': 'Tên thư mục không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        folder_path = os.path.join(PDF_OUTPUT_DIR, safe_folder)
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            shutil.rmtree(folder_path)
            return jsonify({'success': True, 'message': f'Đã xóa đợt {safe_folder}'})
        return jsonify({'success': False, 'error': 'Thư mục không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pdf/delete-file', methods=['POST'])
def pdf_delete_file():
    """Xóa 1 file trong thư mục đợt tách PDF"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        filename = data.get('filename', '')
        if not folder or not filename:
            return jsonify({'success': False, 'error': 'Thông tin không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        safe_file = os.path.basename(filename)
        file_path = os.path.join(PDF_OUTPUT_DIR, safe_folder, safe_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_file}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pdf/face/delete-folder', methods=['POST'])
def pdf_face_delete_folder():
    """Xóa thư mục kết quả quét mặt PDF"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        if not folder:
            return jsonify({'success': False, 'error': 'Tên thư mục không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        folder_path = os.path.join(PDF_FACE_OUTPUT_DIR, safe_folder)
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            shutil.rmtree(folder_path)
            return jsonify({'success': True, 'message': f'Đã xóa kết quả {safe_folder}'})
        return jsonify({'success': False, 'error': 'Thư mục không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/pdf/face/delete-file', methods=['POST'])
def pdf_face_delete_file():
    """Xóa 1 file kết quả quét mặt PDF"""
    try:
        data = request.get_json(silent=True) or {}
        folder = data.get('folder', '')
        filename = data.get('filename', '')
        if not folder or not filename:
            return jsonify({'success': False, 'error': 'Thông tin không hợp lệ'}), 400
        safe_folder = os.path.basename(folder)
        safe_file = os.path.basename(filename)
        file_path = os.path.join(PDF_FACE_OUTPUT_DIR, safe_folder, safe_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_file}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/results/delete', methods=['POST'])
def results_delete_file():
    """Xóa file trong thư mục results"""
    try:
        data = request.get_json(silent=True) or {}
        filename = data.get('filename', '')
        if not filename:
            return jsonify({'success': False, 'error': 'Tên file không hợp lệ'}), 400
        safe_file = os.path.basename(filename)
        file_path = os.path.join(RESULTS_DIR, safe_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': f'Đã xóa {safe_file}'})
        return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ==================== ZALO SERVICE PROXY ====================

ZALO_SERVICE_URL = "http://127.0.0.1:3001"

def _proxy_zalo(path, method='GET', data=None, retry=True):
    """Proxy request to Node.js Zalo service với cơ chế tự phục hồi"""
    import urllib.request
    import urllib.error
    url = f"{ZALO_SERVICE_URL}{path}"
    try:
        if data is not None:
            req = urllib.request.Request(
                url,
                data=json.dumps(data).encode('utf-8'),
                headers={'Content-Type': 'application/json'},
                method=method
            )
        else:
            req = urllib.request.Request(url, method=method)
        
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode('utf-8')
            return json.loads(body), resp.status
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8')
        try:
            return json.loads(body), e.code
        except Exception:
            return {'success': False, 'error': body}, e.code
    except urllib.error.URLError as e:
        # Nếu service chưa khởi động hoặc bị crash, tự phục hồi và thử lại 1 lần
        if retry:
            try:
                from src.zalo_service_manager import ensure_zalo_service
                if ensure_zalo_service(base_dir=BASE_DIR):
                    return _proxy_zalo(path, method=method, data=data, retry=False)
            except Exception as ex:
                logging.error(f"[App] Lỗi khi thử tự khởi động Zalo Service: {ex}")
        return {'success': False, 'error': f'Zalo Service không hoạt động: {str(e)}'}, 503
    except Exception as e:
        return {'success': False, 'error': str(e)}, 500


@app.route('/api/zalo/status')
def zalo_status():
    """Kiểm tra trạng thái kết nối Zalo"""
    result, status = _proxy_zalo('/api/status')
    return jsonify(result), status


@app.route('/api/zalo/login/qr', methods=['POST'])
def zalo_login_qr():
    """Bắt đầu đăng nhập bằng QR code"""
    data = request.get_json(silent=True) or {}
    result, status = _proxy_zalo('/api/login/qr', method='POST', data=data)
    return jsonify(result), status


@app.route('/api/zalo/login/qr/image')
def zalo_login_qr_image():
    """Lấy hình ảnh QR code"""
    result, status = _proxy_zalo('/api/login/qr/image')
    return jsonify(result), status


@app.route('/api/zalo/login/qr/raw')
def zalo_login_qr_raw():
    """Lấy trực tiếp file ảnh QR code (binary PNG)"""
    import urllib.request
    import urllib.error
    url = f"{ZALO_SERVICE_URL}/api/login/qr/raw"
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = resp.read()
            return Response(data, mimetype='image/png')
    except Exception as e:
        return jsonify({'error': str(e)}), 404


@app.route('/api/zalo/login/qr/status')
def zalo_login_qr_status():
    """Kiểm tra trạng thái quét QR"""
    result, status = _proxy_zalo('/api/login/qr/status')
    return jsonify(result), status


@app.route('/api/zalo/logout', methods=['POST'])
def zalo_logout():
    """Đăng xuất Zalo"""
    result, status = _proxy_zalo('/api/logout', method='POST')
    return jsonify(result), status


@app.route('/api/zalo/groups')
def zalo_groups():
    """Lấy danh sách group Zalo"""
    result, status = _proxy_zalo('/api/groups')
    return jsonify(result), status


@app.route('/api/zalo/groups/<group_id>/download', methods=['POST'])
def zalo_download(group_id):
    """Tải ảnh từ group Zalo"""
    data = request.get_json(silent=True) or {}
    result, status = _proxy_zalo(f'/api/groups/{group_id}/download', method='POST', data=data)
    return jsonify(result), status


@app.route('/api/zalo/download/progress')
def zalo_download_progress():
    """SSE endpoint theo dõi tiến trình tải ảnh"""
    import urllib.request
    import urllib.error
    
    def generate():
        try:
            req = urllib.request.Request(f"{ZALO_SERVICE_URL}/api/download/progress")
            with urllib.request.urlopen(req, timeout=300) as resp:
                while True:
                    line = resp.readline()
                    if not line:
                        break
                    yield line.decode('utf-8')
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
    
    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'Connection': 'keep-alive'})


@app.route('/api/zalo/download/progress/poll')
def zalo_download_progress_poll():
    """Polling endpoint cho tiến trình tải ảnh"""
    result, status = _proxy_zalo('/api/download/progress/poll')
    return jsonify(result), status


@app.route('/api/zalo/health')
def zalo_health():
    """Kiểm tra Zalo service có đang chạy không"""
    result, status = _proxy_zalo('/api/health')
    return jsonify(result), status


@app.route('/api/zalo/restart', methods=['POST'])
def zalo_restart():
    """Khởi động lại Zalo Service theo yêu cầu"""
    try:
        from src.zalo_service_manager import restart_zalo_service
        success = restart_zalo_service(base_dir=BASE_DIR)
        if success:
            return jsonify({'success': True, 'message': 'Zalo Service đã được khởi động lại thành công'})
        else:
            return jsonify({'success': False, 'error': 'Không thể khởi động lại Zalo Service. Vui lòng kiểm tra Node.js.'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/open/folder', methods=['POST'])
def open_system_folder():
    """Mở thư mục trong File Explorer (Windows)"""
    data = request.json or {}
    target_type = data.get('type', 'input_images')
    subpath = data.get('subpath', '').strip()
    
    if target_type == 'input_images':
        folder = os.path.join(INPUT_IMAGES_DIR, subpath) if subpath else INPUT_IMAGES_DIR
    elif target_type == 'results':
        folder = os.path.join(RESULTS_DIR, subpath) if subpath else RESULTS_DIR
    else:
        folder = BASE_DIR
        
    os.makedirs(folder, exist_ok=True)
    try:
        if sys.platform == 'win32':
            os.startfile(folder)
        else:
            import subprocess
            subprocess.Popen(['xdg-open', folder])
        return jsonify({'success': True, 'path': folder})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==================== MAIN ====================


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass
    print("=" * 60)
    print(f"Server dang chay tai: http://localhost:{FLASK_PORT}")
    print(f"Input Images: {INPUT_IMAGES_DIR}")
    print(f"Results:      {RESULTS_DIR}")
    print("=" * 60)

    try:
        scan_database()
    except Exception:
        pass

    app.run(host=FLASK_HOST, port=FLASK_PORT, debug=FLASK_DEBUG, threaded=True)
