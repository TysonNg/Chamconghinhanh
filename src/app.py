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

# Thêm thư mục gốc vào path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, Response
from werkzeug.utils import secure_filename

# Cấu hình - Phát hiện đúng thư mục khi chạy từ EXE
def get_base_dir():
    """Lấy thư mục gốc - hỗ trợ cả khi chạy từ source và từ EXE"""
    if getattr(sys, 'frozen', False):
        # Chạy từ EXE (PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # Chạy từ source
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = get_base_dir()
INPUT_IMAGES_DIR = os.path.join(BASE_DIR, "input_images")
DATABASE_DIR = os.path.join(BASE_DIR, "database")
RESULTS_DIR = os.path.join(BASE_DIR, "results")
SUPPORTED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

FLASK_HOST = "0.0.0.0"
FLASK_PORT = 5000
FLASK_DEBUG = True
MAX_WORKERS = 4

# Tạo thư mục nếu chưa tồn tại
for directory in [INPUT_IMAGES_DIR, DATABASE_DIR, RESULTS_DIR]:
    os.makedirs(directory, exist_ok=True)

app = Flask(__name__, 
            template_folder='../templates',
            static_folder='../static')

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

def send_log(message, log_type='default'):
    """Send log message to all connected SSE clients"""
    log_data = json.dumps({'message': message, 'type': log_type})
    with log_lock:
        for q in log_queues:
            try:
                q.put_nowait(log_data)
            except:
                pass
    # Also print to console
    print(message)


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

# ==================== API: ATTENDANCE ====================

# Đường dẫn cho attendance
CHAMCONG_DIR = os.path.join(BASE_DIR, "chamcong")
PORTRAIT_DIR = os.path.join(BASE_DIR, "Ảnh BV")

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
    if _face_matcher is None:
        try:
            send_log("⏳ Đang khởi tạo Face Matcher (DeepFace)...", "info")
            from src.face_matcher import FaceMatcher
            _face_matcher = FaceMatcher(PORTRAIT_DIR, log_callback=send_log)
            send_log("✅ Face Matcher đã sẵn sàng!", "success")
        except Exception as e:
            send_log(f"❌ Lỗi khởi tạo FaceMatcher: {e}", "error")
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
        send_log("📂 Step 1: Đang phân tích file chấm công...", "info")
        processor = AttendanceProcessor(CHAMCONG_DIR)
        processor.scan_all_files()
        missing_records = processor.get_missing_records()
        summary = processor.get_summary()
        send_log(f"📋 Tìm thấy {len(missing_records)} bản ghi thiếu từ {summary.get('total_persons', 0)} người", "info")
        
        # Step 2: Khởi tạo face matcher
        send_log("🔧 Step 2: Đang khởi tạo Face Matcher...", "info")
        matcher = get_face_matcher()
        if matcher:
            send_log("✅ Face Matcher đã sẵn sàng", "success")
        else:
            send_log("⚠️ Face Matcher không khả dụng, sẽ dùng fallback", "warning")
        
        # Step 3: Match ảnh camera cho mỗi bản ghi thiếu
        send_log(f"🔍 Step 3: Bắt đầu matching ảnh cho {len(missing_records)} bản ghi...", "info")
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
                    matched_image = matcher.match_face_in_images(person_name, images)
                    if matched_image:
                        record['matched_image'] = matched_image
                        matched_count += 1
                        send_log(f"  [{i+1}/{len(missing_records)}] ✓ {person_name} -> {os.path.basename(matched_image)}", "success")
                    else:
                        send_log(f"  [{i+1}/{len(missing_records)}] ✗ {person_name}: Không tìm thấy ảnh match", "warning")
                elif images:
                    # Fallback: không có matcher, lấy ảnh đầu tiên
                    send_log(f"  [{i+1}/{len(missing_records)}] ⚠️ FaceMatcher chưa sẵn sàng, lấy ảnh đầu tiên cho {person_name}", "warning")
                    record['matched_image'] = images[0]
                    matched_count += 1
                else:
                    send_log(f"  [{i+1}/{len(missing_records)}] ⚠️ Thư mục {day} rỗng", "warning")
            else:
                send_log(f"  [{i+1}/{len(missing_records)}] ❌ Không tìm thấy thư mục: {day_folder}", "error")
                record['matched_image'] = None
        
        summary['total_matched'] = matched_count
        send_log(f"🎉 Hoàn thành! Matched {matched_count}/{len(missing_records)} bản ghi", "success")
        
        return jsonify({
            'success': True,
            'summary': summary,
            'records': missing_records
        })
    except Exception as e:
        import traceback
        send_log(f"❌ Lỗi: {e}", "error")
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
os.makedirs(PDF_UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)

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

# ==================== MAIN ====================

if __name__ == '__main__':
    print(f"""
╔══════════════════════════════════════════════════════════════╗
║     PHẦN MỀM NHẬN DIỆN KHUÔN MẶT & TRÍCH XUẤT NGÀY THÁNG    ║
╠══════════════════════════════════════════════════════════════╣
║  Server đang chạy tại: http://localhost:{FLASK_PORT}                   ║
║  Input Images: {INPUT_IMAGES_DIR:<44} ║
║  Database:     {DATABASE_DIR:<44} ║
║  Results:      {RESULTS_DIR:<44} ║
╚══════════════════════════════════════════════════════════════╝
    """)
    
    print("Đang quét database...")
    scan_database()
    print(f"Đã load {len(database)} người trong database\n")
    
    app.run(host=FLASK_HOST, port=FLASK_PORT, debug=FLASK_DEBUG, threaded=True)
