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

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, Response, url_for
from werkzeug.utils import secure_filename

# Cấu hình - Phát hiện đúng thư mục khi chạy từ EXE
def get_base_dir():
    """Lấy thư mục gốc (chứa data: input_images, database, ...) - hỗ trợ cả khi chạy từ source và từ EXE"""
    if os.environ.get("ATTENDANCE_DATA_DIR"):
        return os.path.abspath(os.environ["ATTENDANCE_DATA_DIR"])
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


def _build_report_options(data, fallback_project_name):
    """Chuẩn hóa thông tin báo cáo tổng hợp nhận từ giao diện/API."""
    project_name = str(data.get('report_project_name') or fallback_project_name or '').strip()
    if not project_name:
        raise ValueError('Thiếu tên dự án cho báo cáo tổng hợp')

    from_date = str(data.get('from_date') or '').strip()
    to_date = str(data.get('to_date') or '').strip()
    if bool(from_date) != bool(to_date):
        raise ValueError('Vui lòng chọn đầy đủ Từ ngày và Đến ngày')

    if from_date and to_date:
        try:
            start = datetime.strptime(from_date, '%Y-%m-%d').date()
            end = datetime.strptime(to_date, '%Y-%m-%d').date()
        except ValueError as exc:
            raise ValueError('Ngày báo cáo phải có định dạng YYYY-MM-DD') from exc
        if start > end:
            raise ValueError('Từ ngày không được sau Đến ngày')

    options = {'project_name': project_name}
    if from_date:
        options['from_date'] = from_date
        options['to_date'] = to_date
    return options

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
SUPPLEMENT_DIR = os.path.join(BASE_DIR, "supplement_data")

# Tao thu muc neu chua ton tai
for directory in [INPUT_IMAGES_DIR, RESULTS_DIR, CHAMCONG_DIR, PORTRAIT_DIR, NGAY_RONG_DIR, SUPPLEMENT_DIR]:
    os.makedirs(directory, exist_ok=True)

# Lazy-loaded PhotoSupplement instance
_photo_supplement = None
def get_photo_supplement():
    global _photo_supplement
    if _photo_supplement is None:
        from src.photo_supplement import PhotoSupplement
        _photo_supplement = PhotoSupplement(SUPPLEMENT_DIR)
    return _photo_supplement

DEFAULT_PROJECT_NAME = "Chung cư Tân Thuận Đông"
_identity_registry = None
_identity_registry_lock = threading.Lock()

def get_identity_registry():
    global _identity_registry
    from pathlib import Path
    from src.identity_registry import IdentityRegistry
    with _identity_registry_lock:
        wanted = Path(BASE_DIR) / "data" / "identity.sqlite3"
        if (_identity_registry is None or _identity_registry.db_path != wanted
                or _identity_registry.portrait_root != Path(PORTRAIT_DIR).resolve()):
            _identity_registry = IdentityRegistry(wanted, PORTRAIT_DIR)
        return _identity_registry



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

# Legacy migration is explicit; never move source photos on startup.

def ensure_project_structure(project_name: str, create_day_folders: bool = False) -> Tuple[str, str]:
    """Đảm bảo đầy đủ cấu trúc thư mục cho dự án:
    1. Ảnh BV/<project_name> (Thư mục chân dung nhân viên)
    2. input_images/<project_name> (Thư mục ảnh camera theo ngày)
    Thư mục ngày YYYY-MM-DD chỉ được tạo khi có ngày đầy đủ.
    """
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', project_name.strip()) if project_name else DEFAULT_PROJECT_NAME
    p_dir = os.path.join(PORTRAIT_DIR, safe_name)
    i_dir = os.path.join(INPUT_IMAGES_DIR, safe_name)
    os.makedirs(p_dir, exist_ok=True)
    os.makedirs(i_dir, exist_ok=True)
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
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True

from src.supplement_batches import register_batches
register_batches(app, SUPPLEMENT_DIR)

from src.task_manager import TaskManager
task_manager = TaskManager(BASE_DIR)

# Global state
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

from src.identity_routes import register_identity_routes
from src.daily_photo_routes import register_daily_photo_routes
register_identity_routes(app, get_identity_registry, lambda: INPUT_IMAGES_DIR)
register_daily_photo_routes(app, lambda: INPUT_IMAGES_DIR,
                           project_resolver=lambda ref: get_identity_registry().get_project(ref))

# ==================== API: FACE MATCHER ====================

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
            _face_matcher = FaceMatcher(portrait_dir, log_callback=send_log, identity_registry=get_identity_registry())
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
                _face_matcher = FaceMatcher(alt_dir, log_callback=send_log, identity_registry=get_identity_registry())
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
        self.aggregate_report = None
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
            'aggregate_report': self.aggregate_report,
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

        try:
            identity_registry = get_identity_registry()
            selected_project = identity_registry.get_project(data.get('project_id') or data.get('project') or DEFAULT_PROJECT_NAME)
            if not selected_project['active']:
                raise ValueError('Dự án đã được lưu trữ')
            project_name = selected_project['storage_dir']
            project_id = selected_project['project_id']
            report_options = _build_report_options(data, selected_project['display_name'])
        except ValueError as exc:
            return jsonify({'success': False, 'error': str(exc)}), 400
        p_dir = str(identity_registry.project_portrait_dir(project_id))
        i_dir = os.path.join(INPUT_IMAGES_DIR, project_name)

        task_id = f"pdf_face_{int(time.time() * 1000)}"

        payload = {
            'folder': folder,
            'project_name': project_name,
            'project_id': project_id,
            'threshold': threshold,
            'report_options': report_options,
            'input_dir': input_dir,
            'output_dir': output_dir,
            'p_dir': p_dir,
            'i_dir': i_dir,
        }

        def _run(task_rec, cancel_check):
            try:
                send_log(f"🚀 Bắt đầu phân tích khuôn mặt cho thư mục PDF: {folder} (Dự án: {project_name})", "info")
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
                    log_detail=True,
                    project_id=project_id,
                    identity_registry=identity_registry
                )

                def _log(msg, t='default'):
                    send_log(msg, t)

                def _progress(completed, total, name, file_path):
                    task_manager.update_progress(task_rec.task_id, completed, total, name, file_path)

                files = analyzer.analyze_folder(
                    input_dir,
                    output_dir,
                    log_callback=_log,
                    progress_callback=_progress,
                    report_options=report_options,
                    cancel_check=cancel_check,
                )
                aggregate_path = next(
                    (path for path in files if os.path.basename(path).startswith('GIAI_TRINH_')),
                    None,
                )
                if aggregate_path:
                    task_rec.aggregate_report = os.path.basename(aggregate_path)
                    task_rec.files.append({
                        'name': os.path.basename(aggregate_path),
                        'folder': folder,
                        'is_aggregate': True,
                    })
                task_rec.total = len(files)
                task_rec.progress = len(files)
                if cancel_check():
                    send_log(f"⚠️ Đã dừng quét thư mục PDF: {folder}", "warning")
                else:
                    send_log(f"🎉 Hoàn tất! Đã xuất {len(files)} file Word từ PDF", "success")
            except Exception as e:
                import traceback
                task_rec.status = 'failed'
                task_rec.errors.append(str(e))
                send_log(f"❌ Lỗi phân tích PDF: {e}", "error")
                traceback.print_exc()

        task = task_manager.submit_task(
            task_type='pdf_face',
            target_folder=folder,
            project_name=project_name,
            execute_fn=_run,
            payload=payload,
            task_id=task_id,
        )
        pdf_face_tasks[task_id] = task

        return jsonify({
            'success': True,
            'task_id': task_id,
            'status': task.status,
            'message': f'Đã bắt đầu phân tích {folder}' if task.status == 'running' else f'Đã xếp hàng chờ phân tích {folder}',
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/pdf/face/status/<task_id>')
def pdf_face_status(task_id):
    """Kiểm tra tiến độ phân tích khuôn mặt từ PDF"""
    task = task_manager.get_task(task_id) or pdf_face_tasks.get(task_id)
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
        self.aggregate_report = None
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
            'aggregate_report': self.aggregate_report,
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

        try:
            identity_registry = get_identity_registry()
            selected_project = identity_registry.get_project(data.get('project_id') or data.get('project') or DEFAULT_PROJECT_NAME)
            if not selected_project['active']:
                raise ValueError('Dự án đã được lưu trữ')
            project_name = selected_project['storage_dir']
            project_id = selected_project['project_id']
            report_options = _build_report_options(data, selected_project['display_name'])
        except ValueError as exc:
            return jsonify({'success': False, 'error': str(exc)}), 400
        p_dir = str(identity_registry.project_portrait_dir(project_id))
        i_dir = os.path.join(INPUT_IMAGES_DIR, project_name)

        task_id = f"excel_face_{int(time.time() * 1000)}"

        payload = {
            'folder': folder,
            'project_name': project_name,
            'project_id': project_id,
            'threshold': threshold,
            'report_options': report_options,
            'input_dir': input_dir,
            'output_dir': output_dir,
            'p_dir': p_dir,
            'i_dir': i_dir,
        }

        def _run(task_rec, cancel_check):
            try:
                send_log(f"🚀 Bắt đầu phân tích khuôn mặt cho thư mục: {folder} (Dự án: {project_name})", "info")
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
                    log_detail=True,
                    project_id=project_id,
                    identity_registry=identity_registry
                )

                def _log(msg, t='default'):
                    send_log(msg, t)

                def _progress(completed, total, name, file_path):
                    task_manager.update_progress(task_rec.task_id, completed, total, name, file_path)

                files = analyzer.analyze_folder(
                    input_dir,
                    output_dir,
                    log_callback=_log,
                    progress_callback=_progress,
                    report_options=report_options,
                    cancel_check=cancel_check,
                )
                aggregate_path = next(
                    (path for path in files if os.path.basename(path).startswith('GIAI_TRINH_')),
                    None,
                )
                if aggregate_path:
                    task_rec.aggregate_report = os.path.basename(aggregate_path)
                    task_rec.files.append({
                        'name': os.path.basename(aggregate_path),
                        'folder': folder,
                        'is_aggregate': True,
                    })
                task_rec.total = len(files)
                task_rec.progress = len(files)
                if cancel_check():
                    send_log(f"⚠️ Đã dừng quét thư mục: {folder}", "warning")
                else:
                    send_log(f"🎉 Hoàn tất! Đã xuất {len(files)} file Word", "success")
            except Exception as e:
                import traceback
                task_rec.status = 'failed'
                task_rec.errors.append(str(e))
                send_log(f"❌ Lỗi phân tích Excel: {e}", "error")
                traceback.print_exc()

        task = task_manager.submit_task(
            task_type='excel_face',
            target_folder=folder,
            project_name=project_name,
            execute_fn=_run,
            payload=payload,
            task_id=task_id,
        )
        excel_face_tasks[task_id] = task

        return jsonify({
            'success': True,
            'task_id': task_id,
            'status': task.status,
            'message': f'Đã bắt đầu phân tích {folder}' if task.status == 'running' else f'Đã xếp hàng chờ phân tích {folder}',
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/excel/face/status/<task_id>')
def excel_face_status(task_id):
    """Kiểm tra tiến độ phân tích khuôn mặt từ Excel"""
    task = task_manager.get_task(task_id) or excel_face_tasks.get(task_id)
    if not task:
        return jsonify({'error': 'Task không tồn tại'}), 404
    return jsonify(task.to_dict())


@app.route('/api/task/cancel/<task_id>', methods=['POST'])
def cancel_task(task_id):
    """Dừng/hủy tác vụ đang chạy hoặc đang chờ"""
    success = task_manager.cancel_task(task_id)
    task = task_manager.get_task(task_id)
    if not task:
        return jsonify({'success': False, 'error': 'Task không tồn tại'}), 404
    send_log(f"🛑 Đã gửi lệnh hủy tác vụ {task_id} ({task.target_folder})", "warning")
    return jsonify({
        'success': True,
        'task_id': task_id,
        'status': task.status,
        'message': f'Đã gửi yêu cầu dừng tác vụ {task_id}'
    })


@app.route('/api/task/retry/<task_id>', methods=['POST'])
def retry_task(task_id):
    """Quét lại tác vụ bị gián đoạn, lỗi hoặc đã hủy"""
    old_task = task_manager.get_task(task_id)
    if not old_task:
        return jsonify({'success': False, 'error': 'Task không tồn tại'}), 404

    payload = old_task.payload or {}
    if not payload:
        return jsonify({'success': False, 'error': 'Không tìm thấy tham số gốc của tác vụ'}), 400

    folder = payload.get('folder', '')
    project_name = payload.get('project_name', '')
    project_id = payload.get('project_id', '')
    threshold = payload.get('threshold')
    report_options = payload.get('report_options')
    input_dir = payload.get('input_dir', '')
    output_dir = payload.get('output_dir', '')
    p_dir = payload.get('p_dir', '')
    i_dir = payload.get('i_dir', '')
    identity_reg = get_identity_registry()

    if old_task.task_type == 'excel_face':
        new_tid = f"excel_face_{int(time.time() * 1000)}"

        def _retry_excel_run(task_rec, cancel_check):
            try:
                send_log(f"🚀 [Quét lại] Phân tích khuôn mặt thư mục: {folder} (Dự án: {project_name})", "info")
                matcher = get_face_matcher()
                from src.excel_face_analyzer import ExcelFaceAnalyzer
                analyzer = ExcelFaceAnalyzer(
                    p_dir, i_dir, matcher, accuracy_mode=True,
                    match_distance_threshold=threshold, log_detail=True,
                    project_id=project_id, identity_registry=identity_reg
                )
                def _log(msg, t='default'):
                    send_log(msg, t)
                def _progress(completed, total, name, file_path):
                    task_manager.update_progress(task_rec.task_id, completed, total, name, file_path)

                files = analyzer.analyze_folder(
                    input_dir, output_dir, log_callback=_log,
                    progress_callback=_progress, report_options=report_options,
                    cancel_check=cancel_check,
                )
                aggregate_path = next(
                    (path for path in files if os.path.basename(path).startswith('GIAI_TRINH_')),
                    None,
                )
                if aggregate_path:
                    task_rec.aggregate_report = os.path.basename(aggregate_path)
                    task_rec.files.append({
                        'name': os.path.basename(aggregate_path),
                        'folder': folder,
                        'is_aggregate': True,
                    })
                task_rec.total = len(files)
                task_rec.progress = len(files)
                if cancel_check():
                    send_log(f"⚠️ Đã dừng quét lại thư mục: {folder}", "warning")
                else:
                    send_log(f"🎉 Hoàn tất quét lại! Đã xuất {len(files)} file Word", "success")
            except Exception as e:
                import traceback
                task_rec.status = 'failed'
                task_rec.errors.append(str(e))
                send_log(f"❌ Lỗi quét lại Excel: {e}", "error")
                traceback.print_exc()

        new_task = task_manager.submit_task(
            task_type='excel_face',
            target_folder=folder,
            project_name=project_name,
            execute_fn=_retry_excel_run,
            payload=payload,
            task_id=new_tid,
        )
        excel_face_tasks[new_tid] = new_task
        return jsonify({
            'success': True,
            'task_id': new_tid,
            'status': new_task.status,
            'message': f'Đã xếp hàng quét lại cho {folder}'
        })

    elif old_task.task_type == 'pdf_face':
        new_tid = f"pdf_face_{int(time.time() * 1000)}"

        def _retry_pdf_run(task_rec, cancel_check):
            try:
                send_log(f"🚀 [Quét lại] Phân tích khuôn mặt PDF: {folder} (Dự án: {project_name})", "info")
                matcher = get_face_matcher()
                from src.pdf_face_analyzer import PDFFaceAnalyzer
                analyzer = PDFFaceAnalyzer(
                    p_dir, i_dir, matcher, accuracy_mode=True,
                    match_distance_threshold=threshold, log_detail=True,
                    project_id=project_id, identity_registry=identity_reg
                )
                def _log(msg, t='default'):
                    send_log(msg, t)
                def _progress(completed, total, name, file_path):
                    task_manager.update_progress(task_rec.task_id, completed, total, name, file_path)

                files = analyzer.analyze_folder(
                    input_dir, output_dir, log_callback=_log,
                    progress_callback=_progress, report_options=report_options,
                    cancel_check=cancel_check,
                )
                aggregate_path = next(
                    (path for path in files if os.path.basename(path).startswith('GIAI_TRINH_')),
                    None,
                )
                if aggregate_path:
                    task_rec.aggregate_report = os.path.basename(aggregate_path)
                    task_rec.files.append({
                        'name': os.path.basename(aggregate_path),
                        'folder': folder,
                        'is_aggregate': True,
                    })
                task_rec.total = len(files)
                task_rec.progress = len(files)
                if cancel_check():
                    send_log(f"⚠️ Đã dừng quét lại PDF: {folder}", "warning")
                else:
                    send_log(f"🎉 Hoàn tất quét lại! Đã xuất {len(files)} file Word", "success")
            except Exception as e:
                import traceback
                task_rec.status = 'failed'
                task_rec.errors.append(str(e))
                send_log(f"❌ Lỗi quét lại PDF: {e}", "error")
                traceback.print_exc()

        new_task = task_manager.submit_task(
            task_type='pdf_face',
            target_folder=folder,
            project_name=project_name,
            execute_fn=_retry_pdf_run,
            payload=payload,
            task_id=new_tid,
        )
        pdf_face_tasks[new_tid] = new_task
        return jsonify({
            'success': True,
            'task_id': new_tid,
            'status': new_task.status,
            'message': f'Đã xếp hàng quét lại cho {folder}'
        })

    return jsonify({'success': False, 'error': f'Không hỗ trợ retry cho {old_task.task_type}'}), 400


@app.route('/api/tasks/recent')
def get_recent_tasks():
    """Lấy danh sách các tác vụ gần nhất từ TaskManager"""
    limit = request.args.get('limit', default=20, type=int)
    task_type = request.args.get('type', default=None, type=str)
    tasks = task_manager.get_recent_tasks(limit=limit, task_type=task_type)
    return jsonify({'success': True, 'tasks': tasks})


@app.route('/api/cache/face/clear', methods=['POST'])
def api_clear_face_cache():
    """Xóa toàn bộ bộ nhớ đệm vector khuôn mặt"""
    from src.face_matcher import clear_face_cache
    success = clear_face_cache()
    if success:
        send_log("🧹 Đã làm mới toàn bộ bộ nhớ đệm vector khuôn mặt", "success")
        return jsonify({'success': True, 'message': 'Đã xóa toàn bộ cache khuôn mặt'})
    return jsonify({'success': False, 'error': 'Lỗi khi xóa cache'}), 500


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


@app.route('/api/aggregate-reports')
def aggregate_reports():
    """Liệt kê file giải trình tổng hợp từ cả luồng quét Excel và PDF."""
    files = []
    sources = (
        ('Excel', EXCEL_FACE_OUTPUT_DIR, 'excel_face_download'),
        ('PDF', PDF_FACE_OUTPUT_DIR, 'pdf_face_download'),
    )
    for source, base_dir, download_endpoint in sources:
        if not os.path.isdir(base_dir):
            continue
        for folder in os.listdir(base_dir):
            folder_path = os.path.join(base_dir, folder)
            if not os.path.isdir(folder_path):
                continue
            for filename in os.listdir(folder_path):
                if not (
                    filename.upper().startswith('GIAI_TRINH_')
                    and filename.lower().endswith('.docx')
                ):
                    continue
                file_path = os.path.join(folder_path, filename)
                if not os.path.isfile(file_path):
                    continue
                modified_timestamp = os.path.getmtime(file_path)
                files.append({
                    'name': filename,
                    'source': source,
                    'folder': folder,
                    'size': os.path.getsize(file_path),
                    'modified': datetime.fromtimestamp(modified_timestamp).isoformat(),
                    'download_url': url_for(
                        download_endpoint,
                        folder=folder,
                        filename=filename,
                    ),
                    '_modified_timestamp': modified_timestamp,
                })

    files.sort(key=lambda item: item['_modified_timestamp'], reverse=True)
    for item in files:
        item.pop('_modified_timestamp', None)
    return jsonify({'success': True, 'files': files, 'count': len(files)})


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
        root_folder = INPUT_IMAGES_DIR
    elif target_type == 'results':
        root_folder = RESULTS_DIR
    elif target_type == 'excel_output':
        root_folder = EXCEL_OUTPUT_DIR
    else:
        return jsonify({'success': False, 'error': 'Loại thư mục không hợp lệ'}), 400

    root_folder = os.path.abspath(root_folder)
    folder = os.path.abspath(os.path.join(root_folder, subpath)) if subpath else root_folder
    try:
        if os.path.commonpath([root_folder, folder]) != root_folder:
            return jsonify({'success': False, 'error': 'Đường dẫn không hợp lệ'}), 400
    except ValueError:
        return jsonify({'success': False, 'error': 'Đường dẫn không hợp lệ'}), 400

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


# ==================== BỔ SUNG ẢNH CHẤM CÔNG ====================

@app.route('/api/supplement/records', methods=['GET'])
def supplement_get_records():
    """Lấy danh sách tất cả records bổ sung ảnh"""
    try:
        ps = get_photo_supplement()
        status_filter = request.args.get('status', '')
        if status_filter:
            records = ps.get_records_by_status(status_filter)
        else:
            records = ps.get_all_records()
        return jsonify({'success': True, 'records': records})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/upload', methods=['POST'])
def supplement_upload():
    """Upload ảnh chụp bù và tạo record mới"""
    try:
        if 'photo' not in request.files:
            return jsonify({'success': False, 'error': 'Không có file ảnh'}), 400
        
        file = request.files['photo']
        if not file.filename:
            return jsonify({'success': False, 'error': 'File rỗng'}), 400
        
        employee_name = request.form.get('employee_name', '').strip()
        target_date = request.form.get('target_date', '').strip()
        target_time = request.form.get('target_time', '').strip()
        
        if not employee_name or not target_date or not target_time:
            return jsonify({'success': False, 'error': 'Thiếu thông tin: tên nhân viên, ngày, giờ'}), 400
        
        # Validate date format
        try:
            datetime.strptime(target_date, '%Y-%m-%d')
        except ValueError:
            return jsonify({'success': False, 'error': 'Định dạng ngày không hợp lệ (cần YYYY-MM-DD)'}), 400
        
        # Validate time format  
        if not re.match(r'^\d{2}:\d{2}(:\d{2})?$', target_time):
            return jsonify({'success': False, 'error': 'Định dạng giờ không hợp lệ (cần HH:MM hoặc HH:MM:SS)'}), 400
        if len(target_time) == 5:
            target_time += ':00'
        
        # Lưu file tạm
        ext = os.path.splitext(file.filename)[1].lower() or '.jpg'
        temp_path = os.path.join(SUPPLEMENT_DIR, f"temp_upload{ext}")
        file.save(temp_path)
        
        # Tạo record
        ps = get_photo_supplement()
        record = ps.upload_photo(
            temp_path, employee_name, target_date, target_time,
            watermark_style=request.form.get('watermark_style', 'timestamp_camera'),
            watermark_position=request.form.get('watermark_position', 'bottom-left'),
            location_name=request.form.get('location_name', ''),
            gps_coords=request.form.get('gps_coords', ''),
            remove_old_watermark=request.form.get('remove_old_watermark', 'true').lower() == 'true',
            modify_exif=request.form.get('modify_exif', 'true').lower() == 'true',
        )
        
        # Xóa file tạm
        try:
            os.remove(temp_path)
        except Exception:
            pass
        
        return jsonify({'success': True, 'record': record.to_dict()})
    except Exception as e:
        logging.error(f"Supplement upload error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/process', methods=['POST'])
def supplement_process():
    """Xử lý ảnh: xóa watermark cũ + tạo watermark mới + sửa EXIF"""
    try:
        data = request.json or {}
        record_id = data.get('record_id', '')
        record_ids = data.get('record_ids', [])
        
        ps = get_photo_supplement()
        
        if record_id:
            result = ps.process_record(record_id)
            return jsonify({'success': True, 'record': result.to_dict()})
        elif record_ids:
            results = ps.batch_process(record_ids)
            return jsonify({'success': True, 'records': [r.to_dict() for r in results]})
        else:
            # Xử lý tất cả pending
            results = ps.batch_process()
            return jsonify({'success': True, 'records': [r.to_dict() for r in results]})
    except Exception as e:
        logging.error(f"Supplement process error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/download/<record_id>')
def supplement_download(record_id):
    """Tải ảnh đã xử lý"""
    try:
        ps = get_photo_supplement()
        record = ps.get_record(record_id)
        if not record or not record.processed_path:
            return jsonify({'success': False, 'error': 'Record không tồn tại hoặc chưa xử lý'}), 404
        
        if not os.path.exists(record.processed_path):
            return jsonify({'success': False, 'error': 'File không tồn tại'}), 404
        
        return send_file(
            record.processed_path,
            as_attachment=True,
            download_name=os.path.basename(record.processed_path)
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/preview/<record_id>')
def supplement_preview(record_id):
    """Xem preview ảnh (gốc hoặc đã xử lý)"""
    try:
        ps = get_photo_supplement()
        preview_path = ps.get_preview(record_id)
        if not preview_path or not os.path.exists(preview_path):
            return jsonify({'success': False, 'error': 'Không có ảnh preview'}), 404
        
        return send_file(preview_path, mimetype='image/jpeg')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/validate/<record_id>')
def supplement_validate(record_id):
    """Kiểm tra chất lượng ảnh đã xử lý"""
    try:
        ps = get_photo_supplement()
        result = ps.validate_result(record_id)
        return jsonify({'success': True, 'validation': result})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/delete', methods=['POST'])
def supplement_delete():
    """Xóa record và files liên quan"""
    try:
        data = request.json or {}
        record_id = data.get('record_id', '')
        if not record_id:
            return jsonify({'success': False, 'error': 'Thiếu record_id'}), 400
        
        ps = get_photo_supplement()
        success = ps.delete_record(record_id)
        return jsonify({'success': success})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/supplement/missing')
def supplement_missing_days():
    """Lấy danh sách ngày thiếu ảnh từ attendance processor"""
    try:
        from src.attendance_processor import AttendanceProcessor
        processor = AttendanceProcessor(CHAMCONG_DIR)
        processor.scan_all_files()
        missing = processor.get_missing_records()
        return jsonify({'success': True, 'missing': missing})
    except Exception as e:
        logging.error(f"Supplement missing days error: {e}")
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
