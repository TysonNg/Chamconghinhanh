# -*- coding: utf-8 -*-
"""
Main entry point cho desktop app
Tự động mở trình duyệt khi khởi động
"""

import os
import sys
import time
import webbrowser
import threading
import logging
import traceback

# Thêm thư mục gốc vào path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
sys.path.insert(0, BASE_DIR)

# Setup file logging ngay từ đầu (quan trọng khi chạy EXE không có console)
LOG_FILE = os.path.join(BASE_DIR, 'startup.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def log_info(msg):
    """Log ra cả console và file"""
    print(msg)
    logging.info(msg)

def log_error(msg):
    """Log lỗi ra cả console và file"""
    print(f"[ERROR] {msg}")
    logging.error(msg)

def open_browser():
    """Mở trình duyệt sau khi Flask khởi động"""
    time.sleep(2)  # Đợi Flask khởi động
    webbrowser.open('http://127.0.0.1:5000')
    log_info("✓ Đã mở trình duyệt...")
    log_info("✓ Để dừng phần mềm, đóng cửa sổ này hoặc nhấn Ctrl+C")

def main():
    """Entry point chính"""
    try:
        log_info("=" * 60)
        log_info("KHỞI ĐỘNG PHẦN MỀM")
        log_info(f"  Frozen: {getattr(sys, 'frozen', False)}")
        log_info(f"  BASE_DIR: {BASE_DIR}")
        log_info(f"  Python: {sys.version}")
        log_info(f"  Executable: {sys.executable}")
        if getattr(sys, 'frozen', False):
            log_info(f"  _MEIPASS: {sys._MEIPASS}")
        log_info("=" * 60)
        
        print("""
╔══════════════════════════════════════════════════════════════╗
║     PHẦN MỀM NHẬN DIỆN KHUÔN MẶT & CHẤM CÔNG                ║
║     Desktop App - Phiên bản 1.0                              ║
╠══════════════════════════════════════════════════════════════╣
║  Đang khởi động server...                                    ║
╚══════════════════════════════════════════════════════════════╝
        """)
        
        # Mở browser trong thread riêng
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()
        
        # Import và chạy Flask app
        log_info("Đang import Flask app...")
        from src.app import app, scan_database, FLASK_PORT, FLASK_HOST
        log_info("Import Flask app thành công!")
        
        log_info("Đang quét database...")
        scan_database()
        log_info("Quét database xong!")
        
        # Kiểm tra thư mục quan trọng
        important_dirs = {
            'input_images': os.path.join(BASE_DIR, 'input_images'),
            'database': os.path.join(BASE_DIR, 'database'),
            'chamcong': os.path.join(BASE_DIR, 'chamcong'),
            'Ảnh BV': os.path.join(BASE_DIR, 'Ảnh BV'),
            'results': os.path.join(BASE_DIR, 'results'),
        }
        for name, path in important_dirs.items():
            exists = os.path.exists(path)
            log_info(f"  [{('✓' if exists else '✗')}] {name}: {path}")
        
        # Chạy Flask (blocking)
        log_info(f"Đang khởi động Flask server tại 127.0.0.1:{FLASK_PORT}...")
        app.run(host='127.0.0.1', port=FLASK_PORT, debug=False, threaded=True, use_reloader=False)
        
    except Exception as e:
        error_msg = f"LỖI NGHIÊM TRỌNG: {str(e)}\n{traceback.format_exc()}"
        log_error(error_msg)
        
        # Hiển thị lỗi cho user nếu chạy từ EXE
        if getattr(sys, 'frozen', False):
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(
                    0, 
                    f"Lỗi khởi động phần mềm:\n\n{str(e)}\n\nXem chi tiết tại: {LOG_FILE}",
                    "Lỗi - Phần Mềm Chấm Công",
                    0x10  # MB_ICONERROR
                )
            except Exception:
                pass
        
        # Giữ console mở
        input("Nhấn Enter để đóng...")

if __name__ == '__main__':
    main()
