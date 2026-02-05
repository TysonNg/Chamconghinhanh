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

# Thêm thư mục gốc vào path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

def open_browser():
    """Mở trình duyệt sau khi Flask khởi động"""
    time.sleep(2)  # Đợi Flask khởi động
    webbrowser.open('http://127.0.0.1:5000')
    print("\n✓ Đã mở trình duyệt...")
    print("✓ Để dừng phần mềm, đóng cửa sổ này hoặc nhấn Ctrl+C\n")

def main():
    """Entry point chính"""
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
    from src.app import app, scan_database, FLASK_PORT, FLASK_HOST
    
    print("Đang quét database...")
    scan_database()
    
    # Chạy Flask (blocking)
    app.run(host='127.0.0.1', port=FLASK_PORT, debug=False, threaded=True, use_reloader=False)

if __name__ == '__main__':
    main()
