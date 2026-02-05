# -*- coding: utf-8 -*-
"""
Cấu hình cho phần mềm nhận diện khuôn mặt
"""

import os

# Đường dẫn gốc của dự án
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Đường dẫn các thư mục
INPUT_IMAGES_DIR = os.path.join(BASE_DIR, "input_images")
DATABASE_DIR = os.path.join(BASE_DIR, "database")
RESULTS_DIR = os.path.join(BASE_DIR, "results")

# Cấu hình nhận diện khuôn mặt
FACE_RECOGNITION_TOLERANCE = 0.6  # Ngưỡng so sánh (nhỏ hơn = chính xác hơn)
FACE_DETECTION_MODEL = "hog"  # "hog" (nhanh) hoặc "cnn" (chính xác hơn, cần GPU)

# Cấu hình xử lý bất đồng bộ
MAX_WORKERS = 4  # Số thread xử lý song song

# Cấu hình Tesseract OCR (đường dẫn trên Windows)
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Định dạng ảnh hỗ trợ
SUPPORTED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

# Cấu hình Flask
FLASK_HOST = "0.0.0.0"
FLASK_PORT = 5000
FLASK_DEBUG = True

# Tạo thư mục nếu chưa tồn tại
for directory in [INPUT_IMAGES_DIR, DATABASE_DIR, RESULTS_DIR]:
    os.makedirs(directory, exist_ok=True)
