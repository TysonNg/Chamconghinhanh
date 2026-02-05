# -*- coding: utf-8 -*-
"""
Module trích xuất ngày tháng và địa điểm từ ảnh
Sử dụng OCR để đọc watermark (tùy chọn)
"""

import re
from datetime import datetime

# Import cv2 với xử lý lỗi
try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    cv2 = None
    np = None

# Import PIL
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    Image = None

# Import pytesseract với xử lý lỗi
OCR_AVAILABLE = False
try:
    import pytesseract
    try:
        from src.config import TESSERACT_CMD
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    except ImportError:
        pass
    OCR_AVAILABLE = True
except ImportError:
    pytesseract = None

# Mapping tháng tiếng Việt
VIETNAMESE_MONTHS = {
    'Th1': 1, 'Th01': 1,
    'Th2': 2, 'Th02': 2,
    'Th3': 3, 'Th03': 3,
    'Th4': 4, 'Th04': 4,
    'Th5': 5, 'Th05': 5,
    'Th6': 6, 'Th06': 6,
    'Th7': 7, 'Th07': 7,
    'Th8': 8, 'Th08': 8,
    'Th9': 9, 'Th09': 9,
    'Th10': 10,
    'Th11': 11,
    'Th12': 12,
}


def preprocess_image_for_ocr(image):
    """
    Tiền xử lý ảnh để OCR đọc tốt hơn
    """
    # Chuyển sang grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Tăng contrast
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    
    # Threshold để tách chữ
    _, thresh = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    return thresh


def extract_text_region(image, region='bottom'):
    """
    Cắt vùng chứa watermark (thường ở góc dưới)
    """
    height, width = image.shape[:2]
    
    if region == 'bottom':
        # Lấy 30% phía dưới ảnh (nơi thường có watermark)
        y_start = int(height * 0.7)
        cropped = image[y_start:height, 0:width]
    elif region == 'bottom_right':
        y_start = int(height * 0.7)
        x_start = int(width * 0.5)
        cropped = image[y_start:height, x_start:width]
    else:
        cropped = image
    
    return cropped


def parse_vietnamese_datetime(text):
    """
    Parse ngày tháng từ định dạng Việt Nam
    Ví dụ: "24 Th12, 2025 08:43:36"
    """
    # Pattern cho định dạng: DD ThMM, YYYY HH:MM:SS
    pattern = r'(\d{1,2})\s*(Th\d{1,2}),?\s*(\d{4})\s*(\d{2}):(\d{2}):(\d{2})'
    
    match = re.search(pattern, text)
    if match:
        day = int(match.group(1))
        month_str = match.group(2)
        year = int(match.group(3))
        hour = int(match.group(4))
        minute = int(match.group(5))
        second = int(match.group(6))
        
        month = VIETNAMESE_MONTHS.get(month_str, 1)
        
        try:
            dt = datetime(year, month, day, hour, minute, second)
            return dt.strftime('%d/%m/%Y %H:%M:%S')
        except ValueError:
            pass
    
    return None


def extract_location(text):
    """
    Trích xuất địa điểm từ text
    """
    # Tìm các pattern địa chỉ phổ biến
    lines = text.split('\n')
    location_parts = []
    
    for line in lines:
        line = line.strip()
        # Bỏ qua dòng chứa ngày tháng
        if re.search(r'\d{2}:\d{2}:\d{2}', line):
            continue
        # Tìm các pattern địa chỉ
        if any(keyword in line.lower() for keyword in ['đường', 'quận', 'phường', 'thành phố', 'tp', 'số']):
            location_parts.append(line)
        elif re.match(r'^\d+\s+\w+', line):  # Số + tên đường
            location_parts.append(line)
    
    return ', '.join(location_parts) if location_parts else None


def extract_datetime_and_location(image_path):
    """
    Trích xuất ngày tháng và địa điểm từ ảnh
    
    Returns:
        dict: {
            'datetime': str or None,
            'location': str or None,
            'raw_text': str
        }
    """
    result = {
        'datetime': None,
        'location': None,
        'raw_text': ''
    }
    
    if not OCR_AVAILABLE or not CV2_AVAILABLE:
        return result
    
    try:
        # Đọc ảnh
        image = cv2.imread(image_path)
        if image is None:
            return result
        
        # Cắt vùng chứa watermark
        cropped = extract_text_region(image, 'bottom')
        
        # Tiền xử lý
        processed = preprocess_image_for_ocr(cropped)
        
        # OCR
        custom_config = r'--oem 3 --psm 6 -l vie+eng'
        text = pytesseract.image_to_string(processed, config=custom_config)
        
        result['raw_text'] = text
        
        # Parse ngày tháng
        result['datetime'] = parse_vietnamese_datetime(text)
        
        # Trích xuất địa điểm
        result['location'] = extract_location(text)
        
    except Exception as e:
        result['error'] = str(e)
    
    return result


def extract_datetime_simple(image_path):
    """
    Phương pháp đơn giản hơn - đọc trực tiếp từ ảnh gốc
    """
    result = {
        'datetime': None,
        'location': None,
        'raw_text': ''
    }
    
    if not OCR_AVAILABLE or not CV2_AVAILABLE or not PIL_AVAILABLE:
        return result
        
    try:
        image = cv2.imread(image_path)
        if image is None:
            return result
            
        height, width = image.shape[:2]
        
        # Lấy vùng góc dưới phải (thường có timestamp)
        y_start = int(height * 0.75)
        cropped = image[y_start:height, :]
        
        # Chuyển sang RGB cho PIL
        rgb = cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(rgb)
        
        # OCR với config cho tiếng Việt
        text = pytesseract.image_to_string(pil_image, lang='vie+eng')
        result['raw_text'] = text
        
        # Parse
        result['datetime'] = parse_vietnamese_datetime(text)
        result['location'] = extract_location(text)
        
    except Exception as e:
        result['error'] = str(e)
    
    return result
