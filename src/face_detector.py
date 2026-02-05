# -*- coding: utf-8 -*-
"""
Module nhận diện khuôn mặt
Sử dụng thư viện face_recognition (tùy chọn)
"""

import os
import numpy as np

try:
    import face_recognition
    FACE_RECOGNITION_AVAILABLE = True
except ImportError:
    FACE_RECOGNITION_AVAILABLE = False

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

# Import config với xử lý lỗi
try:
    from src.config import FACE_RECOGNITION_TOLERANCE, FACE_DETECTION_MODEL
except ImportError:
    FACE_RECOGNITION_TOLERANCE = 0.6
    FACE_DETECTION_MODEL = "hog"


def detect_faces(image_path):
    """
    Phát hiện khuôn mặt trong ảnh
    
    Returns:
        list: Danh sách vị trí khuôn mặt [(top, right, bottom, left), ...]
    """
    if not FACE_RECOGNITION_AVAILABLE:
        return []
    
    try:
        image = face_recognition.load_image_file(image_path)
        face_locations = face_recognition.face_locations(image, model=FACE_DETECTION_MODEL)
        return face_locations
    except Exception as e:
        print(f"Lỗi phát hiện khuôn mặt: {e}")
        return []


def get_face_encoding(image_path):
    """
    Tạo encoding cho khuôn mặt trong ảnh
    
    Returns:
        numpy.ndarray or None: Face encoding hoặc None nếu không tìm thấy
    """
    if not FACE_RECOGNITION_AVAILABLE:
        return None
    
    try:
        image = face_recognition.load_image_file(image_path)
        face_locations = face_recognition.face_locations(image, model=FACE_DETECTION_MODEL)
        
        if not face_locations:
            return None
        
        # Lấy encoding của khuôn mặt đầu tiên (lớn nhất)
        face_encodings = face_recognition.face_encodings(image, face_locations)
        
        if face_encodings:
            return face_encodings[0]
        
        return None
    except Exception as e:
        print(f"Lỗi tạo face encoding: {e}")
        return None


def get_all_face_encodings(image_path):
    """
    Lấy encoding của tất cả khuôn mặt trong ảnh
    
    Returns:
        list: Danh sách (face_location, face_encoding)
    """
    if not FACE_RECOGNITION_AVAILABLE:
        return []
    
    try:
        image = face_recognition.load_image_file(image_path)
        face_locations = face_recognition.face_locations(image, model=FACE_DETECTION_MODEL)
        
        if not face_locations:
            return []
        
        face_encodings = face_recognition.face_encodings(image, face_locations)
        
        return list(zip(face_locations, face_encodings))
    except Exception as e:
        print(f"Lỗi lấy face encodings: {e}")
        return []


def compare_faces(known_encoding, unknown_encoding, tolerance=None):
    """
    So sánh 2 khuôn mặt
    
    Returns:
        tuple: (is_match: bool, distance: float)
    """
    if not FACE_RECOGNITION_AVAILABLE:
        return False, 1.0
    
    if known_encoding is None or unknown_encoding is None:
        return False, 1.0
    
    if tolerance is None:
        tolerance = FACE_RECOGNITION_TOLERANCE
    
    try:
        # Tính khoảng cách
        distance = face_recognition.face_distance([known_encoding], unknown_encoding)[0]
        is_match = distance <= tolerance
        
        return is_match, float(distance)
    except Exception as e:
        print(f"Lỗi so sánh khuôn mặt: {e}")
        return False, 1.0


def find_best_match(unknown_encoding, known_faces_dict, tolerance=None):
    """
    Tìm khuôn mặt phù hợp nhất trong database
    
    Args:
        unknown_encoding: Encoding của khuôn mặt cần tìm
        known_faces_dict: Dict {person_id: {'encoding': ..., 'branch': ..., 'name': ...}}
        tolerance: Ngưỡng so sánh
    
    Returns:
        dict or None: Thông tin người phù hợp nhất hoặc None
    """
    if not FACE_RECOGNITION_AVAILABLE or unknown_encoding is None:
        return None
    
    if tolerance is None:
        tolerance = FACE_RECOGNITION_TOLERANCE
    
    best_match = None
    best_distance = float('inf')
    
    for person_id, person_data in known_faces_dict.items():
        known_encoding = person_data.get('encoding')
        if known_encoding is None:
            continue
        
        is_match, distance = compare_faces(known_encoding, unknown_encoding, tolerance)
        
        if is_match and distance < best_distance:
            best_distance = distance
            best_match = {
                'person_id': person_id,
                'branch': person_data.get('branch', 'Unknown'),
                'name': person_data.get('name', 'Unknown'),
                'distance': distance,
                'confidence': round((1 - distance) * 100, 2)  # Chuyển sang %
            }
    
    return best_match


def extract_face_image(image_path, face_location, padding=20):
    """
    Cắt khuôn mặt từ ảnh
    
    Returns:
        numpy.ndarray: Ảnh khuôn mặt đã cắt
    """
    if not CV2_AVAILABLE:
        return None
        
    try:
        image = cv2.imread(image_path)
        if image is None:
            return None
        
        top, right, bottom, left = face_location
        height, width = image.shape[:2]
        
        # Thêm padding
        top = max(0, top - padding)
        left = max(0, left - padding)
        bottom = min(height, bottom + padding)
        right = min(width, right + padding)
        
        face_image = image[top:bottom, left:right]
        return face_image
    except Exception as e:
        print(f"Lỗi cắt khuôn mặt: {e}")
        return None
