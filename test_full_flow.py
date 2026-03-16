# -*- coding: utf-8 -*-
"""
Test script để debug luồng quét mặt đầy đủ
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Add project root to path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from src.attendance_processor import AttendanceProcessor
from src.face_matcher import FaceMatcher

# Paths
CHAMCONG_DIR = os.path.join(BASE_DIR, "chamcong")
INPUT_IMAGES_DIR = os.path.join(BASE_DIR, "input_images")
PORTRAIT_DIR = os.path.join(BASE_DIR, "Ảnh BV")

print("=" * 70)
print("TEST FULL FLOW - QUÉT MẶT ẢNH CAMERA")
print("=" * 70)

# Step 1: Check directories
print("\n[STEP 1] Kiểm tra thư mục:")
print(f"  CHAMCONG_DIR: {CHAMCONG_DIR}")
print(f"    Exists: {os.path.exists(CHAMCONG_DIR)}")
if os.path.exists(CHAMCONG_DIR):
    files = os.listdir(CHAMCONG_DIR)
    print(f"    Files: {files}")
    
print(f"  PORTRAIT_DIR: {PORTRAIT_DIR}")
print(f"    Exists: {os.path.exists(PORTRAIT_DIR)}")

print(f"  INPUT_IMAGES_DIR: {INPUT_IMAGES_DIR}")
print(f"    Exists: {os.path.exists(INPUT_IMAGES_DIR)}")
if os.path.exists(INPUT_IMAGES_DIR):
    subfolders = os.listdir(INPUT_IMAGES_DIR)
    print(f"    Subfolders: {subfolders[:10]}...")

# Step 2: Parse attendance
print("\n[STEP 2] Phân tích file chấm công:")
processor = AttendanceProcessor(CHAMCONG_DIR)
processor.scan_all_files()
missing = processor.get_missing_records()
summary = processor.get_summary()

print(f"  Tổng số người: {summary['total_persons']}")
print(f"  Tổng số bản ghi thiếu: {summary['total_missing']}")

# Show first 5 missing records
print("  5 bản ghi đầu tiên:")
for i, rec in enumerate(missing[:5], 1):
    print(f"    {i}. {rec['person_name']} - {rec['date']} ({rec['weekday']}): {rec['issue_description']}")

# Step 3: Test face matcher
print("\n[STEP 3] Khởi tạo FaceMatcher:")
try:
    matcher = FaceMatcher(PORTRAIT_DIR)
    print(f"  ✅ FaceMatcher khởi tạo thành công")
    print(f"  Số người trong cache: {len(matcher.portrait_cache)}")
except Exception as e:
    print(f"  ❌ Lỗi: {e}")
    matcher = None

# Step 4: Test matching for first missing record
print("\n[STEP 4] Test match ảnh cho bản ghi đầu tiên:")
if missing and matcher:
    rec = missing[0]
    person_name = rec['person_name']
    date_str = rec['date']  # format: dd/mm/yyyy
    day = date_str.split('/')[0].zfill(2)
    
    print(f"  Person: {person_name}")
    print(f"  Date: {date_str}")
    print(f"  Day folder: {day}")
    
    # Find portraits for this person
    portraits = matcher.find_portraits(person_name)
    print(f"  Số ảnh chân dung tìm thấy: {len(portraits)}")
    for p in portraits:
        print(f"    - {os.path.basename(p)}")
    
    # Find camera images
    day_folder = os.path.join(INPUT_IMAGES_DIR, day)
    print(f"  Day folder path: {day_folder}")
    print(f"  Exists: {os.path.exists(day_folder)}")
    
    if os.path.exists(day_folder):
        # Get image files
        SUPPORTED = {'.jpg', '.jpeg', '.png', '.bmp'}
        images = []
        for f in os.listdir(day_folder):
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED:
                images.append(os.path.join(day_folder, f))
        print(f"  Số ảnh camera: {len(images)}")
        
        if images and portraits:
            print("\n  Bắt đầu match face (có thể mất vài phút)...")
            matched = matcher.match_face_in_images(person_name, images)
            if matched:
                print(f"  ✅ MATCHED: {os.path.basename(matched)}")
            else:
                print("  ❌ Không tìm thấy ảnh match")
                print("  Thử debug thêm với threshold cao hơn...")
                
                # Debug: try manual match
                from src.face_matcher import get_deepface, copy_to_ascii_path
                DeepFace = get_deepface()
                if DeepFace:
                    portrait = portraits[0]
                    print(f"  Test thủ công với portrait: {os.path.basename(portrait)}")
                    p_ascii = copy_to_ascii_path(portrait)
                    
                    best_dist = float('inf')
                    best_img = None
                    
                    for img in images[:10]:  # Test 10 ảnh đầu
                        try:
                            c_ascii = copy_to_ascii_path(img)
                            result = DeepFace.verify(
                                img1_path=p_ascii,
                                img2_path=c_ascii,
                                model_name="VGG-Face",
                                enforce_detection=False
                            )
                            dist = result.get('distance', 1.0)
                            verified = result.get('verified', False)
                            print(f"    {os.path.basename(img)}: distance={dist:.3f}, verified={verified}")
                            
                            if dist < best_dist:
                                best_dist = dist
                                best_img = img
                                
                        except Exception as e:
                            print(f"    {os.path.basename(img)}: ERROR - {str(e)[:50]}")
                    
                    print(f"\n  Best match: {os.path.basename(best_img) if best_img else 'None'} (distance={best_dist:.3f})")
                    print(f"  Threshold mặc định: 0.6")
                    if best_dist > 0.6:
                        print(f"  ⚠️ Best distance > 0.6, có thể cần tăng threshold!")
    else:
        print(f"  ❌ Thư mục ngày {day} không tồn tại!")
else:
    print("  Không có bản ghi hoặc matcher không khởi tạo được")

print("\n" + "=" * 70)
print("TEST COMPLETE")
print("=" * 70)
