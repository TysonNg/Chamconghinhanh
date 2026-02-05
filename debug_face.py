# -*- coding: utf-8 -*-
"""
Debug script để test face matching cho Le Van Dung ngày 18
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.face_matcher import FaceMatcher, normalize_vietnamese, get_deepface

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORTRAIT_DIR = os.path.join(BASE_DIR, "Ảnh BV")
INPUT_DIR = os.path.join(BASE_DIR, "input_images", "18")

print("=" * 60)
print("DEBUG: Face Matching cho Le Van Dung ngày 18")
print("=" * 60)

# 1. Kiểm tra thư mục
print(f"\n1. Kiểm tra thư mục:")
print(f"   Portrait dir: {PORTRAIT_DIR}")
print(f"   Exists: {os.path.exists(PORTRAIT_DIR)}")
print(f"   Input dir: {INPUT_DIR}")
print(f"   Exists: {os.path.exists(INPUT_DIR)}")

# 2. Load FaceMatcher
print(f"\n2. Load FaceMatcher...")
matcher = FaceMatcher(PORTRAIT_DIR)
print(f"   Số người trong cache: {len(matcher.portrait_cache)}")
print(f"   Các tên trong cache:")
for name in list(matcher.portrait_cache.keys())[:10]:
    print(f"      - {name} ({len(matcher.portrait_cache[name])} ảnh)")

# 3. Test tìm portrait cho "Le Van Dung"
print(f"\n3. Test tìm portrait:")
test_names = ["Le Van Dung", "Lê Văn Dũng", "le van dung"]
for name in test_names:
    portraits = matcher.find_portraits(name)
    print(f"   '{name}' -> {len(portraits)} ảnh")
    for p in portraits:
        print(f"      - {os.path.basename(p)}")

# 4. Test normalize
print(f"\n4. Test normalize:")
for name in ["Le Van Dung", "Lê Văn Dũng"]:
    norm = normalize_vietnamese(name)
    print(f"   '{name}' -> '{norm}'")

# 5. Lấy ảnh camera ngày 18
print(f"\n5. Ảnh camera ngày 18:")
camera_images = []
if os.path.exists(INPUT_DIR):
    for f in os.listdir(INPUT_DIR):
        if f.lower().endswith(('.jpg', '.jpeg', '.png')):
            camera_images.append(os.path.join(INPUT_DIR, f))
print(f"   Số ảnh: {len(camera_images)}")

# 6. Test DeepFace với 1 cặp ảnh
print(f"\n6. Test DeepFace matching:")
DeepFace = get_deepface()
if DeepFace:
    print("   DeepFace loaded OK")
    
    # Lấy portrait của Le Van Dung
    portraits = matcher.find_portraits("Le Van Dung")
    if portraits and camera_images:
        portrait = portraits[0]
        print(f"   Portrait: {os.path.basename(portrait)}")
        
        # Test với 5 ảnh camera đầu tiên
        print(f"\n   Testing với 5 ảnh camera đầu:")
        for i, camera_img in enumerate(camera_images[:5]):
            try:
                from src.face_matcher import copy_to_ascii_path
                p_ascii = copy_to_ascii_path(portrait)
                c_ascii = copy_to_ascii_path(camera_img)
                
                result = DeepFace.verify(
                    img1_path=p_ascii,
                    img2_path=c_ascii,
                    model_name="VGG-Face",
                    enforce_detection=False
                )
                distance = result.get('distance', 1.0)
                verified = result.get('verified', False)
                print(f"   [{i+1}] {os.path.basename(camera_img)}: distance={distance:.3f}, verified={verified}")
            except Exception as e:
                print(f"   [{i+1}] {os.path.basename(camera_img)}: ERROR - {str(e)[:50]}")
else:
    print("   DeepFace KHÔNG load được!")

print("\n" + "=" * 60)
print("DEBUG COMPLETE")
print("=" * 60)
