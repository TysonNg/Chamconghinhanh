# -*- coding: utf-8 -*-
"""
Script để tổ chức lại thư mục Ảnh BV theo tên người
Chạy một lần để setup cấu trúc thư mục mới
"""

import os
import shutil
import re

def normalize_name(name):
    """Chuẩn hóa tên thư mục (bỏ dấu không hợp lệ trong Windows)"""
    # Các ký tự không hợp lệ trong tên thư mục Windows
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '')
    return name.strip()

def organize_portraits(portrait_dir):
    """Tổ chức lại ảnh theo subfolder tên người"""
    
    if not os.path.exists(portrait_dir):
        print(f"Thư mục không tồn tại: {portrait_dir}")
        return
    
    supported_ext = {'.jpg', '.jpeg', '.png', '.bmp'}
    organized_count = 0
    
    # Lấy danh sách file ảnh hiện tại (không phải folder)
    files_to_move = []
    for item in os.listdir(portrait_dir):
        item_path = os.path.join(portrait_dir, item)
        if os.path.isfile(item_path):
            ext = os.path.splitext(item)[1].lower()
            if ext in supported_ext:
                files_to_move.append(item)
    
    if not files_to_move:
        print("Không có file ảnh nào ở cấp root để tổ chức")
        return
    
    print(f"Tìm thấy {len(files_to_move)} file ảnh cần tổ chức")
    
    for filename in files_to_move:
        # Lấy tên người từ tên file (bỏ extension)
        person_name = os.path.splitext(filename)[0]
        person_name = normalize_name(person_name)
        
        # Tạo thư mục cho người này nếu chưa có
        person_dir = os.path.join(portrait_dir, person_name)
        os.makedirs(person_dir, exist_ok=True)
        
        # Di chuyển file vào thư mục
        src = os.path.join(portrait_dir, filename)
        dst = os.path.join(person_dir, filename)
        
        try:
            shutil.move(src, dst)
            print(f"✓ {filename} -> {person_name}/")
            organized_count += 1
        except Exception as e:
            print(f"✗ Lỗi di chuyển {filename}: {e}")
    
    print(f"\nĐã tổ chức {organized_count} file vào thư mục con")
    print("Giờ bạn có thể thêm nhiều ảnh vào mỗi thư mục người để tăng độ chính xác nhận diện")

if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    portrait_dir = r'd:\Projects\phan mem quet mat\Ảnh BV'
    
    print("=== Tổ chức lại thư mục Ảnh Chân Dung ===")
    print(f"Thư mục: {portrait_dir}")
    print()
    
    response = input("Bạn có muốn tổ chức lại ảnh theo thư mục con theo tên? (y/n): ")
    if response.lower() in ['y', 'yes', 'có', 'co']:
        organize_portraits(portrait_dir)
    else:
        print("Đã hủy.")
