# -*- coding: utf-8 -*-
"""
Module quản lý database ảnh chân dung
Cấu trúc: database/chi_nhanh/ten_nguoi/anh.jpg
"""

import os
import json
import pickle
from pathlib import Path

from src.config import DATABASE_DIR, SUPPORTED_IMAGE_EXTENSIONS
from src.face_detector import get_face_encoding


# File cache cho encoding
CACHE_FILE = os.path.join(DATABASE_DIR, '.face_cache.pkl')


class DatabaseManager:
    def __init__(self):
        self.database = {}  # {person_id: {'encoding': ..., 'branch': ..., 'name': ..., 'image_path': ...}}
        self.branches = []  # Danh sách chi nhánh
        self._load_cache()
    
    def _get_person_id(self, branch, name):
        """Tạo ID duy nhất cho người"""
        return f"{branch}/{name}"
    
    def _load_cache(self):
        """Load cache từ file"""
        if os.path.exists(CACHE_FILE):
            try:
                with open(CACHE_FILE, 'rb') as f:
                    self.database = pickle.load(f)
                print(f"Đã load cache: {len(self.database)} người")
            except Exception as e:
                print(f"Lỗi load cache: {e}")
                self.database = {}
    
    def _save_cache(self):
        """Lưu cache vào file"""
        try:
            with open(CACHE_FILE, 'wb') as f:
                pickle.dump(self.database, f)
            print(f"Đã lưu cache: {len(self.database)} người")
        except Exception as e:
            print(f"Lỗi lưu cache: {e}")
    
    def scan_database(self, progress_callback=None):
        """
        Quét toàn bộ thư mục database và tạo encoding
        
        Args:
            progress_callback: Hàm callback(current, total, message)
        """
        self.database = {}
        self.branches = []
        
        if not os.path.exists(DATABASE_DIR):
            os.makedirs(DATABASE_DIR)
            return
        
        # Đếm tổng số người
        total_persons = 0
        for branch in os.listdir(DATABASE_DIR):
            branch_path = os.path.join(DATABASE_DIR, branch)
            if os.path.isdir(branch_path) and not branch.startswith('.'):
                self.branches.append(branch)
                for person in os.listdir(branch_path):
                    person_path = os.path.join(branch_path, person)
                    if os.path.isdir(person_path):
                        total_persons += 1
        
        current = 0
        for branch in self.branches:
            branch_path = os.path.join(DATABASE_DIR, branch)
            
            for person_name in os.listdir(branch_path):
                person_path = os.path.join(branch_path, person_name)
                
                if not os.path.isdir(person_path):
                    continue
                
                current += 1
                if progress_callback:
                    progress_callback(current, total_persons, f"Đang xử lý: {branch}/{person_name}")
                
                # Tìm ảnh trong thư mục của người này
                image_files = []
                for file in os.listdir(person_path):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in SUPPORTED_IMAGE_EXTENSIONS:
                        image_files.append(os.path.join(person_path, file))
                
                if not image_files:
                    continue
                
                # Lấy encoding từ ảnh đầu tiên
                image_path = image_files[0]
                encoding = get_face_encoding(image_path)
                
                if encoding is not None:
                    person_id = self._get_person_id(branch, person_name)
                    self.database[person_id] = {
                        'encoding': encoding,
                        'branch': branch,
                        'name': person_name,
                        'image_path': image_path
                    }
        
        self._save_cache()
        
        if progress_callback:
            progress_callback(total_persons, total_persons, f"Hoàn thành! {len(self.database)} người trong database")
    
    def get_all_faces(self):
        """Lấy tất cả khuôn mặt trong database"""
        return self.database
    
    def get_branches(self):
        """Lấy danh sách chi nhánh"""
        if not self.branches:
            self.branches = []
            if os.path.exists(DATABASE_DIR):
                for branch in os.listdir(DATABASE_DIR):
                    branch_path = os.path.join(DATABASE_DIR, branch)
                    if os.path.isdir(branch_path) and not branch.startswith('.'):
                        self.branches.append(branch)
        return self.branches
    
    def get_persons_in_branch(self, branch):
        """Lấy danh sách người trong chi nhánh"""
        persons = []
        branch_path = os.path.join(DATABASE_DIR, branch)
        
        if os.path.exists(branch_path):
            for person in os.listdir(branch_path):
                person_path = os.path.join(branch_path, person)
                if os.path.isdir(person_path):
                    persons.append({
                        'name': person,
                        'branch': branch,
                        'person_id': self._get_person_id(branch, person)
                    })
        
        return persons
    
    def add_branch(self, branch_name):
        """Thêm chi nhánh mới"""
        branch_path = os.path.join(DATABASE_DIR, branch_name)
        os.makedirs(branch_path, exist_ok=True)
        if branch_name not in self.branches:
            self.branches.append(branch_name)
        return True
    
    def add_person(self, branch, person_name, image_path):
        """
        Thêm người mới vào database
        
        Args:
            branch: Tên chi nhánh
            person_name: Tên người
            image_path: Đường dẫn ảnh chân dung
        """
        person_dir = os.path.join(DATABASE_DIR, branch, person_name)
        os.makedirs(person_dir, exist_ok=True)
        
        # Copy ảnh vào thư mục
        import shutil
        filename = os.path.basename(image_path)
        dest_path = os.path.join(person_dir, filename)
        shutil.copy2(image_path, dest_path)
        
        # Tạo encoding
        encoding = get_face_encoding(dest_path)
        
        if encoding is not None:
            person_id = self._get_person_id(branch, person_name)
            self.database[person_id] = {
                'encoding': encoding,
                'branch': branch,
                'name': person_name,
                'image_path': dest_path
            }
            self._save_cache()
            return True
        
        return False
    
    def get_database_stats(self):
        """Lấy thống kê database"""
        stats = {
            'total_persons': len(self.database),
            'total_branches': len(self.get_branches()),
            'branches': {}
        }
        
        for person_id, data in self.database.items():
            branch = data.get('branch', 'Unknown')
            if branch not in stats['branches']:
                stats['branches'][branch] = 0
            stats['branches'][branch] += 1
        
        return stats


# Singleton instance
_db_manager = None

def get_database_manager():
    global _db_manager
    if _db_manager is None:
        _db_manager = DatabaseManager()
    return _db_manager
