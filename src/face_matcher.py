# -*- coding: utf-8 -*-
"""
Module nhận diện khuôn mặt - so sánh ảnh camera với ảnh chân dung
"""

import os
import re
import shutil
import tempfile
import unicodedata
from typing import List, Dict, Optional, Tuple

# Lazy loading để tránh import lỗi
_deepface = None

# Thư mục temp để lưu ảnh tạm (tránh lỗi đường dẫn tiếng Việt)
_temp_dir = None

def get_temp_dir():
    """Lấy hoặc tạo thư mục temp"""
    global _temp_dir
    if _temp_dir is None or not os.path.exists(_temp_dir):
        _temp_dir = tempfile.mkdtemp(prefix="face_matcher_")
    return _temp_dir

def copy_to_ascii_path(src_path: str) -> str:
    """
    Copy file sang thư mục temp với tên ASCII
    Giải quyết lỗi DeepFace không đọc được đường dẫn tiếng Việt
    """
    if not os.path.exists(src_path):
        print(f"  [copy_to_ascii] Source không tồn tại: {src_path}")
        return src_path
    
    # Luôn copy để tránh lỗi encoding
    try:
        # Tạo tên file ASCII dựa trên hash
        ext = os.path.splitext(src_path)[1]
        # Dùng abs path để hash unique hơn
        abs_path = os.path.abspath(src_path)
        ascii_name = f"img_{hash(abs_path) & 0xFFFFFFFF}{ext}"
        
        temp_dir = get_temp_dir()
        dst_path = os.path.join(temp_dir, ascii_name)
        
        # Copy nếu chưa tồn tại hoặc source mới hơn
        if not os.path.exists(dst_path):
            shutil.copy2(src_path, dst_path)
        
        return dst_path
    except Exception as e:
        print(f"  [copy_to_ascii] Error: {e}")
        return src_path

def get_deepface():
    """Lazy load DeepFace để giảm thời gian khởi động"""
    global _deepface
    if _deepface is None:
        try:
            from deepface import DeepFace
            _deepface = DeepFace
        except ImportError:
            print("DeepFace chưa được cài đặt. Chạy: pip install deepface tf-keras")
            return None
    return _deepface


def normalize_vietnamese(text: str) -> str:
    """
    Chuẩn hóa tên tiếng Việt - loại bỏ dấu và chuyển lowercase
    Ví dụ: "Lê Văn Tòng" -> "levantong"
    """
    if not text:
        return ""
    
    # Bảng chuyển đổi tiếng Việt
    vietnamese_map = {
        'à': 'a', 'á': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a',
        'ă': 'a', 'ằ': 'a', 'ắ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a',
        'â': 'a', 'ầ': 'a', 'ấ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
        'đ': 'd',
        'è': 'e', 'é': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e',
        'ê': 'e', 'ề': 'e', 'ế': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
        'ì': 'i', 'í': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
        'ò': 'o', 'ó': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o',
        'ô': 'o', 'ồ': 'o', 'ố': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o',
        'ơ': 'o', 'ờ': 'o', 'ớ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
        'ù': 'u', 'ú': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u',
        'ư': 'u', 'ừ': 'u', 'ứ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
        'ỳ': 'y', 'ý': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y',
    }
    
    # Lowercase và thay thế ký tự tiếng Việt
    result = text.lower()
    for vn_char, ascii_char in vietnamese_map.items():
        result = result.replace(vn_char, ascii_char)
    
    # Loại bỏ khoảng trắng và ký tự đặc biệt
    result = re.sub(r'[^a-z0-9]', '', result)
    
    return result


def calculate_name_similarity(name1: str, name2: str) -> float:
    """
    Tính độ tương đồng giữa 2 tên (0.0 - 1.0)
    Sử dụng thuật toán đơn giản dựa trên substring matching
    """
    n1 = normalize_vietnamese(name1)
    n2 = normalize_vietnamese(name2)
    
    if not n1 or not n2:
        return 0.0
    
    # Exact match
    if n1 == n2:
        return 1.0
    
    # Substring match
    if n1 in n2 or n2 in n1:
        shorter = min(len(n1), len(n2))
        longer = max(len(n1), len(n2))
        return shorter / longer
    
    # Prefix/suffix match
    common_prefix = 0
    for i in range(min(len(n1), len(n2))):
        if n1[i] == n2[i]:
            common_prefix += 1
        else:
            break
    
    return common_prefix / max(len(n1), len(n2))


class FaceMatcher:
    """So sánh khuôn mặt giữa ảnh camera và ảnh chân dung"""
    
    def __init__(self, portrait_dir: str, model_name: str = "VGG-Face", log_callback=None):
        """
        Args:
            portrait_dir: Thư mục chứa ảnh chân dung (Ảnh BV/)
            model_name: Model nhận diện (VGG-Face, Facenet, ArcFace, etc.)
            log_callback: Hàm callback để gửi log (optional)
        """
        self.portrait_dir = portrait_dir
        self.model_name = model_name
        self.portrait_cache = {}  # {person_name: [portrait_paths]}
        self.log_callback = log_callback
        self._scan_portraits()
    
    def _log(self, message: str, log_type: str = "default"):
        """Gửi log qua callback hoặc print"""
        if self.log_callback:
            self.log_callback(message, log_type)
        print(message)  # Always print to console too
    
    def _scan_portraits(self):
        """Quét thư mục chân dung và cache đường dẫn"""
        if not os.path.exists(self.portrait_dir):
            print(f"Thư mục ảnh chân dung không tồn tại: {self.portrait_dir}")
            return
        
        supported_ext = {'.jpg', '.jpeg', '.png', '.bmp'}
        
        for item in os.listdir(self.portrait_dir):
            item_path = os.path.join(self.portrait_dir, item)
            
            if os.path.isdir(item_path):
                # Thư mục con = tên người
                person_name = item
                images = []
                for file in os.listdir(item_path):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in supported_ext:
                        images.append(os.path.join(item_path, file))
                
                if images:
                    self.portrait_cache[person_name] = images
            else:
                # File trực tiếp = tên file là tên người
                ext = os.path.splitext(item)[1].lower()
                if ext in supported_ext:
                    person_name = os.path.splitext(item)[0]
                    self.portrait_cache[person_name] = [item_path]
        
        print(f"Đã load {len(self.portrait_cache)} người từ thư mục chân dung")
    
    def find_portrait(self, person_name: str) -> Optional[str]:
        """Tìm ảnh chân dung đầu tiên cho một người (backward compatible)"""
        portraits = self.find_portraits(person_name)
        return portraits[0] if portraits else None
    
    def find_portraits(self, person_name: str) -> List[str]:
        """
        Tìm TẤT CẢ ảnh chân dung cho một người
        Hỗ trợ matching tên tiếng Việt có/không dấu
        """
        # 1. Exact match
        if person_name in self.portrait_cache:
            return self.portrait_cache[person_name]
        
        # 2. Normalize và tìm exact match sau khi chuẩn hóa
        person_normalized = normalize_vietnamese(person_name)
        
        for cached_name, images in self.portrait_cache.items():
            cached_normalized = normalize_vietnamese(cached_name)
            
            # Exact match sau khi normalize
            if person_normalized == cached_normalized:
                return images
        
        # 3. Fuzzy match với similarity score
        best_images = None
        best_score = 0.0
        min_threshold = 0.7  # Yêu cầu ít nhất 70% tương đồng
        
        for cached_name, images in self.portrait_cache.items():
            score = calculate_name_similarity(person_name, cached_name)
            
            if score > best_score and score >= min_threshold:
                best_score = score
                best_images = images
        
        if best_images:
            return best_images
        
        # 4. Fallback: substring match
        for cached_name, images in self.portrait_cache.items():
            cached_normalized = normalize_vietnamese(cached_name)
            
            if person_normalized in cached_normalized or cached_normalized in person_normalized:
                return images
        
        return []
    
    def match_face_in_images(self, person_name: str, camera_images: List[str], 
                              distance_threshold: float = 0.6) -> Optional[str]:
        """
        Tìm ảnh camera có khuôn mặt match với người được chỉ định
        
        Args:
            person_name: Tên người cần tìm
            camera_images: Danh sách đường dẫn ảnh camera
            distance_threshold: Ngưỡng khoảng cách (thấp hơn = giống hơn), default 0.6
            
        Returns:
            Đường dẫn ảnh camera match tốt nhất, hoặc None nếu không tìm thấy
        """
        DeepFace = get_deepface()
        if DeepFace is None:
            self._log("  ❌ DeepFace không load được!", "error")
            return None
        
        # Tìm TẤT CẢ ảnh chân dung
        portrait_paths = self.find_portraits(person_name)
        if not portrait_paths:
            self._log(f"  ❌ Không tìm thấy ảnh chân dung cho: {person_name}", "error")
            self._log(f"     Cache có {len(self.portrait_cache)} người: {list(self.portrait_cache.keys())[:5]}...", "warning")
            return None
        
        self._log(f"  → Tìm thấy {len(portrait_paths)} ảnh chân dung", "info")
        for p in portrait_paths:
            exists = os.path.exists(p)
            self._log(f"     - {os.path.basename(p)} (exists={exists})", "default")
        
        best_match = None
        best_distance = float('inf')
        best_portrait = None
        errors_count = 0
        compared_count = 0
        
        # Thử với TỪNG ảnh chân dung
        for portrait_path in portrait_paths:
            if not os.path.exists(portrait_path):
                self._log(f"  ⚠️ Portrait không tồn tại: {portrait_path}", "warning")
                continue
                
            # Copy portrait sang đường dẫn ASCII nếu cần
            portrait_ascii = copy_to_ascii_path(portrait_path)
            
            # Log một lần cho mỗi portrait
            if portrait_path == portrait_paths[0]:
                self._log(f"  📂 Portrait ASCII: {portrait_ascii}", "info")
                self._log(f"     exists={os.path.exists(portrait_ascii)}", "info")
            
            # So sánh với từng ảnh camera
            for i, camera_img in enumerate(camera_images):
                try:
                    if not os.path.exists(camera_img):
                        continue
                        
                    # Copy camera image sang đường dẫn ASCII nếu cần
                    camera_ascii = copy_to_ascii_path(camera_img)
                    
                    result = DeepFace.verify(
                        img1_path=portrait_ascii,
                        img2_path=camera_ascii,
                        model_name=self.model_name,
                        enforce_detection=False  # Không lỗi nếu không detect được mặt
                    )
                    
                    distance = result.get('distance', 1.0)
                    compared_count += 1
                    
                    # Track best match
                    if distance < best_distance:
                        best_distance = distance
                        best_match = camera_img  # Trả về đường dẫn gốc
                        best_portrait = portrait_path
                        
                except Exception as e:
                    errors_count += 1
                    # Log lỗi đầu tiên để debug
                    if errors_count <= 2:
                        self._log(f"    ⚠ Error #{errors_count}: {str(e)[:80]}", "warning")
        
        # Log summary
        self._log(f"  📊 So sánh: {compared_count} lần, lỗi: {errors_count} lần", "info")
        
        # Trả về best match nếu distance đủ thấp
        if best_match and best_distance <= distance_threshold:
            self._log(f"  ✓ Best Match: {os.path.basename(best_match)} (distance={best_distance:.3f})", "success")
            return best_match
        elif best_match:
            self._log(f"  → Best distance={best_distance:.3f} > threshold={distance_threshold}", "warning")
        else:
            self._log(f"  → Không tìm thấy ảnh nào match được", "error")
        
        return None
    
    def match_all_faces(self, person_name: str, camera_images: List[str], 
                        max_matches: int = 1) -> List[Tuple[str, float]]:
        """
        Tìm tất cả ảnh camera match với người được chỉ định
        
        Returns:
            List of (image_path, confidence) tuples
        """
        DeepFace = get_deepface()
        if DeepFace is None:
            return []
        
        portrait_path = self.find_portrait(person_name)
        if not portrait_path:
            return []
        
        matches = []
        
        for camera_img in camera_images:
            try:
                result = DeepFace.verify(
                    img1_path=portrait_path,
                    img2_path=camera_img,
                    model_name=self.model_name,
                    enforce_detection=False
                )
                
                if result.get('verified', False):
                    distance = result.get('distance', 1.0)
                    confidence = max(0, 100 * (1 - distance))  # Convert to percentage
                    matches.append((camera_img, confidence))
                    
                    if len(matches) >= max_matches:
                        break
                        
            except Exception:
                continue
        
        # Sắp xếp theo confidence giảm dần
        matches.sort(key=lambda x: x[1], reverse=True)
        return matches


def simple_face_match(portrait_path: str, camera_images: List[str]) -> Optional[str]:
    """
    So sánh đơn giản - trả về ảnh camera đầu tiên match với portrait
    """
    DeepFace = get_deepface()
    if DeepFace is None:
        return None
    
    for camera_img in camera_images:
        try:
            result = DeepFace.verify(
                img1_path=portrait_path,
                img2_path=camera_img,
                model_name="VGG-Face",
                enforce_detection=False
            )
            
            if result.get('verified', False):
                return camera_img
                
        except Exception:
            continue
    
    return None


# Test
if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    portrait_dir = r'd:\Projects\phan mem quet mat\Ảnh BV'
    matcher = FaceMatcher(portrait_dir)
    
    print("\n=== Test FaceMatcher ===")
    print(f"Số người trong cache: {len(matcher.portrait_cache)}")
    
    # Test tìm portrait
    test_name = "Nguyen Van A"
    portrait = matcher.find_portrait(test_name)
    print(f"Portrait cho {test_name}: {portrait}")
