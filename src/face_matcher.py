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
import numpy as np

import sys
import pickle
import threading

# Cấu hình an toàn console Windows tránh UnicodeEncodeError
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Lazy loading để tránh import lỗi
_deepface = None
_deepface_lock = threading.Lock()

# Thư mục temp để lưu ảnh tạm (tránh lỗi đường dẫn tiếng Việt)
_temp_dir = None
_temp_dir_lock = threading.Lock()

# Bộ nhớ đệm vector khuôn mặt lưu trên ổ đĩa (.cache/)
_CACHE_LOCK = threading.Lock()
_DISK_CACHE = None
_CACHE_FILE = None
_CACHE_DIRTY_COUNT = 0

def _get_cache_file():
    global _CACHE_FILE
    if _CACHE_FILE is None:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        cache_dir = os.path.join(base_dir, ".cache")
        os.makedirs(cache_dir, exist_ok=True)
        _CACHE_FILE = os.path.join(cache_dir, "face_embeddings.pkl")
    return _CACHE_FILE

def load_disk_cache():
    global _DISK_CACHE
    with _CACHE_LOCK:
        if _DISK_CACHE is not None:
            return _DISK_CACHE
        cache_file = _get_cache_file()
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'rb') as f:
                    _DISK_CACHE = pickle.load(f)
                    print(f"[CACHE] Đã nạp {len(_DISK_CACHE)} vector khuôn mặt từ ổ đĩa.")
            except Exception as e:
                print(f"[CACHE] Khởi tạo cache mới: {e}")
                _DISK_CACHE = {}
        else:
            _DISK_CACHE = {}
        return _DISK_CACHE

def save_disk_cache(force=False):
    global _DISK_CACHE, _CACHE_DIRTY_COUNT
    with _CACHE_LOCK:
        if _DISK_CACHE is None or (_CACHE_DIRTY_COUNT == 0 and not force):
            return
        cache_file = _get_cache_file()
        try:
            temp_file = cache_file + ".tmp"
            with open(temp_file, 'wb') as f:
                pickle.dump(_DISK_CACHE, f, protocol=pickle.HIGHEST_PROTOCOL)
            if os.path.exists(cache_file):
                os.remove(cache_file)
            os.rename(temp_file, cache_file)
            _CACHE_DIRTY_COUNT = 0
            print(f"[CACHE] Đã lưu {len(_DISK_CACHE)} vector khuôn mặt vào ổ đĩa.")
        except Exception as e:
            print(f"[CACHE] Lỗi lưu cache: {e}")

def get_temp_dir():
    """Lấy hoặc tạo thư mục temp (thread-safe)"""
    global _temp_dir
    if _temp_dir is None or not os.path.exists(_temp_dir):
        with _temp_dir_lock:
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
    
    try:
        ext = os.path.splitext(src_path)[1]
        abs_path = os.path.abspath(src_path)
        ascii_name = f"img_{hash(abs_path) & 0xFFFFFFFF}{ext}"
        
        temp_dir = get_temp_dir()
        dst_path = os.path.join(temp_dir, ascii_name)
        
        if (not os.path.exists(dst_path) or
                os.path.getmtime(src_path) > os.path.getmtime(dst_path)):
            shutil.copy2(src_path, dst_path)
        
        return dst_path
    except Exception as e:
        print(f"  [copy_to_ascii] Error: {e}")
        return src_path

def get_deepface():
    """Lazy load DeepFace để giảm thời gian khởi động (thread-safe)"""
    global _deepface
    if _deepface is None:
        with _deepface_lock:
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

    # Chuẩn hóa Unicode và loại bỏ dấu (combining marks)
    result = unicodedata.normalize('NFD', str(text))
    result = ''.join(c for c in result if unicodedata.category(c) != 'Mn')

    # Chuyển đ/Đ về d
    result = result.replace('đ', 'd').replace('Đ', 'd')

    # Lowercase + bỏ ký tự không phải chữ số
    result = result.lower()
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
    
    def __init__(
        self,
        portrait_dir: str,
        model_name: str = "ArcFace",
        detector_backend: str = "retinaface",
        distance_metric: str = "cosine",
        enforce_detection: bool = True,
        log_callback=None
    ):
        """
        Args:
            portrait_dir: Thư mục chứa ảnh chân dung (Ảnh BV/)
            model_name: Model nhận diện (VGG-Face, Facenet, ArcFace, etc.)
            detector_backend: Backend detect mặt (retinaface/mtcnn/opencv)
            distance_metric: Metric so sánh (cosine/euclidean)
            enforce_detection: Bắt buộc detect mặt để tăng độ chính xác
            log_callback: Hàm callback để gửi log (optional)
        """
        self.portrait_dir = portrait_dir
        self.model_name = model_name
        self.detector_backend = detector_backend
        self.distance_metric = distance_metric
        self.enforce_detection = enforce_detection
        self.portrait_cache = {}  # {person_name: [portrait_paths]}
        self.project_portrait_cache = {}  # {project_name: {person_name: [portrait_paths]}}
        self._embedding_cache = {}  # {path: (mtime, embedding)}
        self.log_callback = log_callback
        self._scan_portraits()

    def _log(self, message: str, log_type: str = "default"):
        """Gửi log qua callback hoặc print"""
        if self.log_callback:
            self.log_callback(message, log_type)
        print(message)  # Always print to console too
    
    def _scan_portraits(self):
        """Quét thư mục chân dung và cache đường dẫn, hỗ trợ cấu trúc 1 cấp hoặc 2 cấp (theo dự án)"""
        self.portrait_cache = {}
        self.project_portrait_cache = {}
        if not os.path.exists(self.portrait_dir):
            print(f"Thư mục ảnh chân dung không tồn tại: {self.portrait_dir}")
            return
        
        supported_ext = {'.jpg', '.jpeg', '.png', '.bmp'}

        def _add_person_images(p_name: str, paths: List[str], proj_name: Optional[str] = None):
            if not paths:
                return
            self.portrait_cache.setdefault(p_name, []).extend(paths)
            if proj_name:
                self.project_portrait_cache.setdefault(proj_name, {}).setdefault(p_name, []).extend(paths)
        
        for item in os.listdir(self.portrait_dir):
            item_path = os.path.join(self.portrait_dir, item)
            
            if os.path.isdir(item_path):
                sub_items = os.listdir(item_path)
                sub_dirs = [s for s in sub_items if os.path.isdir(os.path.join(item_path, s))]
                img_files = [os.path.join(item_path, s) for s in sub_items if os.path.splitext(s)[1].lower() in supported_ext]
                
                if sub_dirs:
                    # Thư mục Dự Án (chứa các thư mục con là từng nhân viên)
                    project_name = item
                    for s_dir in sub_dirs:
                        person_path = os.path.join(item_path, s_dir)
                        p_images = [
                            os.path.join(person_path, f)
                            for f in os.listdir(person_path)
                            if os.path.splitext(f)[1].lower() in supported_ext
                        ]
                        _add_person_images(s_dir, p_images, project_name)
                    if img_files:
                        for img in img_files:
                            p_name = os.path.splitext(os.path.basename(img))[0]
                            _add_person_images(p_name, [img], project_name)
                else:
                    # Thư mục Nhân Viên trực tiếp
                    _add_person_images(item, img_files)
            else:
                ext = os.path.splitext(item)[1].lower()
                if ext in supported_ext:
                    person_name = os.path.splitext(item)[0]
                    _add_person_images(person_name, [item_path])
        
        print(f"Đã load {len(self.portrait_cache)} người từ thư mục chân dung ({len(self.project_portrait_cache)} dự án)")

    def reload_portraits(self):
        """Tải lại danh sách ảnh chân dung (dùng khi thêm/sửa/xóa/chuyển dự án)"""
        self._scan_portraits()
    
    def find_portrait(self, person_name: str, project_name: Optional[str] = None) -> Optional[str]:
        """Tìm ảnh chân dung đầu tiên cho một người (backward compatible)"""
        portraits = self.find_portraits(person_name, project_name=project_name)
        return portraits[0] if portraits else None
    
    def find_portraits(self, person_name: str, project_name: Optional[str] = None) -> List[str]:
        """
        Tìm TẤT CẢ ảnh chân dung cho một người
        Hỗ trợ matching tên tiếng Việt có/không dấu, ưu tiên tìm trong Dự Án nếu có
        """
        cache = self.portrait_cache
        if project_name and project_name in self.project_portrait_cache:
            cache = self.project_portrait_cache[project_name]

        # 1. Exact match
        if person_name in cache:
            return cache[person_name]
        
        # 2. Normalize và tìm exact match sau khi chuẩn hóa
        person_normalized = normalize_vietnamese(person_name)
        
        for cached_name, images in cache.items():
            cached_normalized = normalize_vietnamese(cached_name)
            
            # Exact match sau khi normalize
            if person_normalized == cached_normalized:
                return images
        
        # 3. Fuzzy match với similarity score
        best_images = None
        best_score = 0.0
        min_threshold = 0.7  # Yêu cầu ít nhất 70% tương đồng
        
        for cached_name, images in cache.items():
            score = calculate_name_similarity(person_name, cached_name)
            
            if score > best_score and score >= min_threshold:
                best_score = score
                best_images = images
        
        if best_images:
            return best_images
        
        # 4. Fallback: substring match
        for cached_name, images in cache.items():
            cached_normalized = normalize_vietnamese(cached_name)
            
            if person_normalized in cached_normalized or cached_normalized in person_normalized:
                return images
        
        # 5. Fallback tìm toàn cục nếu chưa tìm thấy trong project cụ thể
        if project_name and cache is not self.portrait_cache:
            return self.find_portraits(person_name, project_name=None)

        return []
    
    def _cosine_distance(self, a: np.ndarray, b: np.ndarray) -> float:
        denom = (np.linalg.norm(a) * np.linalg.norm(b)) + 1e-12
        return float(1.0 - np.dot(a, b) / denom)

    def _get_embedding(self, image_path: str) -> Optional[np.ndarray]:
        if not image_path or not os.path.exists(image_path):
            return None

        try:
            stat = os.stat(image_path)
            cache_key = (
                os.path.abspath(image_path),
                stat.st_size,
                stat.st_mtime,
                self.model_name,
                self.detector_backend
            )

            # 1. Kiểm tra cache chung
            global _DISK_CACHE, _CACHE_DIRTY_COUNT
            with _CACHE_LOCK:
                if _DISK_CACHE is None:
                    load_disk_cache()
                if cache_key in _DISK_CACHE:
                    return _DISK_CACHE[cache_key]

            # 2. Tính toán vector với DeepFace
            DeepFace = get_deepface()
            if DeepFace is None:
                return None

            img_ascii = copy_to_ascii_path(image_path)
            reps = DeepFace.represent(
                img_path=img_ascii,
                model_name=self.model_name,
                detector_backend=self.detector_backend,
                enforce_detection=self.enforce_detection
            )
            if not reps:
                return None

            embedding = np.array(reps[0]["embedding"], dtype=np.float32)

            # 3. Ghi vào cache chung
            with _CACHE_LOCK:
                if _DISK_CACHE is not None:
                    _DISK_CACHE[cache_key] = embedding
                    _CACHE_DIRTY_COUNT += 1
                    if _CACHE_DIRTY_COUNT >= 20:
                        save_disk_cache()

            return embedding
        except Exception:
            return None

    def save_cache(self):
        """Lưu toàn bộ cache xuống đĩa"""
        save_disk_cache(force=True)

    def _get_default_threshold(self) -> float:
        if self.distance_metric == "cosine":
            model = self.model_name.lower()
            if model == "arcface":
                return 0.37
            if model == "facenet512":
                return 0.38
            if model == "vgg-face":
                return 0.45
        return 0.40

    def match_face_in_images(
        self,
        person_name: str,
        camera_images: List[str],
        distance_threshold: Optional[float] = None,
        fast_mode: bool = True,
        log_detail: bool = False,
        project_name: Optional[str] = None
    ) -> Optional[str]:
        """
        Tìm ảnh camera có khuôn mặt match với người được chỉ định

        Args:
            person_name: Tên người cần tìm
            camera_images: Danh sách đường dẫn ảnh camera
            distance_threshold: Ngưỡng khoảng cách (thấp hơn = giống hơn).
                                Nếu None sẽ dùng ngưỡng mặc định theo model.
            fast_mode: Tối ưu tốc độ
            log_detail: In log chi tiết
            project_name: Tên dự án (tùy chọn) để lọc ảnh chân dung

        Returns:
            Đường dẫn ảnh camera match tốt nhất, hoặc None nếu không tìm thấy
        """
        if distance_threshold is None:
            distance_threshold = self._get_default_threshold()

        # Tìm tất cả ảnh chân dung
        portrait_paths = self.find_portraits(person_name, project_name=project_name)
        if not portrait_paths:
            self._log(f"  [ERROR] Không tìm thấy ảnh chân dung cho: {person_name}", "error")
            self._log(
                f"     Cache có {len(self.portrait_cache)} người: {list(self.portrait_cache.keys())[:5]}...",
                "warning"
            )
            return None

        self._log(f"  -> Tìm thấy {len(portrait_paths)} ảnh chân dung", "info")
        for p in portrait_paths:
            exists = os.path.exists(p)
            self._log(f"     - {os.path.basename(p)} (exists={exists})", "default")

        # Tạo embedding cho ảnh chân dung
        portrait_embeddings = []
        for p in portrait_paths:
            emb = self._get_embedding(p)
            if emb is not None:
                portrait_embeddings.append((p, emb))

        if not portrait_embeddings:
            self._log("  [ERROR] Không tạo được embedding cho ảnh chân dung", "error")
            return None

        best_match = None
        best_distance = float('inf')
        errors_count = 0
        compared_count = 0
        total_camera = len(camera_images)
        early_stop_threshold = distance_threshold * 0.7 if fast_mode else None

        for i, camera_img in enumerate(camera_images):
            if not os.path.exists(camera_img):
                continue

            if i % 5 == 0 or i == total_camera - 1:
                self._log(f"    [SCAN] So sánh ảnh {i+1}/{total_camera}...", "default")

            try:
                cam_emb = self._get_embedding(camera_img)
                if cam_emb is None:
                    continue

                distances = []
                for _, p_emb in portrait_embeddings:
                    if self.distance_metric == "cosine":
                        d = self._cosine_distance(cam_emb, p_emb)
                    else:
                        d = float(np.linalg.norm(cam_emb - p_emb))
                    distances.append(d)

                if not distances:
                    continue

                distance = min(distances)
                if log_detail:
                    self._log(
                        f"    [DIST] {os.path.basename(camera_img)} => {distance:.3f}",
                        "default"
                    )
                compared_count += 1

                if distance < best_distance:
                    best_distance = distance
                    best_match = camera_img
                    self._log(
                        f"    [CAND] Ứng viên: {os.path.basename(camera_img)} (distance={distance:.3f})",
                        "info"
                    )

                if early_stop_threshold is not None and best_distance <= early_stop_threshold:
                    self._log(
                        f"    [EARLY] Match tốt tìm thấy sớm! (distance={best_distance:.3f})",
                        "success"
                    )
                    break
            except Exception as e:
                errors_count += 1
                if errors_count <= 3:
                    self._log(f"    [WARN] Error #{errors_count}: {str(e)}", "warning")

        self._log(
            f"  [STATS] So sánh: {compared_count}/{total_camera} ảnh, lỗi: {errors_count}",
            "info"
        )

        if best_match and best_distance <= distance_threshold:
            self._log(
                f"  [OK] Best Match: {os.path.basename(best_match)} (distance={best_distance:.3f})",
                "success"
            )
            return best_match
        elif best_match:
            self._log(
                f"  -> Best distance={best_distance:.3f} > threshold={distance_threshold}",
                "warning"
            )
        else:
            self._log("  -> Không tìm thấy ảnh nào match được", "error")

        return None

    def match_all_faces(self, person_name: str, camera_images: List[str], 
                        max_matches: int = 1) -> List[Tuple[str, float]]:
        """
        Tìm tất cả ảnh camera match với người được chỉ định

        Returns:
            List of (image_path, confidence) tuples
        """
        portrait_paths = self.find_portraits(person_name)
        if not portrait_paths:
            return []

        portrait_embeddings = []
        for p in portrait_paths:
            emb = self._get_embedding(p)
            if emb is not None:
                portrait_embeddings.append(emb)

        if not portrait_embeddings:
            return []

        matches = []
        for camera_img in camera_images:
            cam_emb = self._get_embedding(camera_img)
            if cam_emb is None:
                continue

            distances = []
            for p_emb in portrait_embeddings:
                if self.distance_metric == "cosine":
                    d = self._cosine_distance(cam_emb, p_emb)
                else:
                    d = float(np.linalg.norm(cam_emb - p_emb))
                distances.append(d)

            if not distances:
                continue

            distance = min(distances)
            confidence = max(0, 100 * (1 - distance))
            matches.append((camera_img, confidence))

            if len(matches) >= max_matches:
                break

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
            portrait_ascii = copy_to_ascii_path(portrait_path)
            camera_ascii = copy_to_ascii_path(camera_img)
            result = DeepFace.verify(
                img1_path=portrait_ascii,
                img2_path=camera_ascii,
                model_name="ArcFace",
                detector_backend="retinaface",
                enforce_detection=True
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




