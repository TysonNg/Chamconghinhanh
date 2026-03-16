# -*- coding: utf-8 -*-
"""
Module nháº­n diá»‡n khuÃ´n máº·t - so sÃ¡nh áº£nh camera vá»›i áº£nh chÃ¢n dung
"""

import os
import re
import shutil
import tempfile
import unicodedata
from typing import List, Dict, Optional, Tuple
import numpy as np

# Lazy loading Ä‘á»ƒ trÃ¡nh import lá»—i
_deepface = None

# ThÆ° má»¥c temp Ä‘á»ƒ lÆ°u áº£nh táº¡m (trÃ¡nh lá»—i Ä‘Æ°á»ng dáº«n tiáº¿ng Viá»‡t)
_temp_dir = None

def get_temp_dir():
    """Láº¥y hoáº·c táº¡o thÆ° má»¥c temp"""
    global _temp_dir
    if _temp_dir is None or not os.path.exists(_temp_dir):
        _temp_dir = tempfile.mkdtemp(prefix="face_matcher_")
    return _temp_dir

def copy_to_ascii_path(src_path: str) -> str:
    """
    Copy file sang thÆ° má»¥c temp vá»›i tÃªn ASCII
    Giáº£i quyáº¿t lá»—i DeepFace khÃ´ng Ä‘á»c Ä‘Æ°á»£c Ä‘Æ°á»ng dáº«n tiáº¿ng Viá»‡t
    """
    if not os.path.exists(src_path):
        print(f"  [copy_to_ascii] Source khÃ´ng tá»“n táº¡i: {src_path}")
        return src_path
    
    # LuÃ´n copy Ä‘á»ƒ trÃ¡nh lá»—i encoding
    try:
        # Táº¡o tÃªn file ASCII dá»±a trÃªn hash
        ext = os.path.splitext(src_path)[1]
        # DÃ¹ng abs path Ä‘á»ƒ hash unique hÆ¡n
        abs_path = os.path.abspath(src_path)
        ascii_name = f"img_{hash(abs_path) & 0xFFFFFFFF}{ext}"
        
        temp_dir = get_temp_dir()
        dst_path = os.path.join(temp_dir, ascii_name)
        
        # Copy náº¿u chÆ°a tá»“n táº¡i hoáº·c source má»›i hÆ¡n
        if (not os.path.exists(dst_path) or
                os.path.getmtime(src_path) > os.path.getmtime(dst_path)):
            shutil.copy2(src_path, dst_path)
        
        return dst_path
    except Exception as e:
        print(f"  [copy_to_ascii] Error: {e}")
        return src_path

def get_deepface():
    """Lazy load DeepFace Ä‘á»ƒ giáº£m thá»i gian khá»Ÿi Ä‘á»™ng"""
    global _deepface
    if _deepface is None:
        try:
            from deepface import DeepFace
            _deepface = DeepFace
        except ImportError:
            print("DeepFace chÆ°a Ä‘Æ°á»£c cÃ i Ä‘áº·t. Cháº¡y: pip install deepface tf-keras")
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
    TÃ­nh Ä‘á»™ tÆ°Æ¡ng Ä‘á»“ng giá»¯a 2 tÃªn (0.0 - 1.0)
    Sá»­ dá»¥ng thuáº­t toÃ¡n Ä‘Æ¡n giáº£n dá»±a trÃªn substring matching
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
    """So sÃ¡nh khuÃ´n máº·t giá»¯a áº£nh camera vÃ  áº£nh chÃ¢n dung"""
    
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
        self._embedding_cache = {}  # {path: (mtime, embedding)}
        self.log_callback = log_callback
        self._scan_portraits()

    def _log(self, message: str, log_type: str = "default"):
        """Gá»­i log qua callback hoáº·c print"""
        if self.log_callback:
            self.log_callback(message, log_type)
        print(message)  # Always print to console too
    
    def _scan_portraits(self):
        """QuÃ©t thÆ° má»¥c chÃ¢n dung vÃ  cache Ä‘Æ°á»ng dáº«n"""
        if not os.path.exists(self.portrait_dir):
            print(f"ThÆ° má»¥c áº£nh chÃ¢n dung khÃ´ng tá»“n táº¡i: {self.portrait_dir}")
            return
        
        supported_ext = {'.jpg', '.jpeg', '.png', '.bmp'}
        
        for item in os.listdir(self.portrait_dir):
            item_path = os.path.join(self.portrait_dir, item)
            
            if os.path.isdir(item_path):
                # ThÆ° má»¥c con = tÃªn ngÆ°á»i
                person_name = item
                images = []
                for file in os.listdir(item_path):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in supported_ext:
                        images.append(os.path.join(item_path, file))
                
                if images:
                    self.portrait_cache[person_name] = images
            else:
                # File trá»±c tiáº¿p = tÃªn file lÃ  tÃªn ngÆ°á»i
                ext = os.path.splitext(item)[1].lower()
                if ext in supported_ext:
                    person_name = os.path.splitext(item)[0]
                    self.portrait_cache[person_name] = [item_path]
        
        print(f"ÄÃ£ load {len(self.portrait_cache)} ngÆ°á»i tá»« thÆ° má»¥c chÃ¢n dung")
    
    def find_portrait(self, person_name: str) -> Optional[str]:
        """TÃ¬m áº£nh chÃ¢n dung Ä‘áº§u tiÃªn cho má»™t ngÆ°á»i (backward compatible)"""
        portraits = self.find_portraits(person_name)
        return portraits[0] if portraits else None
    
    def find_portraits(self, person_name: str) -> List[str]:
        """
        TÃ¬m Táº¤T Cáº¢ áº£nh chÃ¢n dung cho má»™t ngÆ°á»i
        Há»— trá»£ matching tÃªn tiáº¿ng Viá»‡t cÃ³/khÃ´ng dáº¥u
        """
        # 1. Exact match
        if person_name in self.portrait_cache:
            return self.portrait_cache[person_name]
        
        # 2. Normalize vÃ  tÃ¬m exact match sau khi chuáº©n hÃ³a
        person_normalized = normalize_vietnamese(person_name)
        
        for cached_name, images in self.portrait_cache.items():
            cached_normalized = normalize_vietnamese(cached_name)
            
            # Exact match sau khi normalize
            if person_normalized == cached_normalized:
                return images
        
        # 3. Fuzzy match vá»›i similarity score
        best_images = None
        best_score = 0.0
        min_threshold = 0.7  # YÃªu cáº§u Ã­t nháº¥t 70% tÆ°Æ¡ng Ä‘á»“ng
        
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
    
    def _cosine_distance(self, a: np.ndarray, b: np.ndarray) -> float:
        denom = (np.linalg.norm(a) * np.linalg.norm(b)) + 1e-12
        return float(1.0 - np.dot(a, b) / denom)

    def _get_embedding(self, image_path: str) -> Optional[np.ndarray]:
        DeepFace = get_deepface()
        if DeepFace is None:
            return None

        if not os.path.exists(image_path):
            return None

        try:
            mtime = os.path.getmtime(image_path)
            cached = self._embedding_cache.get(image_path)
            if cached and cached[0] == mtime:
                return cached[1]

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
            self._embedding_cache[image_path] = (mtime, embedding)
            return embedding
        except Exception:
            return None

    def _get_default_threshold(self) -> float:
        if self.distance_metric == "cosine":
            model = self.model_name.lower()
            if model == "arcface":
                return 0.35
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
        log_detail: bool = False
    ) -> Optional[str]:
        """
        Tìm ảnh camera có khuôn mặt match với người được chỉ định

        Args:
            person_name: Tên người cần tìm
            camera_images: Danh sách đường dẫn ảnh camera
            distance_threshold: Ngưỡng khoảng cách (thấp hơn = giống hơn).
                                Nếu None sẽ dùng ngưỡng mặc định theo model.

        Returns:
            Đường dẫn ảnh camera match tốt nhất, hoặc None nếu không tìm thấy
        """
        if distance_threshold is None:
            distance_threshold = self._get_default_threshold()

        # Tìm tất cả ảnh chân dung
        portrait_paths = self.find_portraits(person_name)
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
    So sÃ¡nh Ä‘Æ¡n giáº£n - tráº£ vá» áº£nh camera Ä‘áº§u tiÃªn match vá»›i portrait
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
    
    portrait_dir = r'd:\Projects\phan mem quet mat\áº¢nh BV'
    matcher = FaceMatcher(portrait_dir)
    
    print("\n=== Test FaceMatcher ===")
    print(f"Sá»‘ ngÆ°á»i trong cache: {len(matcher.portrait_cache)}")
    
    # Test tÃ¬m portrait
    test_name = "Nguyen Van A"
    portrait = matcher.find_portrait(test_name)
    print(f"Portrait cho {test_name}: {portrait}")




