# -*- coding: utf-8 -*-
"""
Smart Watermark Replacer - Thay thế text timestamp thông minh & chân thật
Giữ nguyên 100% ảnh gốc, phát hiện bằng EasyOCR, xóa nét chữ siêu mỏng
(không xóa nhòe cả mảng nền) và vẽ lại với bóng đổ + viền chuẩn camera.
Tích hợp Google Gemini & OpenAI API để phân tích style chữ chính xác.
"""

import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter

from src.ai_timestamp_service import AITimestampService

logger = logging.getLogger(__name__)

# ==================== LAZY SINGLETON OCR READER ====================

_ocr_reader = None
_ocr_init_error = None


def _get_ocr_reader():
    """Lazy-load EasyOCR Reader (chỉ init 1 lần, cache lại)."""
    global _ocr_reader, _ocr_init_error
    if _ocr_reader is not None:
        return _ocr_reader
    if _ocr_init_error is not None:
        return None
    try:
        import easyocr
        _ocr_reader = easyocr.Reader(['vi', 'en'], gpu=False, verbose=False)
        logger.info("EasyOCR Reader initialized successfully (vi, en)")
        return _ocr_reader
    except Exception as e:
        _ocr_init_error = str(e)
        logger.error(f"Không thể khởi tạo EasyOCR: {e}")
        return None


# ==================== HELPER UTF-8 FILE I/O FOR OPENCV ====================

def imread_utf8(path: str) -> Optional[np.ndarray]:
    """Đọc ảnh hỗ trợ đường dẫn tiếng Việt Unicode trên Windows."""
    try:
        with open(path, 'rb') as f:
            bytes_data = np.frombuffer(f.read(), dtype=np.uint8)
            return cv2.imdecode(bytes_data, cv2.IMREAD_COLOR)
    except Exception as e:
        logger.error(f"Lỗi imread_utf8({path}): {e}")
        return None


def imwrite_utf8(path: str, img: np.ndarray, quality: int = 95) -> bool:
    """Ghi ảnh hỗ trợ đường dẫn tiếng Việt Unicode trên Windows."""
    try:
        ext = Path(path).suffix.lower()
        if not ext:
            ext = '.jpg'
        params = [int(cv2.IMWRITE_JPEG_QUALITY), quality] if ext in ('.jpg', '.jpeg') else []
        success, encoded = cv2.imencode(ext, img, params)
        if success:
            with open(path, 'wb') as f:
                f.write(encoded.tobytes())
            return True
        return False
    except Exception as e:
        logger.error(f"Lỗi imwrite_utf8({path}): {e}")
        return False


def hex_to_rgb(hex_str: str, default: Tuple[int, int, int] = (255, 255, 255)) -> Tuple[int, int, int]:
    """Chuyển đổi màu Hex (#FFFFFF hoặc FFFFFF) sang RGB tuple."""
    if not hex_str or not isinstance(hex_str, str):
        return default
    hex_clean = hex_str.strip().lstrip('#')
    if len(hex_clean) == 3:
        hex_clean = ''.join(c * 2 for c in hex_clean)
    if len(hex_clean) == 6:
        try:
            return (int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))
        except ValueError:
            pass
    return default


# ==================== DATA CLASSES ====================

@dataclass
class TimestampRegion:
    """Thông tin một vùng text timestamp được phát hiện."""
    bbox: List[List[int]]         # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
    text: str                     # Nội dung text gốc
    confidence: float             # Độ tin cậy OCR
    detected_color: Tuple[int, int, int] = (255, 255, 255)  # RGB
    estimated_font_size: int = 20
    has_date: bool = False
    has_time: bool = False
    is_weekday: bool = False
    is_timemark: bool = False
    timemark_info: Optional[Dict[str, Any]] = None
    date_format: str = ""         # "DD/MM/YYYY", "YYYY-MM-DD", etc.
    line_index: int = 0           # Thứ tự dòng trong block
    ref_right_x: Optional[int] = None  # Tọa độ lề phải tham chiếu từ các dòng địa chỉ bên dưới


# ==================== MAIN CLASS ====================

class SmartWatermarkReplacer:
    """
    Thay thế text timestamp trên ảnh chấm công.
    Giữ nguyên 100% ảnh gốc (mặt, nền), chỉ thay đổi text ngày/giờ.
    Sử dụng xóa nét chữ siêu mỏng và kết hợp AI Vision để khớp màu/bóng đổ.
    """

    VN_WEEKDAYS = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật']
    WEEKDAY_PATTERN = re.compile(r'\b(Thứ\s*(?:Hai|Ba|Tư|Bốn|Năm|Sáu|Bảy|2|3|4|5|6|7|[^\s,]+)|Chủ\s*Nhật|CN)\b', re.I)

    # Regex patterns cho ngày
    DATE_PATTERNS = [
        # Format tiếng Anh: "18 Jan 2026", "18 January 2026", "Jan 18, 2026"
        (re.compile(r'(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})', re.I), 'DD_ENG_YYYY'),
        (re.compile(r'(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})', re.I), 'ENG_DD_YYYY'),
        # DD/MM/YYYY hoặc DD-MM-YYYY hoặc DD.MM.YYYY
        (re.compile(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})'), 'DD/MM/YYYY'),
        # YYYY/MM/DD hoặc YYYY-MM-DD
        (re.compile(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})'), 'YYYY/MM/DD'),
        # Tháng viết dạng tiếng Việt: "11 Thl, 2026", "14 Th9, 2026", "15 Thg 9, 2026", "15 Tháng 9, 2026", "16 Th09, 2026", "(hang 2,2026", v.v.
        (re.compile(r'(\d{1,2})\s*(?:[\(tT][hH]?[aáàạảãAÁÀẠẢÃ]?[nN]?[gG]?|th[a-zA-Z0-9\.]*|tháng)\s*(\d{1,2})\s*,?\s*(\d{4})', re.I), 'DD_THG_MM_YYYY'),
        (re.compile(r'(?:[\(tT][hH]?[aáàạảãAÁÀẠẢÃ]?[nN]?[gG]?|tháng|thg)\s*(\d{1,2})\s*,?\s*(\d{4})', re.I), 'THG_MM_YYYY'),
        (re.compile(r'(\d{1,2})\s*(?:th[a-zA-Z0-9\.]*(?:\s*\d+)?|tháng\s*\d*)\s*,?\s*(\d{4})', re.I), 'DD_THG_MM_YYYY'),
    ]

    @staticmethod
    def is_in_safe_zone(bbox: List[List[int]], h: int) -> bool:
        """
        Watermark chấm công thực tế luôn nằm ở viền ngoài:
        - Top safe zone: y < 22% chiều cao ảnh (ví dụ: Timestamp Camera góc trên)
        - Bottom safe zone: y > 75% chiều cao ảnh (ví dụ: GPS Map Camera góc dưới)
        Loại bỏ hoàn toàn vùng giữa (22% - 75%) nơi có khuôn mặt, đồng phục, bảng tên, biển số xe.
        """
        if not bbox or h <= 0:
            return False
        y_coords = [p[1] for p in bbox]
        if not y_coords:
            return False
        mid_y = sum(y_coords) / len(y_coords)
        return (mid_y < 0.22 * h) or (mid_y > 0.75 * h)

    # Regex patterns cho giờ (ví dụ: 07:46-51, 07:46:54, 20.52.10,931, 08:00, 21.-00)
    TIME_PATTERN = re.compile(r'(\d{1,2})[:.\-h]+(\d{2})(?:[:.\-h]+(\d{2})(?:,\d+)?)?')

    # Font paths phổ biến - Ưu tiên Roboto (chuẩn tuyệt đối của Timestamp Camera di động)
    _ROBOTO_STATIC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'static', 'fonts', 'Roboto.ttf')
    _ROBOTO_INTERNAL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '_internal', 'static', 'fonts', 'Roboto.ttf')

    FONT_PATHS = [
        _ROBOTO_STATIC,
        _ROBOTO_INTERNAL,
        r"C:\Windows\Fonts\segoeui.ttf",   # Segoe UI Regular
        r"C:\Windows\Fonts\arial.ttf",     # Arial Regular
        r"C:\Windows\Fonts\segoeuib.ttf",  # Segoe UI Bold
        r"C:\Windows\Fonts\arialbd.ttf",   # Arial Bold
        r"C:\Windows\Fonts\calibrib.ttf",  # Calibri Bold
        r"C:\Windows\Fonts\tahoma.ttf",
    ]

    def __init__(self, ai_service: Optional[AITimestampService] = None):
        self._font_path = self._find_font()
        self.ai_service = ai_service

    def _find_font(self) -> Optional[str]:
        """Tìm font có sẵn trên hệ thống Windows."""
        for path in self.FONT_PATHS:
            if os.path.exists(path):
                return path
        return None

    # ==================== 1. DETECT TIMESTAMP REGIONS ====================

    def detect_timestamp_regions(self, image_path: str) -> List[TimestampRegion]:
        """
        Dùng EasyOCR scan toàn ảnh, tìm tất cả vùng text chứa ngày/giờ.
        Phát hiện và gom nhóm thông minh các khối Timemark (Giờ lớn | Ngày / Thứ).
        Trả về danh sách TimestampRegion.
        """
        reader = _get_ocr_reader()
        if reader is None:
            logger.warning("EasyOCR không khả dụng, bỏ qua phát hiện timestamp")
            return []

        # Đọc ảnh an toàn qua UTF-8
        img = imread_utf8(image_path)
        if img is None:
            logger.error(f"Không thể đọc ảnh: {image_path}")
            return []
        h, w = img.shape[:2]

        try:
            results = reader.readtext(img, detail=1)
        except Exception as e:
            logger.error(f"Lỗi OCR: {e}")
            return []

        if not results:
            logger.info("OCR không phát hiện text nào trong ảnh")
            return []

        # Tìm lề phải chuẩn từ các dòng địa chỉ hoặc thông tin camera bên dưới
        ref_right_x = None
        bottom_right_xs = []
        is_timemark_logo = False

        for bbox, text, _ in results:
            text_str = str(text).strip()
            if re.search(r'\b(timemark|toonemark|100%|chân thực|chan thuc)\b', text_str, re.I):
                is_timemark_logo = True
            bx = [p[0] for p in bbox]
            by = [p[1] for p in bbox]
            if min(by) > h * 0.70 and max(bx) > w * 0.60:
                bottom_right_xs.append(max(bx))
        if bottom_right_xs:
            ref_right_x = int(max(bottom_right_xs))

        detected_candidates: List[TimestampRegion] = []
        for idx, (bbox, text, confidence) in enumerate(results):
            text_str = str(text).strip()
            if not text_str:
                continue

            bbox_int = [[int(round(p[0])), int(round(p[1]))] for p in bbox]

            # 1. Kiểm tra Safe Zone: Chặn tuyệt đối vùng giữa ảnh (22% - 75% chiều cao ảnh)
            if not self.is_in_safe_zone(bbox_int, h):
                continue

            # Bỏ qua logo watermark app góc dưới
            if re.search(r'\b(timemark|toonemark|100%|chân thực|chan thuc)\b', text_str, re.I):
                continue

            # 2. Kiểm tra xem text có chứa date, time hoặc weekday pattern không
            has_date = False
            has_time = False
            is_weekday = False
            date_format = ""

            for pattern, fmt in self.DATE_PATTERNS:
                if pattern.search(text_str):
                    has_date = True
                    date_format = fmt
                    break

            if self.TIME_PATTERN.search(text_str):
                has_time = True

            if self.WEEKDAY_PATTERN.search(text_str) or re.search(r'\b(thứ|thu|chủ nhật)\b', text_str, re.I):
                is_weekday = True

            # Chỉ giữ lại text có chứa ngày, giờ hoặc thứ
            if not has_date and not has_time and not is_weekday:
                continue

            # 3. Bỏ qua nếu text khớp dạng biển số xe Việt Nam
            if re.search(r'\b\d{2}[-–\s]*[A-Z]{1,2}\d?[-–\s.]*\d{3,5}\b', text_str, re.I) and not has_date:
                continue

            color, font_size = self._analyze_text_style(img, bbox_int)

            region = TimestampRegion(
                bbox=bbox_int,
                text=text_str,
                confidence=confidence,
                detected_color=color,
                estimated_font_size=font_size,
                has_date=has_date,
                has_time=has_time,
                is_weekday=is_weekday,
                date_format=date_format,
                line_index=idx,
                ref_right_x=ref_right_x
            )
            detected_candidates.append(region)

        # 4. Phát hiện và nhóm layout Timemark 2 cột (Giờ lớn | Ngày / Thứ)
        time_items = [r for r in detected_candidates if r.has_time and (r.bbox[2][1] - r.bbox[0][1] >= 35)]
        date_items = [r for r in detected_candidates if r.has_date]
        weekday_items = [r for r in detected_candidates if r.is_weekday]

        is_timemark = False
        timemark_region = None

        if time_items and (date_items or weekday_items or is_timemark_logo):
            # Kiểm tra xem time và date/weekday có nằm cạnh nhau theo chiều ngang không
            t_reg = time_items[0]
            t_x2 = max(p[0] for p in t_reg.bbox)
            t_y1 = min(p[1] for p in t_reg.bbox)
            t_y2 = max(p[1] for p in t_reg.bbox)

            adjacent_right = [
                r for r in (date_items + weekday_items)
                if min(p[0] for p in r.bbox) >= t_x2 - 10
                and abs(min(p[1] for p in r.bbox) - t_y1) < (t_y2 - t_y1) * 1.5
            ]

            # Quét tất cả các mảnh text OCR trong dải ngang bên phải giờ của Timemark
            t_x1 = min(p[0] for p in t_reg.bbox)
            extra_right_boxes = []
            for b, t, c in results:
                bx = [int(p[0]) for p in b]
                by = [int(p[1]) for p in b]
                min_bx, max_bx = min(bx), max(bx)
                min_by, max_by = min(by), max(by)
                if min_bx >= t_x2 - 15 and max_bx <= t_x1 + 560 and min_by >= t_y1 - 20 and max_by <= t_y2 + 25:
                    extra_right_boxes.append([[min_bx, min_by], [max_bx, min_by], [max_bx, max_by], [min_bx, max_by]])

            if adjacent_right or is_timemark_logo or extra_right_boxes:
                is_timemark = True
                all_boxes = [t_reg.bbox] + [r.bbox for r in adjacent_right] + extra_right_boxes
                all_x = [p[0] for b in all_boxes for p in b]
                all_y = [p[1] for b in all_boxes for p in b]

                pad = 4
                union_bbox = [
                    [max(0, min(all_x) - pad), max(0, min(all_y) - pad)],
                    [min(w, max(all_x) + pad), max(0, min(all_y) - pad)],
                    [min(w, max(all_x) + pad), min(h, max(all_y) + pad)],
                    [max(0, min(all_x) - pad), min(h, max(all_y) + pad)]
                ]

                # Tọa độ time_box chuẩn
                time_box = t_reg.bbox
                d_box = date_items[0].bbox if date_items else None
                w_box = weekday_items[0].bbox if weekday_items else None

                timemark_info = {
                    'time_box': time_box,
                    'date_box': d_box,
                    'weekday_box': w_box,
                    'time_text': t_reg.text,
                    'date_text': date_items[0].text if date_items else '',
                    'weekday_text': weekday_items[0].text if weekday_items else '',
                    'separator_color': (212, 168, 75),
                }

                timemark_region = TimestampRegion(
                    bbox=union_bbox,
                    text=f"{t_reg.text} | {timemark_info['date_text']} {timemark_info['weekday_text']}".strip(),
                    confidence=t_reg.confidence,
                    detected_color=(255, 255, 255),
                    estimated_font_size=t_reg.estimated_font_size,
                    has_date=True,
                    has_time=True,
                    is_weekday=True,
                    is_timemark=True,
                    timemark_info=timemark_info,
                    ref_right_x=ref_right_x
                )

        if is_timemark and timemark_region:
            # Loại bỏ các region con đã được gộp vào Timemark
            used_ids = set([id(t_reg)] + [id(r) for r in adjacent_right])
            other_regions = [r for r in detected_candidates if id(r) not in used_ids]
            timestamp_regions = [timemark_region] + other_regions
            logger.info(f"Đã phát hiện Timemark 2-column layout: time='{timemark_region.text}'")
        else:
            timestamp_regions = detected_candidates

        logger.info(f"Phát hiện {len(timestamp_regions)} vùng timestamp trong ảnh")
        return timestamp_regions

    # ==================== 2. ANALYZE TEXT STYLE ====================

    def _analyze_text_style(
        self, img: np.ndarray, bbox: List[List[int]]
    ) -> Tuple[Tuple[int, int, int], int]:
        """Phân tích màu sắc và kích thước font từ vùng text trong ảnh."""
        h, w = img.shape[:2]
        x_coords = [p[0] for p in bbox]
        y_coords = [p[1] for p in bbox]
        x1 = max(0, min(x_coords))
        y1 = max(0, min(y_coords))
        x2 = min(w, max(x_coords))
        y2 = min(h, max(y_coords))

        if x2 <= x1 or y2 <= y1:
            return (255, 255, 255), 20

        roi = img[y1:y2, x1:x2]
        font_size = max(14, int((y2 - y1) * 0.70))

        color = self._detect_dominant_text_color(roi)
        return color, font_size

    def _detect_dominant_text_color(self, roi: np.ndarray) -> Tuple[int, int, int]:
        """Phát hiện màu text chủ đạo trong vùng ROI (thường là màu sáng trắng/vàng)."""
        if roi.size == 0:
            return (255, 255, 255)

        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        # Các camera timestamp hầu như luôn dùng màu trắng sáng (> 180)
        bright_pixels = roi[gray >= 180]
        if len(bright_pixels) > 10:
            mean_bgr = np.mean(bright_pixels, axis=0).astype(int)
            return (int(mean_bgr[2]), int(mean_bgr[1]), int(mean_bgr[0]))

        # Nếu không có pixel quá sáng, lấy top 15% pixel sáng nhất
        thresh_val = np.percentile(gray, 85)
        top_pixels = roi[gray >= thresh_val]
        if len(top_pixels) > 0:
            mean_bgr = np.mean(top_pixels, axis=0).astype(int)
            return (int(mean_bgr[2]), int(mean_bgr[1]), int(mean_bgr[0]))

        return (255, 255, 255)

    # ==================== 3. TIGHT CHARACTER-STROKE MASK ====================

    def _build_text_mask(
        self,
        img: np.ndarray,
        bbox: List[List[int]],
        padding: int = 2,
        detected_color: Optional[Tuple[int, int, int]] = None,
    ) -> np.ndarray:
        """
        Tạo mask bao phủ nét chữ sáng hoặc tối và bóng đổ của watermark camera.
        Sử dụng phân tích ngưỡng sáng cục bộ và dải viền ký tự (candidate band):
        - Nhận diện lõi chữ trắng sáng (gray >= 165)
        - Mở rộng dải ứng viên chứa ký tự và bóng đổ (ellipse 7x7)
        - Trên nền sáng (đường, vỉa hè, tường): bóc tách sạch bóng đổ đen (absdiff > 10 so với nền cục bộ)
        - Trên nền tối (quần, áo tối màu): bóc tách sạch ký tự và viền khử răng cưa (gray > 45)
        - Giãn viền mượt mà (ellipse 5x5) để triệt tiêu hoàn toàn bóng ma 100%.
        """
        h, w = img.shape[:2]
        mask = np.zeros((h, w), dtype=np.uint8)

        x_coords = [p[0] for p in bbox]
        y_coords = [p[1] for p in bbox]
        pad = max(padding, 2)
        x1 = max(0, min(x_coords) - pad)
        y1 = max(0, min(y_coords) - pad)
        x2 = min(w, max(x_coords) + pad)
        y2 = min(h, max(y_coords) + pad)

        roi = img[y1:y2, x1:x2]
        if roi.size == 0:
            return mask

        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

        # 1. Ước lượng nền cục bộ
        k_size = 25 if min(roi.shape[:2]) >= 25 else (min(roi.shape[:2]) // 2 * 2 + 1)
        k_size = max(3, k_size)
        bg_large = cv2.medianBlur(gray, k_size)

        color_luma = 255.0
        if detected_color is not None:
            red, green, blue = detected_color
            color_luma = 0.299 * red + 0.587 * green + 0.114 * blue

        if detected_color is not None and color_luma < 130:
            # Watermark tối: bám theo màu AI/OCR đã phát hiện và tương phản âm với nền.
            expected_bgr = np.array(
                [detected_color[2], detected_color[1], detected_color[0]],
                dtype=np.int16,
            )
            color_distance = np.linalg.norm(
                roi.astype(np.int16) - expected_bgr,
                axis=2,
            )
            dark_delta = bg_large.astype(np.int16) - gray.astype(np.int16)
            dark_core = (dark_delta >= 12) & (color_distance <= 75)

            # Chữ xanh đậm có thể nằm trên mái cũng tối tương đương về luminance.
            # Khi màu đích có kênh trội rõ ràng, dùng độ trội màu để lấy lại các nét đó.
            strongest = int(np.argmax(expected_bgr))
            weakest = int(np.argmin(expected_bgr))
            middle = 3 - strongest - weakest
            chroma_span = int(expected_bgr[strongest] - expected_bgr[weakest])
            middle_span = int(expected_bgr[strongest] - expected_bgr[middle])
            if chroma_span >= 20:
                roi_i16 = roi.astype(np.int16)
                dominant_delta = roi_i16[:, :, strongest] - roi_i16[:, :, weakest]
                middle_delta = roi_i16[:, :, strongest] - roi_i16[:, :, middle]
                chroma_core = (
                    (dominant_delta >= max(18, int(chroma_span * 0.45)))
                    & (middle_delta >= max(8, int(middle_span * 0.40)))
                )
                dark_core |= chroma_core
            if not np.any(dark_core):
                return mask

            candidate_band = cv2.dilate(
                dark_core.astype(np.uint8),
                cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7)),
                iterations=1,
            )
            dark_edges = (dark_delta >= 5) & (color_distance <= 105)
            if chroma_span >= 20:
                dark_edges |= (
                    (dominant_delta >= max(14, int(chroma_span * 0.35)))
                    & (middle_delta >= max(6, int(middle_span * 0.30)))
                )
            mask_roi = (candidate_band == 1) & dark_edges
        else:
            # Watermark sáng: giữ nguyên chiến lược lõi trắng hiện có.
            white_core = (gray >= 165)
            kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
            candidate_band = cv2.dilate(white_core.astype(np.uint8), kernel, iterations=1)

            # Phân tách nền sáng và nền tối cho chữ trắng.
            diff_light = (bg_large > 80) & (cv2.absdiff(gray, bg_large) > 10)
            diff_dark = (bg_large <= 80) & (gray > 45)
            mask_roi = (candidate_band == 1) & (diff_light | diff_dark)

        char_mask_dilated = cv2.dilate(
            mask_roi.astype(np.uint8) * 255,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5)),
            iterations=1
        )

        # Giới hạn mask bên trong polygon bbox (có padding an toàn)
        poly_mask = np.zeros((h, w), dtype=np.uint8)
        pts = np.array([[x1, y1], [x2, y1], [x2, y2], [x1, y2]], dtype=np.int32)
        cv2.fillPoly(poly_mask, [pts], 255)

        roi_target = mask[y1:y2, x1:x2]
        mask[y1:y2, x1:x2] = cv2.bitwise_or(roi_target, char_mask_dilated)

        final_mask = cv2.bitwise_and(mask, poly_mask)
        if detected_color is not None and color_luma < 130:
            logger.info(
                "Dark watermark mask: color=%s, bbox=[%s,%s,%s,%s], pixels=%s",
                detected_color,
                x1, y1, x2, y2,
                int(np.count_nonzero(final_mask)),
            )
        return final_mask

    def _build_timemark_mask(
        self, img: np.ndarray, region: TimestampRegion, padding: int = 3
    ) -> np.ndarray:
        """Tạo mask thích ứng cho nét sáng, vạch vàng và bóng tối của Timemark."""
        h, w = img.shape[:2]
        mask = np.zeros((h, w), dtype=np.uint8)

        x_coords = [p[0] for p in region.bbox]
        y_coords = [p[1] for p in region.bbox]
        region_h = max(y_coords) - min(y_coords)
        adaptive_pad = int(round(region_h * 0.08))
        pad = max(padding, min(14, max(3, adaptive_pad)))
        x1 = max(0, min(x_coords) - pad)
        y1 = max(0, min(y_coords) - pad)
        x2 = min(w, max(x_coords) + pad)
        y2 = min(h, max(y_coords) + pad)

        roi = img[y1:y2, x1:x2]
        if roi.size == 0:
            return mask

        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)

        # So sánh với nền cục bộ để không coi toàn bộ nền sáng là chữ.
        local_kernel = min(31, max(9, int(round(region_h * 0.32))))
        if local_kernel % 2 == 0:
            local_kernel += 1
        local_background = cv2.medianBlur(gray, local_kernel)
        signed_delta = gray.astype(np.int16) - local_background.astype(np.int16)

        # Lõi chữ phải vừa sáng vừa nổi bật so với nền lân cận.
        white_core = (gray >= 165) & (signed_delta >= 10)
        # Vạch ngăn vàng/cam
        yellow = (hsv[:, :, 0] >= 12) & (hsv[:, :, 0] <= 48) & (hsv[:, :, 1] >= 60) & (hsv[:, :, 2] >= 100)

        combined_core = (white_core | yellow).astype(np.uint8) * 255
        if not np.any(combined_core):
            return mask

        # Chỉ xét bóng tối trong dải gần nét thật; chi tiết nền ở xa không bị mask.
        shadow_radius = min(10, max(4, int(round(region_h * 0.11))))
        band_size = shadow_radius * 2 + 1
        shadow_band = cv2.dilate(
            combined_core,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (band_size, band_size)),
            iterations=1,
        )
        dark_contrast = (signed_delta <= -12) & (gray <= 150)
        shadow = dark_contrast & (shadow_band > 0)

        combined = cv2.bitwise_or(combined_core, shadow.astype(np.uint8) * 255)
        close_size = min(5, max(3, int(round(region_h * 0.045)) | 1))
        combined = cv2.morphologyEx(
            combined,
            cv2.MORPH_CLOSE,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (close_size, close_size)),
        )

        # Nới nhẹ để phủ anti-alias nhưng không lan ra toàn khối OCR.
        edge_size = min(5, max(3, int(round(region_h * 0.04)) | 1))
        combined = cv2.dilate(
            combined,
            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (edge_size, edge_size)),
            iterations=1,
        )

        mask[y1:y2, x1:x2] = combined
        return mask

    @staticmethod
    def _timemark_inpaint_radius(mask: np.ndarray) -> int:
        """Suy ra bán kính inpaint từ độ dày mask, có giới hạn để bảo toàn nền."""
        if mask.size == 0 or not np.any(mask):
            return 2
        thickness = float(cv2.distanceTransform(mask, cv2.DIST_L2, 5).max())
        return min(5, max(2, int(round(thickness * 0.65))))

    def _render_timemark_block(
        self,
        pil_img: Image.Image,
        region: TimestampRegion,
        target_date: str,
        target_time: str,
    ) -> Image.Image:
        """
        Render khối watermark Timemark 2 cột chuẩn:
        - Cột trái: Giờ phút kích thước lớn (Roboto scaled, viền/stroke tự nhiên)
        - Giữa: Vạch ngăn màu vàng/cam (gold separator line)
        - Cột phải dòng 1: Ngày tháng năm tiếng Việt (DD Tháng MM, YYYY)
        - Cột phải dòng 2: Thứ trong tuần tiếng Việt (Thứ Hai .. Chủ Nhật)
        """
        info = region.timemark_info or {}
        time_bbox = info.get('time_box')
        if not time_bbox:
            x1 = min(p[0] for p in region.bbox)
            y1 = min(p[1] for p in region.bbox)
            x2 = max(p[0] for p in region.bbox)
            y2 = max(p[1] for p in region.bbox)
            time_w = int((x2 - x1) * 0.45)
            time_h = y2 - y1
            time_x1, time_y1 = x1, y1
        else:
            time_x1 = min(p[0] for p in time_bbox)
            time_y1 = min(p[1] for p in time_bbox)
            time_x2 = max(p[0] for p in time_bbox)
            time_y2 = max(p[1] for p in time_bbox)
            time_w = time_x2 - time_x1
            time_h = time_y2 - time_y1

        try:
            dt = datetime.strptime(target_date, "%Y-%m-%d")
        except ValueError:
            dt = datetime.now()

        t_parts = target_time.split(':')
        hour = t_parts[0] if len(t_parts) > 0 else '08'
        minute = t_parts[1] if len(t_parts) > 1 else '00'

        time_str = f"{hour}:{minute}"
        date_str = f"{dt.day:02d} Tháng {dt.month},{dt.year}"
        weekday_str = self.VN_WEEKDAYS[dt.weekday()]

        font_path = self._ROBOTO_STATIC if os.path.exists(self._ROBOTO_STATIC) else (self._find_font() or "arial.ttf")

        # 1. Render Time scaled to target_w, target_h
        f_time = ImageFont.truetype(font_path, 120)
        t_bbox = f_time.getbbox(time_str)
        raw_tw = t_bbox[2] - t_bbox[0]
        raw_th = t_bbox[3] - t_bbox[1]

        txt_canvas = Image.new('RGBA', (raw_tw + 20, raw_th + 20), (0, 0, 0, 0))
        d_cv = ImageDraw.Draw(txt_canvas)
        d_cv.text((10 - t_bbox[0], 10 - t_bbox[1]), time_str, font=f_time, fill=(255, 255, 255), stroke_width=2, stroke_fill=(255, 255, 255))

        s_canvas = Image.new('RGBA', (raw_tw + 20, raw_th + 20), (0, 0, 0, 0))
        sd_cv = ImageDraw.Draw(s_canvas)
        sd_cv.text((10 - t_bbox[0], 10 - t_bbox[1]), time_str, font=f_time, fill=(0, 0, 0, 48), stroke_width=2, stroke_fill=(0, 0, 0, 48))

        target_tw = max(180, time_w)
        target_th = max(80, time_h)
        scaled_time = txt_canvas.resize((target_tw, target_th), Image.Resampling.LANCZOS)
        scaled_shadow = s_canvas.resize((target_tw, target_th), Image.Resampling.LANCZOS).filter(ImageFilter.GaussianBlur(radius=0.8))

        time_shadow_offset = min(3, max(1, int(round(target_th * 0.015))))
        pil_img.paste(
            scaled_shadow,
            (time_x1 + time_shadow_offset, time_y1 + time_shadow_offset),
            mask=scaled_shadow,
        )
        pil_img.paste(scaled_time, (time_x1, time_y1), mask=scaled_time)

        # 2. Vạch ngăn vàng / cam
        sep_x = time_x1 + target_tw + 16
        sep_y1 = time_y1 + int(target_th * 0.03)
        sep_y2 = time_y1 + int(target_th * 0.97)
        sep_color = info.get('separator_color', (212, 168, 75))

        draw = ImageDraw.Draw(pil_img)
        draw.line([(sep_x + 1, sep_y1 + 1), (sep_x + 1, sep_y2 + 1)], fill=(0, 0, 0, 70), width=4)
        draw.line([(sep_x, sep_y1), (sep_x, sep_y2)], fill=sep_color, width=4)

        # 3. Ngày và Thứ ở cột phải
        right_x = sep_x + 18
        date_font_size = max(24, int(target_th * 0.38))
        weekday_font_size = max(22, int(target_th * 0.36))

        f_date = ImageFont.truetype(font_path, date_font_size)
        f_weekday = ImageFont.truetype(font_path, weekday_font_size)

        date_y = time_y1 + int(target_th * 0.02)
        weekday_y = time_y1 + int(target_th * 0.52)

        right_shadow_offset = min(2, max(1, int(round(date_font_size * 0.035))))
        draw.text(
            (right_x + right_shadow_offset, date_y + right_shadow_offset),
            date_str,
            font=f_date,
            fill=(0, 0, 0, 45),
        )
        draw.text(
            (right_x + right_shadow_offset, weekday_y + right_shadow_offset),
            weekday_str,
            font=f_weekday,
            fill=(0, 0, 0, 45),
        )

        draw.text((right_x, date_y), date_str, font=f_date, fill=(255, 255, 255))
        draw.text((right_x, weekday_y), weekday_str, font=f_weekday, fill=(255, 255, 255))

        return pil_img

    # ==================== 3b. ADAPTIVE COLOR SAMPLING ====================

    def _sample_reference_text_color(
        self, img: np.ndarray, timestamp_bbox: List[List[int]], h: int, w: int
    ) -> Optional[Tuple[int, int, int]]:
        """
        Lấy mẫu màu pixel thực tế từ các dòng text tham chiếu (địa chỉ, quận, TP...)
        nằm NGAY BÊN DƯỚI dòng timestamp trên ảnh gốc/inpainted.
        Đảm bảo text mới render đồng dạng 100% về màu sắc với text gốc camera.

        Thuật toán:
        1. Xác định vùng tham chiếu: từ đáy timestamp_bbox đến cuối ảnh (chiều cao)
        2. Tìm pixel sáng (text trắng/xám camera) trong vùng tham chiếu
        3. Trả về RGB trung bình achromatic (R ≈ G ≈ B)
        """
        y_bottom = max(p[1] for p in timestamp_bbox)
        x_left = min(p[0] for p in timestamp_bbox)
        x_right = max(p[0] for p in timestamp_bbox)

        # Vùng tham chiếu: từ đáy timestamp đến cuối ảnh, mở rộng sang phải
        ref_y1 = min(h - 1, y_bottom + 2)
        ref_y2 = min(h, int(h * 0.995))
        ref_x1 = max(0, int(w * 0.25))
        ref_x2 = min(w, int(w * 0.99))

        if ref_y2 <= ref_y1 or ref_x2 <= ref_x1:
            return None

        roi = img[ref_y1:ref_y2, ref_x1:ref_x2]
        if roi.size == 0:
            return None

        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

        # Tìm pixel sáng (text) trong vùng tham chiếu
        # Camera timestamp text thường có gray >= 170
        bright_mask = gray >= 170
        bright_pixels = roi[bright_mask]

        if len(bright_pixels) < 20:
            # Thử ngưỡng thấp hơn
            bright_mask = gray >= 150
            bright_pixels = roi[bright_mask]

        if len(bright_pixels) < 20:
            return None

        # Tính trung bình BGR → RGB
        mean_bgr = np.mean(bright_pixels, axis=0).astype(int)
        r, g, b = int(mean_bgr[2]), int(mean_bgr[1]), int(mean_bgr[0])

        # Kiểm tra: text camera thật luôn achromatic (R ≈ G ≈ B, chênh < 15)
        # Nếu chênh lệch quá lớn → có thể nhầm lẫn với nền sáng → bỏ qua
        max_diff = max(abs(r - g), abs(g - b), abs(r - b))
        if max_diff > 25:
            logger.info(f"Reference color ({r},{g},{b}) too chromatic (diff={max_diff}), skipping adaptive matching")
            return None

        logger.info(f"Sampled reference text color: RGB=({r},{g},{b}) from {len(bright_pixels)} pixels")
        return (r, g, b)

    # ==================== 4. FORMAT DATE/TIME ====================

    def _format_new_datetime(
        self, original_text: str, date_format: str,
        target_date: str, target_time: str
    ) -> str:
        """Tạo text mới giữ nguyên format của text cũ nhưng thay số ngày/giờ."""
        try:
            dt = datetime.strptime(target_date, "%Y-%m-%d")
        except ValueError:
            dt = datetime.now()

        time_parts = target_time.split(':')
        hour = time_parts[0] if len(time_parts) > 0 else '08'
        minute = time_parts[1] if len(time_parts) > 1 else '00'
        second = time_parts[2] if len(time_parts) > 2 else '00'

        new_text = original_text
        date_replaced = False

        ENG_MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        ENG_MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        target_eng_short = ENG_MONTHS_SHORT[dt.month - 1]
        target_eng_full = ENG_MONTHS_FULL[dt.month - 1]

        # 1. Thử pattern Tiếng Anh: "18 Jan 2026", "18 January 2026", "Jan 18, 2026"
        for pattern, fmt in self.DATE_PATTERNS:
            if fmt == 'DD_ENG_YYYY':
                match = pattern.search(new_text)
                if match:
                    is_full = len(match.group(2)) > 3
                    m_str = target_eng_full if is_full else target_eng_short
                    new_date = f"{dt.day:02d} {m_str} {dt.year}"
                    new_text = new_text[:match.start()] + new_date + new_text[match.end():]
                    date_replaced = True
                    break
            elif fmt == 'ENG_DD_YYYY':
                match = pattern.search(new_text)
                if match:
                    is_full = len(match.group(1)) > 3
                    m_str = target_eng_full if is_full else target_eng_short
                    new_date = f"{m_str} {dt.day:02d}, {dt.year}"
                    new_text = new_text[:match.start()] + new_date + new_text[match.end():]
                    date_replaced = True
                    break

        # 2. Thử pattern Tiếng Việt (phổ biến nhất trên camera VN: 11 Thl, 2026 / 14 Th9, 2026 / 15 Thg 9, 2026)
        if not date_replaced:
            for pattern, fmt in self.DATE_PATTERNS:
                if fmt == 'DD_THG_MM_YYYY':
                    match = pattern.search(new_text)
                    if match:
                        matched_str = match.group(0)
                        has_comma = ',' in matched_str
                        comma_part = ',' if has_comma else ''
                        th_prefix = "Th"
                        if "tháng" in matched_str.lower():
                            th_prefix = "Tháng "
                        elif "thg" in matched_str.lower():
                            th_prefix = "Thg "
                        new_date = f"{dt.day:02d} {th_prefix}{dt.month}{comma_part} {dt.year}"
                        new_text = new_text[:match.start()] + new_date + new_text[match.end():]
                        date_replaced = True
                        break

        # 3. Thử DD/MM/YYYY
        if not date_replaced:
            for pattern, fmt in self.DATE_PATTERNS:
                if fmt == 'DD/MM/YYYY':
                    match = pattern.search(new_text)
                    if match:
                        sep = new_text[match.start(1) + len(match.group(1))]
                        new_date = f"{dt.day:02d}{sep}{dt.month:02d}{sep}{dt.year}"
                        new_text = new_text[:match.start()] + new_date + new_text[match.end():]
                        date_replaced = True
                        break

        # 4. Thử YYYY/MM/DD
        if not date_replaced:
            for pattern, fmt in self.DATE_PATTERNS:
                if fmt == 'YYYY/MM/DD':
                    match = pattern.search(new_text)
                    if match:
                        sep = new_text[match.start(1) + len(match.group(1))]
                        new_date = f"{dt.year}{sep}{dt.month:02d}{sep}{dt.day:02d}"
                        new_text = new_text[:match.start()] + new_date + new_text[match.end():]
                        date_replaced = True
                        break

        # 5. Thay thế giờ
        time_match = self.TIME_PATTERN.search(new_text)
        if time_match:
            full_match = time_match.group(0)
            sep = ':'
            for s in [':', '.', '-']:
                if s in full_match:
                    sep = s
                    break
            parts = re.split(r'[:.\-h]', full_match)
            ms_suffix = ""
            if ',' in full_match:
                ms_suffix = full_match[full_match.index(','):]

            if len(parts) >= 3 and parts[2].strip():
                new_time = f"{hour}{sep}{minute}{sep}{second}{ms_suffix}"
            else:
                new_time = f"{hour}{sep}{minute}{ms_suffix}"
            new_text = new_text[:time_match.start()] + new_time + new_text[time_match.end():]

        # 6. Thay thế thứ trong tuần (weekday tiếng Việt)
        if self.WEEKDAY_PATTERN.search(new_text):
            weekday_vn = self.VN_WEEKDAYS[dt.weekday()]
            new_text = self.WEEKDAY_PATTERN.sub(weekday_vn, new_text)

        return new_text

    # ==================== 5. MAIN REPLACE FUNCTION ====================

    def replace_timestamp(
        self,
        image_path: str,
        output_path: str,
        target_date: str,
        target_time: str
    ) -> bool:
        """
        Quy trình chính:
        1. Kiểm tra OCR trước: Nếu là watermark Timemark (Giờ lớn | Ngày / Thứ),
           xử lý chuyên biệt bằng Timemark Engine để giữ nguyên 100% bố cục 2 cột và ngày/thứ chuẩn.
        2. Thử AI Image Editing trước (Gemini trực tiếp xóa text cũ + render text mới trên ảnh) cho các watermark khác.
        3. Fallback sang pipeline PIL chuẩn: detect → inpaint siêu mỏng nét chữ → render lại chữ mới.
        """
        # 0. Đọc ảnh gốc an toàn qua UTF-8
        img = imread_utf8(image_path)
        if img is None:
            logger.error(f"Không đọc được ảnh: {image_path}")
            return False

        h, w = img.shape[:2]

        # ======================== KIỂM TRA TIMEMARK 2 CỘT ========================
        ocr_regions = self.detect_timestamp_regions(image_path)
        timemark_regions = [r for r in ocr_regions if r.is_timemark]
        if timemark_regions:
            logger.info("Phát hiện watermark dạng Timemark 2 cột (Giờ | Ngày / Thứ). Sử dụng Timemark Engine trực tiếp...")
            tm_region = timemark_regions[0]
            mask = self._build_timemark_mask(img, tm_region)
            inpainted = cv2.inpaint(
                img,
                mask,
                inpaintRadius=self._timemark_inpaint_radius(mask),
                flags=cv2.INPAINT_NS,
            )
            result_pil = Image.fromarray(cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB))
            result_pil = self._render_timemark_block(result_pil, tm_region, target_date, target_time)

            result_bgr = cv2.cvtColor(np.array(result_pil), cv2.COLOR_RGB2BGR)
            saved = imwrite_utf8(output_path, result_bgr, quality=95)
            if not saved:
                result_pil.save(output_path, format='JPEG', quality=95)

            if self.ai_service and self.ai_service.get_full_config().get('anti_ai_enabled', True):
                try:
                    from src.anti_ai_detector import clean_ai_traces
                    clean_ai_traces(output_path, image_path)
                except Exception as ae:
                    logger.warning(f"Lỗi Anti-AI clean traces: {ae}")

            logger.info(f"Thay thế watermark Timemark thành công: {output_path}")
            return True

        # ======================== BƯỚC 1: AI IMAGE EDITING ========================
        # Ưu tiên dùng Gemini AI để chỉnh sửa ảnh trực tiếp (text tự nhiên nhất)
        ai_info = None
        if self.ai_service and self.ai_service.is_configured():
            try:
                raw_bytes = Path(image_path).read_bytes()

                # Gọi detect_and_analyze_watermark trước để lấy original_text + new_text chính xác
                try:
                    ai_info = self.ai_service.detect_and_analyze_watermark(raw_bytes, target_date, target_time)
                except Exception as e:
                    logger.warning(f"Lỗi AI detect: {e}")

                original_text = ""
                new_text = ""
                if ai_info:
                    original_text = ai_info.get('original_text', '')
                    new_text = ai_info.get('new_text', '')

                # Thử AI Image Editing
                logger.info("Đang thử AI Image Editing (Gemini sửa ảnh trực tiếp)...")
                edited_bytes = self.ai_service.edit_timestamp_image(
                    raw_bytes=raw_bytes,
                    target_date=target_date,
                    target_time=target_time,
                    original_text=original_text,
                    new_text=new_text,
                    timestamp_info=ai_info,
                )

                if edited_bytes:
                    # Lưu ảnh AI đã chỉnh sửa
                    try:
                        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                        with open(output_path, 'wb') as f:
                            f.write(edited_bytes)
                        logger.info(f"AI Image Editing thành công! Đã lưu: {output_path}")

                        # Áp dụng Anti-AI Watermark & Metadata Sanitizer
                        should_clean = True
                        if self.ai_service:
                            should_clean = self.ai_service.get_full_config().get('anti_ai_enabled', True)
                        if should_clean:
                            try:
                                from src.anti_ai_detector import clean_ai_traces
                                clean_ai_traces(output_path, image_path)
                            except Exception as ae:
                                logger.warning(f"Lỗi Anti-AI clean traces: {ae}")

                        return True
                    except Exception as save_err:
                        logger.warning(f"Lỗi lưu ảnh AI edited: {save_err}")

                logger.info("AI Image Editing không thành công, chuyển sang fallback PIL rendering...")
            except Exception as e:
                logger.warning(f"Lỗi AI Image Editing pipeline: {e}")

        # ======================== BƯỚC 2: FALLBACK PIL RENDERING ========================
        # Pipeline cũ: detect → inpaint → PIL render
        # Tái sử dụng ai_info nếu đã có từ Bước 1 để tránh gọi API trùng lặp
        if ai_info is None and self.ai_service and self.ai_service.is_configured():
            try:
                raw_bytes = Path(image_path).read_bytes()
                ai_info = self.ai_service.detect_and_analyze_watermark(raw_bytes, target_date, target_time)
            except Exception as e:
                logger.warning(f"Lỗi AI detect_and_analyze_watermark: {e}")

        regions = []
        ai_style = None

        if ai_info and ai_info.get('box_2d'):
            ymin, xmin, ymax, xmax = ai_info['box_2d']
            pad = 4
            x1 = max(0, int(xmin * w / 1000.0) - pad)
            y1 = max(0, int(ymin * h / 1000.0) - pad)
            x2 = min(w, int(xmax * w / 1000.0) + pad)
            y2 = min(h, int(ymax * h / 1000.0) + pad)
            bbox = [[x1, y1], [x2, y1], [x2, y2], [x1, y2]]

            # Kiểm tra Safe Zone cho box AI trả về
            if not self.is_in_safe_zone(bbox, h):
                logger.warning(f"AI trả về box [{x1},{y1},{x2},{y2}] ngoài Safe Zone (giữa ảnh). Hủy bỏ AI box này!")
                ai_info = None
            else:
                orig_text = ai_info.get('original_text', '')
                new_text = ai_info.get('new_text', '')
                if not new_text:
                    new_text = self._format_new_datetime(orig_text, "", target_date, target_time)

                color = hex_to_rgb(ai_info.get('text_color'), (250, 248, 246))
                font_size = max(16, int((y2 - y1) * 0.70))

                regions = [TimestampRegion(
                    bbox=bbox,
                    text=orig_text,
                    confidence=1.0,
                    detected_color=color,
                    estimated_font_size=font_size
                )]
                ai_style = ai_info
                logger.info(f"Đã phát hiện watermark qua AI: box=[{x1},{y1},{x2},{y2}], new_text='{new_text}'")

        if not regions:
            # Fallback sang OCR nếu có
            regions = ocr_regions if ocr_regions else self.detect_timestamp_regions(image_path)
            if not regions:
                # Heuristic: Watermark camera thường ở góc dưới phải (y: 93%-99%, x: 52%-98%)
                rx1, ry1 = int(w * 0.52), int(h * 0.93)
                rx2, ry2 = int(w * 0.99), int(h * 0.99)
                r_bbox = [[rx1, ry1], [rx2, ry1], [rx2, ry2], [rx1, ry2]]
                color, font_size = self._analyze_text_style(img, r_bbox)
                regions = [TimestampRegion(
                    bbox=r_bbox,
                    text="",
                    confidence=0.5,
                    detected_color=color,
                    estimated_font_size=max(18, int((ry2 - ry1) * 0.75))
                )]

        # 3. Tạo mask siêu mỏng cho vùng watermark
        combined_mask = np.zeros((h, w), dtype=np.uint8)
        replacement_items = []

        for region in regions:
            logger.info(
                "Watermark mask pipeline=dark-text-v3 module=%s color=%s",
                __file__,
                region.detected_color,
            )
            region_mask = self._build_text_mask(
                img,
                region.bbox,
                padding=1,
                detected_color=region.detected_color,
            )
            # Không fill toàn bộ polygon để tránh bôi đen loang lổ nền
            combined_mask = cv2.bitwise_or(combined_mask, region_mask)

            item_new_text = (ai_style.get('new_text') if ai_style and ai_style.get('new_text')
                             else self._format_new_datetime(region.text, region.date_format, target_date, target_time))

            replacement_items.append({
                'region': region,
                'new_text': item_new_text
            })

        # 4. Inpaint chuẩn xác với Navier-Stokes để xóa sạch 100% chữ cũ và bóng đổ mà không làm loang nền hay tạo vệt
        inpainted = cv2.inpaint(img, combined_mask, inpaintRadius=3, flags=cv2.INPAINT_NS)

        # 5. Vẽ text mới lên ảnh đã inpaint với Drop Shadow tự nhiên đồng dạng với các dòng dưới
        result_pil = Image.fromarray(cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB))

        for item in replacement_items:
            region = item['region']
            new_text = item['new_text']

            bbox = region.bbox
            x1 = min(p[0] for p in bbox)
            y1 = min(p[1] for p in bbox)
            x2 = max(p[0] for p in bbox)
            y2 = max(p[1] for p in bbox)

            font_size = region.estimated_font_size
            font = None
            for fpath in self.FONT_PATHS:
                if os.path.exists(fpath):
                    try:
                        font = ImageFont.truetype(fpath, font_size)
                        break
                    except Exception:
                        pass
            if font is None:
                font = ImageFont.load_default()

            # Tính toán kích thước text bằng Pillow
            temp_draw = ImageDraw.Draw(result_pil)
            text_bbox = temp_draw.textbbox((0, 0), new_text, font=font)
            text_w = text_bbox[2] - text_bbox[0]
            text_h = text_bbox[3] - text_bbox[1]

            # Căn lề phải tự nhiên đồng dạng với các dòng địa chỉ phía dưới
            right_margin_x = ai_style.get('right_margin_x') if ai_style else None
            if right_margin_x is not None:
                right_x = int(right_margin_x * w / 1000.0)
            elif region.ref_right_x is not None:
                right_x = min(w - 6, region.ref_right_x)
            else:
                right_x = x2

            alignment = ai_style.get('alignment', 'right') if ai_style else ('right' if (x1 + x2)/2 > w * 0.5 else 'left')
            if alignment == 'right':
                text_x = right_x - text_w
            else:
                text_x = x1

            region_h = y2 - y1
            # Căn chân chữ (baseline) khớp tuyệt đối với dòng ngày giờ cũ và các dòng địa chỉ
            text_y = (y2 - int(region_h * 0.15)) - text_bbox[3]

            text_x = max(0, min(text_x, w - text_w - 1))
            text_y = max(0, min(text_y, h - text_h - 1))

            # Xác định màu sắc, bóng đổ từ AI hoặc fallback chuẩn
            # Phân tích pixel cho thấy text gốc camera là achromatic white (R ≈ G ≈ B ≈ 194)
            # KHÔNG dùng warm tone (R>G>B) vì sẽ bị phát hiện khác biệt
            if ai_style:
                text_color = hex_to_rgb(ai_style.get("text_color"), region.detected_color)
                has_shadow = ai_style.get("has_shadow", True)
            else:
                text_color = region.detected_color if region.detected_color != (0, 0, 0) else (255, 255, 255)
                has_shadow = True

            # Adaptive color matching: Lấy mẫu màu thực tế từ các dòng text tham chiếu
            # trên ảnh đã inpaint (Đường số 14, Quận 7, TP HCM...) để đồng bộ chính xác
            ref_color = self._sample_reference_text_color(inpainted, region.bbox, h, w)
            if ref_color is not None:
                text_color = ref_color
                logger.info(f"Adaptive color: sampled reference text color RGB={ref_color}")

            # Drop shadow nhẹ nhàng đồng dạng camera gốc
            # Phân tích pixel: shadow gốc camera có alpha thấp (40-65), không quá đậm
            if has_shadow:
                shadow_ov = Image.new('RGBA', result_pil.size, (0, 0, 0, 0))
                sdraw = ImageDraw.Draw(shadow_ov)
                # 1. Ambient halo mỏng (viền tối 1px xung quanh, alpha nhẹ)
                for dx, dy in [(-1, 0), (1, 0), (0, -1), (0, 1), (-1, -1), (1, -1), (-1, 1), (1, 1)]:
                    sdraw.text((text_x + dx, text_y + dy), new_text, font=font, fill=(0, 0, 0, 35))
                # 2. Directional drop shadow (+1px dưới phải, nhẹ hơn)
                sdraw.text((text_x + 1, text_y + 1), new_text, font=font, fill=(0, 0, 0, 70))
                sdraw.text((text_x + 1, text_y + 2), new_text, font=font, fill=(0, 0, 0, 40))
                shadow_ov = shadow_ov.filter(ImageFilter.GaussianBlur(radius=0.8))
                result_pil = Image.alpha_composite(result_pil.convert('RGBA'), shadow_ov).convert('RGB')

            # Chữ chính: full opacity (alpha=255) với optical blur cực nhẹ (0.3px)
            # Phân tích pixel cho thấy text gốc camera gần như sắc nét hoàn toàn
            # chỉ có anti-aliasing tự nhiên, không cần blur mạnh
            txt_ov = Image.new('RGBA', result_pil.size, (0, 0, 0, 0))
            tdraw = ImageDraw.Draw(txt_ov)
            tdraw.text((text_x, text_y), new_text, font=font, fill=(*text_color, 255))
            txt_ov = txt_ov.filter(ImageFilter.GaussianBlur(radius=0.3))
            result_pil = Image.alpha_composite(result_pil.convert('RGBA'), txt_ov).convert('RGB')

            logger.info(
                f"Đã thay '{region.text}' → '{new_text}' "
                f"tại ({text_x},{text_y}), color={text_color}, font={font.getname() if hasattr(font, 'getname') else 'default'}"
            )

        # 6. Lưu kết quả an toàn qua UTF-8
        result_bgr = cv2.cvtColor(np.array(result_pil), cv2.COLOR_RGB2BGR)
        saved = imwrite_utf8(output_path, result_bgr, quality=95)
        if not saved:
            result_pil.save(output_path, format='JPEG', quality=95)

        # Áp dụng Anti-AI Watermark & Metadata Sanitizer
        should_clean = True
        if self.ai_service:
            should_clean = self.ai_service.get_full_config().get('anti_ai_enabled', True)
        if should_clean:
            try:
                from src.anti_ai_detector import clean_ai_traces
                clean_ai_traces(output_path, image_path)
            except Exception as ae:
                logger.warning(f"Lỗi Anti-AI clean traces: {ae}")

        logger.info(f"Đã thay thế {len(replacement_items)} vùng timestamp thành công")
        return True
