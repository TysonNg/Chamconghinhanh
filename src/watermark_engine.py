# -*- coding: utf-8 -*-
"""
Watermark Engine - Xử lý watermark ảnh chấm công
Bao gồm: Scanner, Remover, Generator, EXIF Editor
"""

import os
import re
import json
import struct
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ExifTags

logger = logging.getLogger(__name__)

# ==================== EXIF EDITOR ====================

class ExifEditor:
    """Chỉnh sửa EXIF metadata của ảnh"""
    
    @staticmethod
    def read_exif(image_path: str) -> Dict:
        """Đọc EXIF metadata từ ảnh"""
        try:
            img = Image.open(image_path)
            exif_data = img._getexif()
            if not exif_data:
                return {}
            
            result = {}
            for tag_id, value in exif_data.items():
                tag_name = ExifTags.TAGS.get(tag_id, tag_id)
                result[tag_name] = value
            return result
        except Exception as e:
            logger.warning(f"Không đọc được EXIF: {e}")
            return {}
    
    @staticmethod
    def modify_exif_date(image_path: str, output_path: str, 
                         target_date: str, target_time: str) -> bool:
        """
        Sửa ngày giờ trong EXIF metadata.
        target_date: 'YYYY-MM-DD'  
        target_time: 'HH:MM:SS'
        """
        try:
            # Đọc file ảnh gốc dưới dạng bytes
            with open(image_path, 'rb') as f:
                data = f.read()
            
            # Format EXIF datetime: "YYYY:MM:DD HH:MM:SS"
            date_parts = target_date.split('-')
            exif_datetime = f"{date_parts[0]}:{date_parts[1]}:{date_parts[2]} {target_time}"
            exif_datetime_bytes = exif_datetime.encode('ascii')
            
            # Tìm và thay thế tất cả datetime tags trong EXIF
            # EXIF datetime format luôn là 19 bytes + null terminator = 20 bytes
            # Pattern: YYYY:MM:DD HH:MM:SS (19 chars)
            exif_date_pattern = re.compile(
                rb'(\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2})'
            )
            
            modified_data = exif_date_pattern.sub(exif_datetime_bytes, data)
            
            with open(output_path, 'wb') as f:
                f.write(modified_data)
            
            logger.info(f"Đã sửa EXIF date thành {exif_datetime}")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi sửa EXIF: {e}")
            # Nếu không sửa được EXIF, copy file gốc
            if image_path != output_path:
                import shutil
                shutil.copy2(image_path, output_path)
            return False
    
    @staticmethod
    def strip_exif(image_path: str, output_path: str) -> bool:
        """Xóa toàn bộ EXIF data (dùng khi muốn tạo ảnh sạch)"""
        try:
            img = Image.open(image_path)
            # Tạo ảnh mới không có EXIF
            data = list(img.getdata())
            clean_img = Image.new(img.mode, img.size)
            clean_img.putdata(data)
            clean_img.save(output_path, quality=95)
            return True
        except Exception as e:
            logger.error(f"Lỗi xóa EXIF: {e}")
            return False


# ==================== WATERMARK SCANNER ====================

class WatermarkScanner:
    """Phân tích format watermark từ ảnh mẫu"""
    
    def __init__(self):
        self.templates = []
    
    def scan_image(self, image_path: str) -> Optional[Dict]:
        """
        Phân tích ảnh để tìm watermark.
        Trả về thông tin: vị trí, kích thước, màu sắc, nội dung text
        """
        img = cv2.imread(image_path)
        if img is None:
            logger.error(f"Không đọc được ảnh: {image_path}")
            return None
        
        h, w = img.shape[:2]
        
        # Watermark thường nằm ở 4 vùng: top-left, top-right, bottom-left, bottom-right
        # Và thường chiếm ~10-25% chiều cao ảnh
        regions = {
            'bottom-left': (0, int(h * 0.7), int(w * 0.7), h),
            'bottom-right': (int(w * 0.3), int(h * 0.7), w, h),
            'top-left': (0, 0, int(w * 0.7), int(h * 0.3)),
            'top-right': (int(w * 0.3), 0, w, int(h * 0.3)),
            'bottom-full': (0, int(h * 0.75), w, h),
            'top-full': (0, 0, w, int(h * 0.25)),
        }
        
        results = []
        for region_name, (x1, y1, x2, y2) in regions.items():
            roi = img[y1:y2, x1:x2]
            analysis = self._analyze_region(roi, region_name)
            if analysis and analysis.get('has_text'):
                analysis['region'] = region_name
                analysis['bounds'] = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2}
                analysis['image_size'] = {'width': w, 'height': h}
                results.append(analysis)
        
        if not results:
            return None
        
        # Chọn vùng có nhiều text nhất (khả năng cao là watermark)
        best = max(results, key=lambda r: r.get('text_density', 0))
        return best
    
    def _analyze_region(self, roi: np.ndarray, region_name: str) -> Dict:
        """Phân tích một vùng ảnh để tìm watermark text"""
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape
        
        # Tìm text bằng cách phát hiện vùng có contrast cao
        # Text watermark thường là màu trắng/vàng trên nền tối hoặc ngược lại
        
        # Kiểm tra vùng sáng (text trắng/vàng)
        _, bright_mask = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)
        bright_ratio = np.sum(bright_mask > 0) / (h * w)
        
        # Kiểm tra vùng tối (text đen)  
        _, dark_mask = cv2.threshold(gray, 50, 255, cv2.THRESH_BINARY_INV)
        dark_ratio = np.sum(dark_mask > 0) / (h * w)
        
        # Watermark text thường chiếm 5-30% diện tích vùng
        has_text = (0.01 < bright_ratio < 0.4) or (0.01 < dark_ratio < 0.4)
        
        # Phân tích màu chủ đạo của text
        text_color = self._detect_text_color(roi)
        
        # Phân tích kích thước font ước lượng
        font_size_estimate = self._estimate_font_size(gray)
        
        return {
            'has_text': has_text,
            'text_density': max(bright_ratio, dark_ratio),
            'text_color': text_color,
            'font_size_estimate': font_size_estimate,
            'region_size': {'width': w, 'height': h},
            'bg_color': self._get_dominant_color(roi),
        }
    
    def _detect_text_color(self, roi: np.ndarray) -> Tuple[int, int, int]:
        """Phát hiện màu text chủ đạo trong vùng watermark"""
        hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)
        
        # Text watermark thường là trắng hoặc vàng
        # Kiểm tra trắng
        white_mask = cv2.inRange(hsv, (0, 0, 200), (180, 30, 255))
        white_count = np.sum(white_mask > 0)
        
        # Kiểm tra vàng/cam  
        yellow_mask = cv2.inRange(hsv, (15, 100, 150), (35, 255, 255))
        yellow_count = np.sum(yellow_mask > 0)
        
        # Kiểm tra đen
        black_mask = cv2.inRange(hsv, (0, 0, 0), (180, 255, 50))
        black_count = np.sum(black_mask > 0)
        
        if white_count > yellow_count and white_count > black_count:
            return (255, 255, 255)
        elif yellow_count > black_count:
            return (255, 200, 0)
        else:
            return (0, 0, 0)
    
    def _estimate_font_size(self, gray: np.ndarray) -> int:
        """Ước lượng kích thước font từ ảnh grayscale"""
        edges = cv2.Canny(gray, 50, 150)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            return 20
        
        heights = []
        for cnt in contours:
            _, _, w, h = cv2.boundingRect(cnt)
            if 5 < h < gray.shape[0] * 0.8 and w > 3:
                heights.append(h)
        
        if not heights:
            return 20
        
        # Font size ≈ median height of text-like contours
        return int(np.median(heights))
    
    def _get_dominant_color(self, roi: np.ndarray) -> Tuple[int, int, int]:
        """Tìm màu chủ đạo (background) của vùng"""
        pixels = roi.reshape(-1, 3).astype(np.float32)
        criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 10, 1.0)
        _, _, centers = cv2.kmeans(pixels, 2, None, criteria, 3, cv2.KMEANS_PP_CENTERS)
        
        # Chọn cluster có nhiều pixel nhất (background)
        dominant = centers[0].astype(int)
        return tuple(dominant.tolist())
    
    def create_template_from_sample(self, image_path: str) -> Optional[Dict]:
        """Tạo template watermark từ ảnh mẫu"""
        analysis = self.scan_image(image_path)
        if not analysis:
            return None
        
        # Tạo template
        template = {
            'source_image': os.path.basename(image_path),
            'region': analysis['region'],
            'text_color': analysis['text_color'],
            'font_size_ratio': analysis['font_size_estimate'] / analysis['image_size']['height'],
            'bg_color': analysis.get('bg_color', (0, 0, 0)),
            'created_at': datetime.now().isoformat(),
        }
        
        self.templates.append(template)
        return template


# ==================== WATERMARK REMOVER ====================

class WatermarkRemover:
    """Xóa watermark từ ảnh bằng inpainting"""
    
    @staticmethod
    def remove_watermark_region(image_path: str, output_path: str,
                                region: str = 'bottom-left',
                                bounds: Optional[Dict] = None) -> bool:
        """
        Xóa watermark ở vùng chỉ định bằng inpainting.
        region: 'bottom-left', 'bottom-right', 'top-left', 'top-right', 'bottom-full'
        bounds: {'x1', 'y1', 'x2', 'y2'} - toạ độ cụ thể nếu biết
        """
        img = cv2.imread(image_path)
        if img is None:
            return False
        
        h, w = img.shape[:2]
        
        if bounds:
            x1, y1, x2, y2 = bounds['x1'], bounds['y1'], bounds['x2'], bounds['y2']
        else:
            # Tính toạ độ dựa trên region
            region_map = {
                'bottom-left': (0, int(h * 0.85), int(w * 0.6), h),
                'bottom-right': (int(w * 0.4), int(h * 0.85), w, h),
                'top-left': (0, 0, int(w * 0.6), int(h * 0.15)),
                'top-right': (int(w * 0.4), 0, w, int(h * 0.15)),
                'bottom-full': (0, int(h * 0.85), w, h),
                'top-full': (0, 0, w, int(h * 0.15)),
            }
            if region not in region_map:
                logger.error(f"Region không hợp lệ: {region}")
                return False
            x1, y1, x2, y2 = region_map[region]
        
        # Tạo mask cho vùng cần xóa
        mask = np.zeros((h, w), dtype=np.uint8)
        
        # Phát hiện text trong vùng watermark
        roi = img[y1:y2, x1:x2]
        roi_gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        
        # Phát hiện vùng text bằng nhiều phương pháp
        # 1. Text sáng (trắng/vàng)
        _, bright = cv2.threshold(roi_gray, 180, 255, cv2.THRESH_BINARY)
        # 2. Text tối
        _, dark = cv2.threshold(roi_gray, 60, 255, cv2.THRESH_BINARY_INV)
        
        # Kết hợp
        text_mask = cv2.bitwise_or(bright, dark)
        
        # Mở rộng mask một chút để cover sạch watermark
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        text_mask = cv2.dilate(text_mask, kernel, iterations=2)
        
        # Đặt mask vào đúng vị trí
        mask[y1:y2, x1:x2] = text_mask
        
        # Inpainting
        result = cv2.inpaint(img, mask, inpaintRadius=7, flags=cv2.INPAINT_TELEA)
        
        cv2.imwrite(output_path, result, [cv2.IMWRITE_JPEG_QUALITY, 95])
        logger.info(f"Đã xóa watermark vùng {region}")
        return True
    
    @staticmethod
    def remove_watermark_smart(image_path: str, output_path: str) -> bool:
        """
        Tự động phát hiện và xóa watermark.
        Phân tích cả 4 góc + 2 cạnh để tìm vùng watermark.
        """
        scanner = WatermarkScanner()
        analysis = scanner.scan_image(image_path)
        
        if not analysis:
            logger.warning("Không tìm thấy watermark để xóa")
            # Copy nguyên file
            import shutil
            shutil.copy2(image_path, output_path)
            return True
        
        return WatermarkRemover.remove_watermark_region(
            image_path, output_path,
            region=analysis['region'],
            bounds=analysis.get('bounds')
        )


# ==================== WATERMARK GENERATOR ====================

class WatermarkGenerator:
    """Tạo watermark mới cho ảnh"""
    
    # Font paths phổ biến trên Windows
    FONT_PATHS = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        r"C:\Windows\Fonts\tahoma.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\verdana.ttf",
    ]
    
    def __init__(self):
        self.font_path = self._find_font()
    
    def _find_font(self) -> Optional[str]:
        """Tìm font có sẵn trên hệ thống"""
        for path in self.FONT_PATHS:
            if os.path.exists(path):
                return path
        return None
    
    def generate_timestamp_watermark(self, image_path: str, output_path: str,
                                      target_date: str, target_time: str,
                                      position: str = 'bottom-left',
                                      text_color: Tuple[int, int, int] = (255, 165, 0),
                                      font_size: Optional[int] = None,
                                      extra_lines: Optional[List[str]] = None,
                                      bg_opacity: float = 0.5) -> bool:
        """
        Tạo watermark timestamp trên ảnh.
        
        Args:
            image_path: Đường dẫn ảnh gốc
            output_path: Đường dẫn ảnh output
            target_date: Ngày (YYYY-MM-DD)
            target_time: Giờ (HH:MM:SS)
            position: Vị trí watermark
            text_color: Màu text RGB
            font_size: Kích thước font (None = tự tính)
            extra_lines: Các dòng text bổ sung (GPS, địa điểm, v.v.)
            bg_opacity: Độ mờ background (0-1)
        """
        try:
            img = Image.open(image_path).convert('RGBA')
            w, h = img.size
            
            # Tự tính font size nếu không chỉ định
            if font_size is None:
                font_size = max(16, int(h * 0.035))
            
            # Load font
            try:
                font = ImageFont.truetype(self.font_path, font_size) if self.font_path else ImageFont.load_default()
                font_small = ImageFont.truetype(self.font_path, max(12, int(font_size * 0.75))) if self.font_path else ImageFont.load_default()
            except Exception:
                font = ImageFont.load_default()
                font_small = font
            
            # Format ngày giờ theo kiểu phổ biến của app chấm công
            date_parts = target_date.split('-')
            formatted_date = f"{date_parts[2]}/{date_parts[1]}/{date_parts[0]}"
            
            # Tạo các dòng text
            lines = [
                (f"{formatted_date} {target_time}", font),
            ]
            
            if extra_lines:
                for line in extra_lines:
                    lines.append((line, font_small))
            
            # Tính kích thước tổng của watermark
            overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)
            
            # Tính kích thước text
            padding = int(font_size * 0.5)
            line_spacing = int(font_size * 0.3)
            
            total_text_height = 0
            max_text_width = 0
            line_metrics = []
            
            for text, f in lines:
                bbox = draw.textbbox((0, 0), text, font=f)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                line_metrics.append({'text': text, 'font': f, 'width': tw, 'height': th})
                total_text_height += th
                max_text_width = max(max_text_width, tw)
            
            total_text_height += line_spacing * (len(lines) - 1) if len(lines) > 1 else 0
            
            # Tính vị trí
            box_w = max_text_width + padding * 2
            box_h = total_text_height + padding * 2
            
            position_map = {
                'bottom-left': (padding, h - box_h - padding),
                'bottom-right': (w - box_w - padding, h - box_h - padding),
                'top-left': (padding, padding),
                'top-right': (w - box_w - padding, padding),
                'bottom-center': ((w - box_w) // 2, h - box_h - padding),
            }
            
            box_x, box_y = position_map.get(position, position_map['bottom-left'])
            
            # Vẽ background bán trong suốt
            bg_alpha = int(255 * bg_opacity)
            draw.rectangle(
                [box_x, box_y, box_x + box_w, box_y + box_h],
                fill=(0, 0, 0, bg_alpha)
            )
            
            # Vẽ text
            text_color_rgba = (*text_color, 255)
            current_y = box_y + padding
            
            for metric in line_metrics:
                draw.text(
                    (box_x + padding, current_y),
                    metric['text'],
                    font=metric['font'],
                    fill=text_color_rgba
                )
                current_y += metric['height'] + line_spacing
            
            # Composite
            result = Image.alpha_composite(img, overlay)
            
            # Convert back to RGB for JPEG
            if output_path.lower().endswith(('.jpg', '.jpeg')):
                result = result.convert('RGB')
            
            result.save(output_path, quality=95)
            logger.info(f"Đã tạo watermark: {target_date} {target_time} tại {position}")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi tạo watermark: {e}")
            return False
    
    def generate_gps_camera_style(self, image_path: str, output_path: str,
                                    target_date: str, target_time: str,
                                    location_name: str = "",
                                    gps_coords: str = "",
                                    app_style: str = 'timestamp_camera') -> bool:
        """
        Tạo watermark theo style của các app phổ biến.
        
        app_style:
            - 'timestamp_camera': Timestamp Camera style
            - 'gps_map_camera': GPS Map Camera style
            - 'simple': Chỉ ngày giờ đơn giản
        """
        extra_lines = []
        
        if app_style == 'timestamp_camera':
            if location_name:
                extra_lines.append(f"📍 {location_name}")
            if gps_coords:
                extra_lines.append(f"🌐 {gps_coords}")
            
            return self.generate_timestamp_watermark(
                image_path, output_path,
                target_date, target_time,
                position='bottom-left',
                text_color=(255, 165, 0),  # Cam
                extra_lines=extra_lines,
                bg_opacity=0.6
            )
        
        elif app_style == 'gps_map_camera':
            if location_name:
                extra_lines.append(location_name)
            if gps_coords:
                extra_lines.append(gps_coords)
            
            return self.generate_timestamp_watermark(
                image_path, output_path,
                target_date, target_time,
                position='bottom-left',
                text_color=(255, 255, 255),  # Trắng
                extra_lines=extra_lines,
                bg_opacity=0.4
            )
        
        else:  # simple
            return self.generate_timestamp_watermark(
                image_path, output_path,
                target_date, target_time,
                position='bottom-left',
                text_color=(255, 255, 255),
                bg_opacity=0.5
            )
