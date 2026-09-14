# -*- coding: utf-8 -*-
"""Các thành phần phân tích thuần cục bộ cho chỉnh ngày/giờ ảnh."""

import re
import random
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


class EditRequestNormalizer:
    """Chuẩn hóa ngày/giờ theo từng ảnh; dữ liệu có cấu trúc luôn là nguồn chuẩn."""

    _WEEKDAYS = (
        "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm",
        "Thứ Sáu", "Thứ Bảy", "Chủ Nhật",
    )

    def normalize_item(self, target_date: str, target_time: str) -> Dict[str, str]:
        parsed_date = datetime.strptime(target_date, "%Y-%m-%d")
        parsed_time = None
        for time_format in ("%H:%M:%S", "%H:%M"):
            try:
                parsed_time = datetime.strptime(target_time, time_format)
                break
            except ValueError:
                continue
        if parsed_time is None:
            raise ValueError("Giờ không hợp lệ; cần HH:MM hoặc HH:MM:SS")
        return {
            "target_date": parsed_date.strftime("%Y-%m-%d"),
            "target_time": parsed_time.strftime("%H:%M:%S"),
            "weekday": self._WEEKDAYS[parsed_date.weekday()],
        }


class TimestampAnalyzer:
    """Nhận diện và xếp hạng dòng có ngày/giờ từ kết quả OCR."""

    DATE_PATTERNS = (
        re.compile(r"\b\d{1,2}\s*(?:thg?|tháng)\s*\d{1,2}\s*,?\s*\d{4}\b", re.I),
        re.compile(r"\b\d{1,2}[\-/\.]\d{1,2}[\-/\.]\d{4}\b"),
        re.compile(r"\b\d{4}[\-/\.]\d{1,2}[\-/\.]\d{1,2}\b"),
    )
    TIME_PATTERN = re.compile(r"\b(?:[01]?\d|2[0-3])[:.]\d{2}(?:[:.]\d{2})?(?:\.\d{1,3})?\b")

    def __init__(self, ocr_engine: Any = None):
        self._ocr_engine = ocr_engine

    @classmethod
    def _has_date(cls, text: str) -> bool:
        return any(pattern.search(text) for pattern in cls.DATE_PATTERNS)

    @staticmethod
    def _box_to_bounds(box: Any) -> List[int]:
        if isinstance(box, (list, tuple)) and len(box) == 4 and all(
            isinstance(value, (int, float)) for value in box
        ):
            return [int(round(value)) for value in box]
        points = list(box or [])
        if len(points) >= 4:
            xs = [float(point[0]) for point in points]
            ys = [float(point[1]) for point in points]
            return [int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))]
        raise ValueError("OCR box không hợp lệ")

    def select_candidate(
        self, lines: Iterable[Dict[str, Any]], image_size: Tuple[int, int]
    ) -> Optional[Dict[str, Any]]:
        width, height = image_size
        candidates: List[Tuple[float, Dict[str, Any]]] = []
        for line in lines:
            text = str(line.get("text", "")).strip()
            has_date = self._has_date(text)
            has_time = bool(self.TIME_PATTERN.search(text))
            if not has_date:
                continue
            bounds = self._box_to_bounds(line.get("box"))
            confidence = float(line.get("score", 0.0) or 0.0)
            vertical_position = max(0.0, min(1.0, bounds[1] / max(1, height)))
            semantic_score = 1.0 + (0.5 if has_time else 0.0)
            candidates.append((semantic_score + confidence + vertical_position * 0.1, {
                "text": text,
                "box": bounds,
                "confidence": confidence,
                "status": "quick_ready" if has_time and confidence >= 0.80 else "needs_region_review",
            }))
        if not candidates:
            return None
        return max(candidates, key=lambda entry: entry[0])[1]


class ImageSourceService:
    """Đọc và lấy mẫu ảnh chân dung trong đúng thư mục dự án/nhân viên."""

    IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

    def __init__(self, portrait_root: Any, rng: Optional[random.Random] = None):
        self.portrait_root = Path(portrait_root).resolve()
        self.rng = rng or random.SystemRandom()

    def _employee_dir(self, project: str, employee: str) -> Path:
        if not project.strip() or not employee.strip():
            raise ValueError("Thiếu dự án hoặc nhân viên")
        candidate = (self.portrait_root / project / employee).resolve()
        try:
            candidate.relative_to(self.portrait_root)
        except ValueError as exc:
            raise ValueError("Đường dẫn dự án/nhân viên không hợp lệ") from exc
        return candidate

    def sample_employee_images(self, project: str, employee: str, count: int) -> List[str]:
        employee_dir = self._employee_dir(project, employee)
        if not employee_dir.is_dir():
            return []
        images = sorted(
            str(path) for path in employee_dir.iterdir()
            if path.is_file() and path.suffix.lower() in self.IMAGE_EXTENSIONS
        )
        limit = min(max(1, int(count)), 31, len(images))
        return self.rng.sample(images, limit) if images else []
