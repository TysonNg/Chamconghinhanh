"""Full calendar dates and explicit handling of legacy day-only folders."""
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import json
import re

@dataclass(frozen=True)
class DayResolution:
    status: str
    path: Path | None
    reason: str = ""

def parse_attendance_date(value: str) -> date:
    if not isinstance(value, str):
        raise ValueError("Ngày phải có định dạng YYYY-MM-DD hoặc DD/MM/YYYY")
    value = value.strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return date.fromisoformat(value)
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        return datetime.strptime(value, "%d/%m/%Y").date()
    raise ValueError("Ngày phải có đầy đủ ngày, tháng, năm (YYYY-MM-DD)")

def canonical_day_path(project_dir: Path, attendance_date: date) -> Path:
    return Path(project_dir) / attendance_date.isoformat()

def resolve_day_folder(project_dir: Path, attendance_date: date) -> DayResolution:
    root = Path(project_dir)
    if not root.is_dir():
        return DayResolution("missing", None, "Không có thư mục dự án")
    mappings = {}
    manifest = root / "attendance-period.json"
    if manifest.exists():
        try:
            payload = json.loads(manifest.read_text(encoding="utf-8"))
            if payload.get("version") != 1 or not isinstance(payload.get("mappings"), dict):
                raise ValueError("Invalid mapping manifest")
            mappings = payload["mappings"]
        except (ValueError, OSError, AttributeError):
            return DayResolution("ambiguous", None, "Mapping thư mục cũ không hợp lệ")
    candidates, unmapped = [], False
    for folder in root.iterdir():
        if not folder.is_dir() or folder.is_symlink() or not folder.resolve().is_relative_to(root.resolve()):
            continue
        folder_date = None
        try:
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", folder.name):
                folder_date = date.fromisoformat(folder.name)
            elif re.fullmatch(r"\d{2}-\d{2}-\d{4}", folder.name):
                folder_date = datetime.strptime(folder.name, "%d-%m-%Y").date()
            elif re.fullmatch(r"\d{1,2}", folder.name) and int(folder.name) == attendance_date.day:
                entry = mappings.get(folder.name, {})
                if not isinstance(entry, dict) or not entry.get("confirmed_by") or not entry.get("confirmed_at"):
                    unmapped = True
                else:
                    folder_date = parse_attendance_date(entry.get("date"))
                    if entry.get("migrated_to"):
                        import hashlib
                        hashes = entry.get("source_hashes")
                        target = root / folder_date.isoformat()
                        current = {}
                        for item in folder.rglob("*"):
                            if item.is_symlink():
                                return DayResolution("ambiguous", None, "Nguồn cũ có liên kết cần kiểm tra")
                            if item.is_file():
                                current[item.relative_to(folder).as_posix()] = hashlib.sha256(item.read_bytes()).hexdigest()
                        if entry["migrated_to"] != folder_date.isoformat() or current != hashes or not target.is_dir() or target.is_symlink():
                            return DayResolution("ambiguous", None, "Nguồn cũ thay đổi sau chuyển đổi; cần kiểm tra")
                        continue

        except (ValueError, TypeError):
            if folder.name.isdigit():
                unmapped = True
        if folder_date == attendance_date:
            candidates.append(folder)
    if len(candidates) > 1:
        return DayResolution("ambiguous", None, "Nhiều thư mục cùng ngày; cần hợp nhất dữ liệu")
    if candidates:
        return DayResolution("found", candidates[0])
    if unmapped:
        return DayResolution("legacy_unmapped", None, "Thư mục DD chưa được xác nhận tháng/năm")
    return DayResolution("missing", None, "Không có ảnh của ngày hồ sơ")

def compare_capture_date(attendance_date: date, observed_date: date | None) -> str:
    if observed_date is None:
        return "unknown"
    return "consistent" if observed_date == attendance_date else "mismatch"

def image_date_status(image_path: Path, attendance_date: date) -> str:
    """EXIF/observed text is evidence, never a guarantee of authenticity."""
    from PIL import Image
    image_path = Path(image_path)
    observed = []
    metadata_path = image_path.with_suffix(image_path.suffix + ".json")
    if metadata_path.exists():
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            if metadata.get("derived"):
                return "unknown"
            # A send/request date must never stand in for an observed capture date.
            if metadata.get("original_visible_date"):
                observed.append(parse_attendance_date(metadata["original_visible_date"]))
        except (ValueError, OSError, TypeError, AttributeError):
            return "unknown"
    try:
        with Image.open(image_path) as image:
            exif = image.getexif()
            capture = exif.get(36867)
            if not capture:
                capture = exif.get_ifd(34665).get(36867)
            if isinstance(capture, bytes):
                capture = capture.decode("ascii")
            if capture:
                observed.append(datetime.strptime(capture.rstrip("\x00"), "%Y:%m:%d %H:%M:%S").date())
    except (OSError, ValueError, AttributeError, KeyError, UnicodeError):
        pass
    if any(value != attendance_date for value in observed):
        return "mismatch"
    return "consistent" if observed else "unknown"
