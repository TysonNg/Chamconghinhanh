"""Daily photo endpoints with strict calendar and project boundaries."""
import calendar
import io
import json
import os
import re
import threading
import uuid
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from flask import Blueprint, jsonify, request, send_file
from PIL import Image
from src.attendance_dates import canonical_day_path, parse_attendance_date, resolve_day_folder

_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif"}
_manifest_lock = threading.Lock()

def safe_component(value):
    if not isinstance(value, str) or not value.strip() or value.strip() in (".", "..") or re.search(r'[<>:"/\\|?*\x00-\x1f]', value):
        raise ValueError("Tên dự án hoặc tệp không hợp lệ")
    return value.strip()

def register_daily_photo_routes(app, input_root, project_resolver=None):
    bp = Blueprint("daily_photos", __name__)

    @bp.errorhandler(ValueError)
    def invalid(exc):
        return jsonify(success=False, error=str(exc)), 400

    @bp.errorhandler(RuntimeError)
    def conflict(exc):
        return jsonify(success=False, error=str(exc)), 409

    def project_path(values):
        reference = values.get("project_id") or values.get("project", "")
        if project_resolver:
            selected = project_resolver(reference)
            if not selected["active"]:
                raise ValueError("Dự án đã được lưu trữ")
            project = selected["storage_dir"]
        else:
            project = safe_component(reference)
        root = Path(input_root()).resolve()
        path = root / project
        if not path.resolve().is_relative_to(root):
            raise ValueError("Đường dẫn dự án không hợp lệ")
        return project, path

    def requested_day(values, fallback=None):
        value = values.get("date") or fallback or values.get("day")
        day = parse_attendance_date(value)
        if str(value).strip() != day.isoformat():
            raise ValueError("API yêu cầu ngày YYYY-MM-DD")
        return day

    def folder_for(values, day, write=False):
        project, root = project_path(values)
        result = resolve_day_folder(root, day)
        if result.status not in ("found", "missing"):
            raise RuntimeError(result.reason)
        # Never split a mapped legacy day's data across two directories.
        folder = result.path if result.path else canonical_day_path(root, day)
        if write:
            folder.mkdir(parents=True, exist_ok=True)
        return project, folder

    def images(folder):
        if not folder.is_dir():
            return []
        return sorted(p for p in folder.iterdir()
                      if p.is_file() and not p.is_symlink() and p.suffix.lower() in _EXTENSIONS)

    @bp.get("/api/photos/daily")
    def stats():
        project, root = project_path(request.args)
        period = request.args.get("period", "")
        if not re.fullmatch(r"\d{4}-\d{2}", period):
            raise ValueError("Chọn tháng/năm theo định dạng YYYY-MM")
        first = parse_attendance_date(period + "-01")
        days = []
        for d in range(1, calendar.monthrange(first.year, first.month)[1] + 1):
            day = date(first.year, first.month, d)
            result = resolve_day_folder(root, day)
            count = len(images(result.path)) if result.path else 0
            days.append({"day": f"{d:02}", "date": day.isoformat(), "image_count": count,
                         "has_images": bool(count), "status": result.status, "reason": result.reason})
        legacy = [p.name for p in root.iterdir() if p.is_dir() and re.fullmatch(r"\d{1,2}", p.name)] if root.exists() else []
        return jsonify(success=True, project=project, period=period, days=days, legacy_folders=sorted(legacy))

    @bp.get("/api/photos/daily/<day>")
    def listing(day):
        target = requested_day(request.args, day)
        project, folder = folder_for(request.args, target)
        photos = [{"filename": p.name, "size_kb": round(p.stat().st_size / 1024, 1),
                   "url": "/api/photos/view/daily?" + urlencode({"project": project, "date": target.isoformat(), "filename": p.name})}
                  for p in images(folder)]
        return jsonify(success=True, project=project, date=target.isoformat(), day=f"{target.day:02}",
                       photos=photos, total=len(photos))

    @bp.get("/api/photos/view/daily")
    def view():
        name = safe_component(request.args.get("filename"))
        _, folder = folder_for(request.args, requested_day(request.args))
        p = folder / name
        if p not in images(folder):
            return jsonify(error="Không tìm thấy ảnh"), 404
        return send_file(p)

    @bp.post("/api/photos/daily/upload")
    def upload():
        target = requested_day(request.form)
        project, root = project_path(request.form)
        files = request.files.getlist("files") or request.files.getlist("photos")
        if not files:
            raise ValueError("Không có ảnh")
        prepared = []
        for file in files:
            name = safe_component(file.filename)
            if Path(name).suffix.lower() not in _EXTENSIONS:
                raise ValueError("Định dạng ảnh không hỗ trợ")
            raw = file.read(20 * 1024 * 1024 + 1)
            if len(raw) > 20 * 1024 * 1024:
                raise ValueError("Ảnh tối đa 20MB")
            try:
                with Image.open(io.BytesIO(raw)) as image:
                    if image.width * image.height > 40000000:
                        raise ValueError("Ảnh vượt quá 40 megapixel")
                    image.verify()
            except OSError as exc:
                raise ValueError("Tệp ảnh không hợp lệ") from exc
            prepared.append((name, raw))
        # New writes always use the ISO folder; a legacy mapping must be migrated first.
        resolution = resolve_day_folder(root, target)
        if resolution.status not in ("missing", "found"):
            raise RuntimeError(resolution.reason)
        destination = canonical_day_path(root, target)
        if resolution.path and resolution.path != destination:
            raise RuntimeError("Chuyển thư mục cũ sang YYYY-MM-DD trước khi tải thêm ảnh")
        destination.mkdir(parents=True, exist_ok=True)
        saved = []
        for name, raw in prepared:
            p = destination / name
            while True:
                try:
                    with p.open("xb") as f:
                        f.write(raw)
                    break
                except FileExistsError:
                    p = destination / f"{Path(name).stem}_{uuid.uuid4().hex[:12]}{Path(name).suffix}"
            saved.append(p.name)
        return jsonify(success=True, project=project, date=target.isoformat(), saved_count=len(saved), files=saved)

    @bp.post("/api/photos/daily/delete")
    def delete():
        data = request.get_json()
        if not isinstance(data, dict):
            raise ValueError("Yêu cầu phải là JSON object")
        name = safe_component(data["filename"]) if data.get("filename") else None
        _, folder = folder_for(data, requested_day(data))
        if data.get("delete_all") is not True and not name:
            raise ValueError("Chọn ảnh cần xóa")
        selected = images(folder) if data.get("delete_all") is True else [p for p in images(folder) if p.name == name]
        for p in selected:
            p.unlink()
            p.with_suffix(p.suffix + ".json").unlink(missing_ok=True)
        return jsonify(success=True, deleted_count=len(selected), message=f"Đã xóa {len(selected)} ảnh")

    @bp.post("/api/photos/daily/legacy-map")
    def legacy_map():
        data = request.get_json()
        if not isinstance(data, dict):
            raise ValueError("Yêu cầu phải là JSON object")
        _, root = project_path(data)
        folder = data.get("folder", "")
        day = requested_day(data)
        reviewer = str(data.get("reviewer") or "").strip()
        if not re.fullmatch(r"\d{1,2}", folder) or int(folder) != day.day:
            raise ValueError("Thư mục DD phải khớp ngày")
        if not (root / folder).is_dir() or (root / folder).is_symlink():
            raise ValueError("Không tìm thấy thư mục cũ")
        if data.get("confirm_single_period") is not True or not reviewer:
            raise ValueError("Cần xác nhận mọi ảnh thuộc cùng ngày và ghi người xác nhận")
        path = root / "attendance-period.json"
        with _manifest_lock:
            payload = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {"version": 1, "mappings": {}}
            previous = payload["mappings"].get(folder)
            if previous and previous.get("date") != day.isoformat():
                raise RuntimeError("Thư mục đã được gán ngày khác; kiểm tra dữ liệu trước khi thay đổi")
            if previous and previous.get("migrated_to"):
                raise ValueError("Thư mục đã chuyển đổi; không ghi đè lịch sử mapping")
            payload["mappings"][folder] = {"date": day.isoformat(), "confirmed_by": reviewer,
                                           "confirmed_at": datetime.now(timezone.utc).isoformat()}
            temp = path.with_suffix(".json.tmp")
            temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(temp, path)
        return jsonify(success=True, mapping=payload["mappings"][folder])

    app.register_blueprint(bp)
