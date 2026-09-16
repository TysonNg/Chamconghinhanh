"""Join identity, full date and face evidence without inferring missing data."""
from pathlib import Path
from src.attendance_dates import parse_attendance_date, resolve_day_folder, image_date_status

def match_attendance_record(record, person, *, project_id, registry, matcher,
                            input_images_dir, threshold=None, accuracy_mode=True):
    record.update(matched_image_path=None,candidate_image_path=None,
                  identity_status="needs_confirmation",date_status="unknown",
                  face_match_status="not_run",face_distance=None,review_reason="")
    if record.get("record_conflict"):
        record["review_reason"]="Các nguồn chấm công cùng ngày có dữ liệu mâu thuẫn"
        return
    try:
        day=parse_attendance_date(record.get("date"))
    except ValueError:
        record.update(date_status="invalid_date",review_reason="Ngày hồ sơ không hợp lệ")
        return
    if not registry or not project_id:
        record["review_reason"]="Chưa xác nhận mã nhân viên và dự án"
        return
    resolution=registry.resolve_employee(project_id,str(person.get("id") or ""),day)
    record["identity_status"]=resolution.status
    record["employee_id"]=resolution.employee_id
    record["project_id"]=project_id
    if resolution.status!="resolved":
        record["review_reason"]=resolution.reason
        return
    if not input_images_dir:
        record["review_reason"]="Không có thư mục ảnh của dự án"
        return
    folder=resolve_day_folder(Path(input_images_dir),day)
    record["date_status"]=folder.status
    if folder.status!="found":
        record["review_reason"]=folder.reason
        return
    images=sorted(p for p in folder.path.rglob("*")
                  if p.is_file() and not p.is_symlink() and p.resolve().is_relative_to(folder.path.resolve())
                  and p.suffix.lower() in {".jpg",".jpeg",".png",".bmp",".webp"})
    if not images:
        record.update(date_status="missing",review_reason="Thư mục ngày không có ảnh")
        return
    by_date={"consistent":[],"unknown":[],"mismatch":[]}
    for path in images:
        by_date[image_date_status(path,day)].append(str(path))
    if matcher is None:
        record.update(face_match_status="error",review_reason="Bộ nhận diện chưa sẵn sàng")
        return
    for status in ("consistent","unknown"):
        candidates=by_date[status]
        if not candidates:
            continue
        try:
            result=matcher.match_employee_in_images(
                project_id=project_id,employee_id=resolution.employee_id,attendance_date=day,
                camera_images=candidates,distance_threshold=threshold,fast_mode=not accuracy_mode)
        except Exception as exc:
            record.update(face_match_status="error",date_status=status,
                          review_reason=f"Lỗi nhận diện cần kiểm tra: {exc}")
            return
        record.update(face_match_status=result.status,face_distance=result.distance,date_status=status)
        if result.status=="matched":
            if status=="consistent":
                record["matched_image_path"]=result.image_path
            else:
                record["candidate_image_path"]=result.image_path
                record["review_reason"]="Khớp mặt nhưng chưa xác định được ngày chụp ảnh gốc"
            return
        record["review_reason"]=result.reason
    if not by_date["consistent"] and not by_date["unknown"]:
        record.update(date_status="mismatch",review_reason="Ngày trên ảnh gốc khác ngày hồ sơ")
