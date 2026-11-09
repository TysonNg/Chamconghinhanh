"""Join identity, full date and face evidence without inferring missing data."""
from pathlib import Path
from src.supplement_evidence import evidence_is_visible
from src.attendance_dates import parse_attendance_date, resolve_day_folder, attendance_image_date_status
from src.scan_runtime import ScanCancelled, report_scan_activity
from src.face_inference import FaceInferenceError

def match_attendance_record(record, person, **kwargs):
    try:
        return _match_attendance_record(record, person, **kwargs)
    finally:
        # Cancellation must propagate without being reported as a completed result.
        import sys
        if sys.exc_info()[0] is None:
            if record.get('record_conflict'):
                conflict_note = "Các nguồn chấm công cùng ngày có dữ liệu mâu thuẫn; cần kiểm tra giờ công"
                reason = record.get('review_reason') or ''
                record['review_reason'] = '; '.join(filter(None, (reason, conflict_note)))
            status = record.get('face_match_status', 'not_run')
            labels = {'matched': 'Mặt khớp', 'no_match': 'Mặt không khớp',
                      'no_face': 'Không thấy khuôn mặt', 'no_portrait': 'Thiếu chân dung',
                      'error': 'Thất bại', 'identity_unresolved': 'Chưa xác nhận danh tính',
                      'not_run': 'Chưa thể đối soát'}
            report_scan_activity(labels.get(status, status), result_status=status,
                                 day=str(record.get('date') or ''),
                                 image=record.get('matched_image_path') or record.get('candidate_image_path') or '',
                                 reason=record.get('review_reason') or '',
                                 distance=record.get('face_distance'))


def _match_attendance_record(record, person, *, project_id, registry, matcher,
                            input_images_dir, threshold=None, accuracy_mode=False):
    record.update(matched_image_path=None,candidate_image_path=None,
                  identity_status="needs_confirmation",date_status="unknown",
                  face_match_status="not_run",face_distance=None,review_reason="")
    try:
        day=parse_attendance_date(record.get("date"))
    except ValueError:
        record.update(date_status="invalid_date",review_reason="Ngày hồ sơ không hợp lệ")
        return
    report_scan_activity('Đang đối soát ngày', day=day.isoformat(), image='', image_index=0, image_total=0)
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
                  and p.suffix.lower() in {".jpg",".jpeg",".png",".bmp",".webp"} and evidence_is_visible(p))
    if not images:
        record.update(date_status="missing",review_reason="Thư mục ngày không có ảnh")
        return
    by_date={"consistent":[],"message_date":[],"unknown":[],"mismatch":[]}
    for path in images:
        by_date[attendance_image_date_status(path,day)].append(str(path))
    if matcher is None:
        record.update(face_match_status="error",review_reason="Bộ nhận diện chưa sẵn sàng")
        return
    for status in ("consistent","message_date","unknown"):
        candidates=by_date[status]
        if not candidates:
            continue
        try:
            result=matcher.match_employee_in_images(
                project_id=project_id,employee_id=resolution.employee_id,attendance_date=day,
                camera_images=candidates,distance_threshold=threshold,fast_mode=not accuracy_mode)
        except (ScanCancelled, FaceInferenceError):
            raise
        except Exception as exc:
            record.update(face_match_status="error",date_status=status,
                          review_reason=f"Lỗi nhận diện cần kiểm tra: {exc}")
            return
        record.update(face_match_status=result.status,face_distance=result.distance,date_status=status)
        if result.status=="matched":
            if status in ("consistent", "message_date"):
                record["matched_image_path"]=result.image_path
                record["review_reason"] = ""
                if status == "message_date":
                    record["review_reason"] = "Đối soát theo ngày gửi Zalo; chưa xác định ngày chụp gốc"
            else:
                record["candidate_image_path"]=result.image_path
                record["review_reason"]="Khớp mặt nhưng chưa xác định được ngày chụp ảnh gốc"
            return
        record["review_reason"]=result.reason
    if not any(by_date[status] for status in ("consistent", "message_date", "unknown")):
        record.update(date_status="mismatch",review_reason="Ngày trên ảnh gốc khác ngày hồ sơ")
