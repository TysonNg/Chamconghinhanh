"""Merge records by verified identity/code and period, never by display name."""
from copy import deepcopy
import hashlib
import json
from src.attendance_dates import parse_attendance_date

def report_identity_suffix(person):
    value = person.get("employee_id") or person.get("id")
    if value:
        # Hash avoids unsafe codes and collisions after filename sanitization.
        return hashlib.sha256(str(value).encode("utf-8")).hexdigest()[:16]
    raw = person.get("source_id") or json.dumps(person,ensure_ascii=False,sort_keys=True,default=str)
    return "pending_" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

def merge_attendance_people(persons, project_id="", registry=None):
    grouped = {}
    for index, source in enumerate(persons):
        code = str(source.get("id") or "").strip()
        source_id = source.get("source_id") or f"source:{index}"
        for original in source.get("records", []):
            record = deepcopy(original)
            try:
                day = parse_attendance_date(record["date"])
                period = (day.year, day.month)
                record["date"] = day.strftime("%d/%m/%Y")
            except (ValueError, KeyError):
                day = None
                period = (source.get("year"), source.get("month"))
                record.update(date_status="invalid_date", review_reason="Ngày hồ sơ không hợp lệ")
            employee_id, status = None, "needs_confirmation"
            if registry and project_id and day:
                resolution = registry.resolve_employee(project_id,code,day)
                employee_id, status = resolution.employee_id,resolution.status
            identity = employee_id or (("code",code) if code and registry is None else ("source",source_id,index))
            key = (project_id,identity,period)
            if key not in grouped:
                item = {k:deepcopy(v) for k,v in source.items() if k!="records"}
                item.update(records=[],source_id=source_id,project_id=project_id,employee_id=employee_id,
                            identity_status=status,year=period[0],month=period[1],raw_name=source.get("raw_name",source.get("name","")))
                if employee_id:
                    item["name"] = registry.get_employee(employee_id)["display_name"]
                grouped[key] = item
            person = grouped[key]
            record.update(project_id=project_id,employee_id=employee_id,identity_status=status)
            duplicate = next((r for r in person["records"] if r.get("date")==record.get("date")),None)
            if duplicate is None:
                person["records"].append(record)
                continue
            fields = ("gio_vao","gio_ra","is_absent","missing_checkin","missing_checkout","same_in_out","under_3h")
            if all(duplicate.get(k)==record.get(k) for k in fields):
                continue
            if not duplicate.get("record_conflict"):
                duplicate["conflicting_records"] = [{k:deepcopy(v) for k,v in duplicate.items() if k!="conflicting_records"}]
            duplicate["conflicting_records"].append(record)
            duplicate["record_conflict"] = True
            duplicate["review_reason"] = "Các nguồn chấm công cùng ngày có dữ liệu mâu thuẫn"
    return list(grouped.values())
