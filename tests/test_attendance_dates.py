import json
from datetime import date
import pytest
from src.excel_extractor import ExcelToWordExporter

def test_exporter_does_not_use_same_day_from_other_month(tmp_path):
    (tmp_path / "2026-08-01").mkdir()
    exporter = object.__new__(ExcelToWordExporter)
    exporter.input_images_dir = str(tmp_path)
    assert exporter._find_day_folder("01", "01/09/2026") is None

@pytest.mark.parametrize("names,requested,status,chosen", [
    (["2026-08-01"], "2026-09-01", "missing", None),
    (["2025-09-01"], "2026-09-01", "missing", None),
    (["2026-09-01"], "2026-09-01", "found", "2026-09-01"),
    (["01-09-2026"], "2026-09-01", "found", "01-09-2026"),
    (["01"], "2026-09-01", "legacy_unmapped", None),
    (["2026-09-01","01-09-2026"], "2026-09-01", "ambiguous", None),
])
def test_full_date_resolution(tmp_path, names, requested, status, chosen):
    from src.attendance_dates import resolve_day_folder
    for name in names:
        (tmp_path / name).mkdir()
    result = resolve_day_folder(tmp_path, date.fromisoformat(requested))
    assert result.status == status
    assert result.path == (tmp_path / chosen if chosen else None)

def test_legacy_mapping_is_specific_to_folder_and_period(tmp_path):
    from src.attendance_dates import resolve_day_folder
    (tmp_path / "01").mkdir()
    (tmp_path / "attendance-period.json").write_text(json.dumps({
        "version": 1, "mappings": {"01": {"date": "2026-09-01", "confirmed_by": "Reviewer", "confirmed_at": "2026-09-16T00:00:00Z"}}
    }), encoding="utf-8")
    assert resolve_day_folder(tmp_path, date(2026,9,1)).path == tmp_path / "01"
    assert resolve_day_folder(tmp_path, date(2026,8,1)).path is None

def test_no_parent_fallback(tmp_path):
    from src.attendance_dates import resolve_day_folder
    (tmp_path / "2026-09-01").mkdir()
    project = tmp_path / "Project"
    project.mkdir()
    assert resolve_day_folder(project, date(2026,9,1)).status == "missing"

@pytest.mark.parametrize("value", ["01", "2026-02-29", "31/04/2026", "", "20260901", "2026-9-1"])
def test_invalid_dates(value):
    from src.attendance_dates import parse_attendance_date
    with pytest.raises(ValueError):
        parse_attendance_date(value)

def test_leap_day_and_observed_date():
    from src.attendance_dates import parse_attendance_date, compare_capture_date
    assert parse_attendance_date("29/02/2024") == date(2024,2,29)
    assert compare_capture_date(date(2026,9,1), date(2026,8,1)) == "mismatch"
    assert compare_capture_date(date(2026,9,1), None) == "unknown"
    assert compare_capture_date(date(2026,9,1), date(2026,9,1)) == "consistent"

def test_exif_uses_capture_date_not_file_modified(tmp_path):
    from PIL import Image
    from src.attendance_dates import image_date_status
    p = tmp_path / "photo.jpg"
    exif = Image.Exif()
    exif[306] = "2026:09:01 08:00:00"
    Image.new("RGB", (10,10)).save(p, exif=exif)
    assert image_date_status(p, date(2026,9,1)) == "unknown"
    exif[36867] = "2026:08:01 08:00:00"
    Image.new("RGB", (10,10)).save(p, exif=exif)
    assert image_date_status(p, date(2026,9,1)) == "mismatch"

def test_modified_image_metadata_cannot_confirm_capture_date(tmp_path):
    from PIL import Image
    from src.attendance_dates import image_date_status
    p = tmp_path / "photo.jpg"
    exif = Image.Exif()
    exif[36867] = "2026:09:01 08:00:00"
    Image.new("RGB", (10,10)).save(p, exif=exif)
    p.with_suffix(".jpg.json").write_text(json.dumps({"derived": True, "requested_date": "2026-09-01"}))
    assert image_date_status(p, date(2026,9,1)) == "unknown"
