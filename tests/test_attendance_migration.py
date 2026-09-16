import json
from datetime import date
import pytest
from src.attendance_dates import resolve_day_folder

def setup(tmp_path):
    root=tmp_path/"A";(root/"01").mkdir(parents=True)
    (root/"01"/"photo.jpg").write_bytes(b"original")
    (root/"attendance-period.json").write_text(json.dumps({"version":1,"mappings":{"01":{
        "date":"2026-09-01","confirmed_by":"Reviewer","confirmed_at":"2026-09-16"}}}),encoding="utf-8")
    return root

def test_preview_does_not_write_and_apply_preserves_source(tmp_path):
    from tools.preview_attendance_migration import migrate_project
    root=setup(tmp_path)
    report=migrate_project(root)
    assert report["files"][0]["sha256"]
    assert not (root/"2026-09-01").exists()
    migrate_project(root,apply=True)
    assert (root/"01"/"photo.jpg").read_bytes()==b"original"
    assert (root/"2026-09-01"/"photo.jpg").read_bytes()==b"original"
    assert resolve_day_folder(root,date(2026,9,1)).path==root/"2026-09-01"
    migrate_project(root,apply=True)
    assert len(list((root/"2026-09-01").iterdir()))==1
    assert list(root.glob("attendance-period.*.backup.json"))

def test_collision_rejected_without_overwrite(tmp_path):
    from tools.preview_attendance_migration import migrate_project
    root=setup(tmp_path);(root/"2026-09-01").mkdir()
    (root/"2026-09-01"/"photo.jpg").write_bytes(b"other")
    with pytest.raises(ValueError): migrate_project(root,apply=True)
    assert (root/"2026-09-01"/"photo.jpg").read_bytes()==b"other"

def test_changed_legacy_source_requires_review(tmp_path):
    from tools.preview_attendance_migration import migrate_project
    root=setup(tmp_path);migrate_project(root,apply=True)
    (root/"01"/"new.jpg").write_bytes(b"new")
    assert resolve_day_folder(root,date(2026,9,1)).status=="ambiguous"
