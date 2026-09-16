from datetime import date
from pathlib import Path
import pytest

@pytest.fixture
def registry(tmp_path):
    from src.identity_registry import IdentityRegistry
    return IdentityRegistry(tmp_path/"identity.sqlite3", tmp_path/"portraits")

def setup_person(r, project="A", code="001", name="Nguyễn Văn An"):
    p=r.register_project(project)
    e=r.create_employee(name)
    r.assign_employee(p["project_id"],e["employee_id"],code,"2026-01-01")
    return p,e

def test_ids_survive_restart_and_project_rename(registry):
    from src.identity_registry import IdentityRegistry
    p,e=setup_person(registry)
    registry.rename_project(p["project_id"],"Tên mới")
    other=IdentityRegistry(registry.db_path,registry.portrait_root)
    assert other.get_project(p["project_id"])["display_name"] == "Tên mới"
    result=other.resolve_employee(p["project_id"],"001",date(2026,9,1))
    assert result.status == "resolved"
    assert result.employee_id == e["employee_id"]
    assert other.resolve_employee(p["project_id"],"1",date(2026,9,1)).status == "missing"

def test_same_name_and_code_are_scoped_to_project(registry):
    pa,ea=setup_person(registry)
    pb,eb=setup_person(registry,"B")
    ec=registry.create_employee(ea["display_name"])
    registry.assign_employee(pa["project_id"],ec["employee_id"],"002","2026-01-01")
    assert registry.resolve_employee(pa["project_id"],"001",date(2026,9,1)).employee_id == ea["employee_id"]
    assert registry.resolve_employee(pb["project_id"],"001",date(2026,9,1)).employee_id == eb["employee_id"]
    assert registry.resolve_employee(pa["project_id"],"",date(2026,9,1)).status == "needs_confirmation"

def test_overlapping_code_fails_without_partial_assignment(registry):
    p,e=setup_person(registry)
    other=registry.create_employee("Other")
    with pytest.raises(ValueError):
        registry.assign_employee(p["project_id"],other["employee_id"],"001","2026-05-01")
    assert registry.resolve_employee(p["project_id"],"001",date(2026,9,1)).employee_id == e["employee_id"]

def test_pending_import_idempotent_and_never_matches_until_confirmed(registry):
    p=registry.register_project("A")
    folder=registry.project_portrait_dir(p["project_id"])/"Nguyễn Văn An"
    folder.mkdir(parents=True)
    (folder/"face.jpg").write_bytes(b"face")
    first=registry.import_legacy(p["project_id"])
    second=registry.import_legacy(p["project_id"])
    assert first[0]["employee_id"] == second[0]["employee_id"]
    assert registry.portrait_paths(p["project_id"],first[0]["employee_id"],date(2026,9,1)) == []
    registry.confirm_source(p["project_id"],first[0]["employee_id"],"001","2026-01-01","Reviewer")
    assert registry.portrait_paths(p["project_id"],first[0]["employee_id"],date(2026,9,1)) == [folder/"face.jpg"]

def test_portraits_never_escape_project(registry):
    p,e=setup_person(registry)
    registry.project_portrait_dir(p["project_id"]).mkdir(parents=True)
    with pytest.raises(ValueError):
        registry.bind_portrait(p["project_id"],e["employee_id"],"../B/face.jpg","Reviewer")

def test_transfer_keeps_identity_and_historical_membership(registry):
    pa,e=setup_person(registry)
    pb=registry.register_project("B")
    folder=registry.project_portrait_dir(pa["project_id"])/e["employee_id"]
    folder.mkdir(parents=True)
    (folder/"face.jpg").write_bytes(b"portrait")
    registry.bind_portrait(pa["project_id"],e["employee_id"],e["employee_id"],"Reviewer")
    registry.transfer_employee(pa["project_id"],pb["project_id"],e["employee_id"],"2026-09-16","B001","Reviewer")
    assert registry.resolve_employee(pa["project_id"],"001",date(2026,9,15)).employee_id == e["employee_id"]
    assert registry.resolve_employee(pa["project_id"],"001",date(2026,9,16)).status == "missing"
    assert registry.resolve_employee(pb["project_id"],"B001",date(2026,9,16)).employee_id == e["employee_id"]
    assert (folder/"face.jpg").exists()
    assert registry.portrait_paths(pa["project_id"],e["employee_id"],date(2026,9,15)) == [folder/"face.jpg"]
    assert registry.portrait_paths(pb["project_id"],e["employee_id"],date(2026,9,16))[0].read_bytes() == b"portrait"

def test_transfer_conflict_preserves_source_membership(registry):
    pa,e=setup_person(registry)
    pb,other=setup_person(registry,"B","B001")
    with pytest.raises(ValueError):
        registry.transfer_employee(pa["project_id"],pb["project_id"],e["employee_id"],"2026-09-16","B001","Reviewer")
    assert registry.resolve_employee(pa["project_id"],"001",date(2026,9,17)).employee_id == e["employee_id"]

def test_referenced_project_is_archived_not_deleted(registry):
    p,e=setup_person(registry)
    registry.archive_project(p["project_id"])
    assert not registry.get_project(p["project_id"])["active"]
    assert registry.resolve_employee(p["project_id"],"001",date(2026,9,1)).employee_id == e["employee_id"]
