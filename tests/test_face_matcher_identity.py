from datetime import date
import numpy as np
from src.face_matcher import FaceMatcher

def test_no_fuzzy_or_cross_project_portrait_fallback(tmp_path):
    for project, name in (("A","Nguyen Van An"),("B","Nguyen Van Anh")):
        folder=tmp_path/project/name
        folder.mkdir(parents=True)
        (folder/"p.jpg").write_bytes(b"portrait")
    matcher=FaceMatcher(str(tmp_path))
    assert matcher.find_portraits("Nguyen Van Anh",project_name="A") == []
    assert matcher.find_portraits("Nguyen Van An",project_name="Missing") == []
    assert matcher.find_portraits("Nguyen Van An") == []

def test_identity_match_uses_only_bound_project_portraits(tmp_path,monkeypatch):
    from src.identity_registry import IdentityRegistry
    r=IdentityRegistry(tmp_path/"ids.db", tmp_path/"portraits")
    pa=r.register_project("A")
    pb=r.register_project("B")
    ea=r.create_employee("Same Name")
    eb=r.create_employee("Same Name")
    for p,e in ((pa,ea),(pb,eb)):
        r.assign_employee(p["project_id"],e["employee_id"],"001","2026-01-01")
        folder=r.project_portrait_dir(p["project_id"])/e["employee_id"]
        folder.mkdir(parents=True)
        (folder/"p.jpg").write_bytes(b"portrait")
        r.bind_portrait(p["project_id"],e["employee_id"],e["employee_id"],"Reviewer")
    camera=tmp_path/"camera.jpg"; camera.write_bytes(b"camera")
    matcher=FaceMatcher(str(r.portrait_root),identity_registry=r)
    def embedding(p):
        return np.array([1.,0.]) if ea["employee_id"] in str(p) or str(p)==str(camera) else np.array([0.,1.])
    monkeypatch.setattr(matcher,"_get_embedding",embedding)
    good=matcher.match_employee_in_images(project_id=pa["project_id"],employee_id=ea["employee_id"],
        attendance_date=date(2026,9,1),camera_images=[str(camera)])
    assert good.status == "matched"
    assert good.image_path == str(camera)
    wrong=matcher.match_employee_in_images(project_id=pa["project_id"],employee_id=eb["employee_id"],
        attendance_date=date(2026,9,1),camera_images=[str(camera)])
    assert wrong.status == "identity_unresolved"
    assert wrong.image_path is None
