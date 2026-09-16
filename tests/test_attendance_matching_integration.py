from datetime import date
from pathlib import Path
from types import SimpleNamespace
from PIL import Image
import pytest
from src.identity_registry import IdentityRegistry
from src.attendance_matching import match_attendance_record

@pytest.mark.parametrize("observed,status,matched,candidate", [
    ("2026:09:01 08:00:00","consistent",True,False),
    ("2026:08:01 08:00:00","mismatch",False,False),
    (None,"unknown",False,True),
])
def test_pipeline_scopes_identity_and_original_date(tmp_path, observed, status, matched, candidate):
    r=IdentityRegistry(tmp_path/"identity.db",tmp_path/"portraits")
    a=r.register_project("A"); b=r.register_project("B")
    employees=[]
    for project,code in [(a,"001"),(a,"002"),(b,"001")]:
        e=r.create_employee("Same Name");employees.append(e)
        r.assign_employee(project["project_id"],e["employee_id"],code,"2026-01-01")
    day=tmp_path/"camera"/"A"/"2026-09-01";day.mkdir(parents=True)
    image=day/"face.jpg"
    exif=Image.Exif()
    if observed: exif[36867]=observed
    Image.new("RGB",(20,20)).save(image,exif=exif)
    wrong=tmp_path/"camera"/"A"/"2026-08-01";wrong.mkdir()
    Image.new("RGB",(20,20)).save(wrong/"other.jpg")
    calls=[]
    class Matcher:
        def match_employee_in_images(self, **kwargs):
            calls.append(kwargs)
            assert kwargs["project_id"]==a["project_id"]
            assert kwargs["employee_id"]==employees[0]["employee_id"]
            assert kwargs["camera_images"]==[str(image)]
            return SimpleNamespace(status="matched",image_path=str(image),distance=0.1,reason="")
    record={"date":"01/09/2026"}
    match_attendance_record(record,{"id":"001"},project_id=a["project_id"],registry=r,matcher=Matcher(),input_images_dir=day.parent)
    assert record["date_status"]==status
    assert bool(record["matched_image_path"])==matched
    assert bool(record["candidate_image_path"])==candidate
    assert len(calls)==(0 if status=="mismatch" else 1)
    assert record["employee_id"]==employees[0]["employee_id"]
