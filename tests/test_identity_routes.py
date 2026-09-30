import io
from datetime import date
from flask import Flask
from PIL import Image
import pytest
from src.identity_registry import IdentityRegistry

def client(tmp_path):
    from src.identity_routes import register_identity_routes
    r=IdentityRegistry(tmp_path/"ids.db",tmp_path/"portraits")
    app=Flask(__name__); app.testing=True
    register_identity_routes(app,lambda:r,lambda:tmp_path/"camera")
    return app.test_client(),r

def photo():
    b=io.BytesIO(); Image.new("RGB",(12,12),"blue").save(b,"PNG"); b.seek(0); return b

def test_existing_portraits_require_confirmation_and_use_ids(tmp_path):
    c,r=client(tmp_path)
    folder=r.portrait_root/"Site"/"Same Name"
    folder.mkdir(parents=True); (folder/"p.png").write_bytes(photo().getvalue())
    projects=c.get("/api/projects").json["projects"]
    pid=projects[0]["project_id"]
    employees=c.get("/api/portraits",query_string={"project_id":pid}).json["employees"]
    emp=employees[0]
    assert emp["identity_status"]=="needs_confirmation"
    assert c.get(emp["avatar_url"]).status_code==200
    response=c.post("/api/portraits/employee/confirm",json={"project_id":pid,
        "employee_id":emp["employee_id"],"payroll_code":"001","valid_from":"2026-01-01","reviewer":"Reviewer"})
    assert response.status_code==200
    assert r.resolve_employee(pid,"001",date(2026,9,1)).employee_id==emp["employee_id"]

def test_create_same_name_different_ids_and_upload_uses_id(tmp_path):
    c,r=client(tmp_path)
    p=c.post("/api/projects/create",json={"name":"Site"}).json
    employees=[]
    for code in ("001","002"):
        response=c.post("/api/portraits/employee/create",json={"project_id":p["project_id"],"name":"Same Name",
            "payroll_code":code,"valid_from":"2026-01-01","reviewer":"Reviewer"})
        assert response.status_code==200
        employees.append(response.json["employee_id"])
    assert employees[0]!=employees[1]
    response=c.post("/api/portraits/employee/upload",data={"project_id":p["project_id"],"employee_id":employees[1],
        "reviewer":"Reviewer","files":(photo(),"portrait.png")})
    assert response.status_code==200
    assert r.portrait_paths(p["project_id"],employees[0],date(2026,9,1))==[]
    assert len(r.portrait_paths(p["project_id"],employees[1],date(2026,9,1)))==1
    # An old name-only caller cannot select between the employees.
    response=c.post("/api/portraits/employee/upload",data={"project_id":p["project_id"],"name":"Same Name",
        "files":(photo(),"portrait.png")})
    assert response.status_code==400

def test_project_rename_preserves_storage_and_identity(tmp_path):
    c,r=client(tmp_path)
    p=c.post("/api/projects/create",json={"name":"Site"}).json
    response=c.post("/api/projects/rename",json={"project_id":p["project_id"],"new_name":"Renamed"})
    assert response.status_code==200
    row=r.get_project(p["project_id"])
    assert row["storage_dir"]=="Site"
    assert row["display_name"]=="Renamed"
    assert (r.portrait_root/"Site").exists()

def test_cross_project_employee_upload_rejected(tmp_path):
    c,r=client(tmp_path)
    a=c.post("/api/projects/create",json={"name":"A"}).json
    b=c.post("/api/projects/create",json={"name":"B"}).json
    e=c.post("/api/portraits/employee/create",json={"project_id":a["project_id"],"name":"One",
        "payroll_code":"001","valid_from":"2026-01-01","reviewer":"Reviewer"}).json
    result=c.post("/api/portraits/employee/upload",data={"project_id":b["project_id"],"employee_id":e["employee_id"],
        "reviewer":"Reviewer","files":(photo(),"photo.png")})
    assert result.status_code==400
    assert not list((r.portrait_root/"B").rglob("*.png"))

def test_archiving_project_keeps_photos(tmp_path):
    c,r=client(tmp_path)
    p=c.post("/api/projects/create",json={"name":"Site"}).json
    f=r.portrait_root/"Site"/"old.png"; f.write_bytes(photo().getvalue())
    assert c.post("/api/projects/delete",json={"project_id":p["project_id"]}).status_code==200
    assert f.exists()
    assert c.get("/api/projects").json["projects"]==[]

def test_create_employee_assigns_internal_code(tmp_path):
    c, r = client(tmp_path)
    p = c.post("/api/projects/create", json={"name": "Site"}).json
    res = c.post("/api/portraits/employee/create", json={
        "project_id": p["project_id"],
        "name": "Nguyen Van Test",
        "payroll_code": "0099",
        "valid_from": "2026-01-01",
        "reviewer": "Reviewer"
    })
    assert res.status_code == 200
    data = res.json
    assert data["success"] is True
    assert "internal_code" in data
    assert data["internal_code"].startswith("NV-")
    assert len(data["internal_code"]) == 11

    # Check via list portraits endpoint
    portraits = c.get("/api/portraits", query_string={"project_id": p["project_id"]}).json
    emp = next(e for e in portraits["employees"] if e["employee_id"] == data["employee_id"])
    assert emp["internal_code"] == data["internal_code"]


def test_create_multiple_employees_without_payroll_codes_and_upload_without_reviewer(tmp_path):
    c, r = client(tmp_path)
    project = c.post("/api/projects/create", json={"name": "Site"}).json

    employees = []
    for name in ("Nguyen Van A", "Tran Van B"):
        response = c.post("/api/portraits/employee/create", json={
            "project_id": project["project_id"],
            "name": name,
            "payroll_code": ""
        })
        assert response.status_code == 200
        employees.append(response.json["employee_id"])

    listed = c.get("/api/portraits", query_string={"project_id": project["project_id"]}).json["employees"]
    assert {employee["employee_id"] for employee in listed} == set(employees)
    assert all(employee["payroll_code"] == "" for employee in listed)

    uploaded = c.post("/api/portraits/employee/upload", data={
        "project_id": project["project_id"],
        "employee_id": employees[0],
        "files": (photo(), "portrait.png")
    })
    assert uploaded.status_code == 200
    assert len(r.portrait_paths(project["project_id"], employees[0], date.today())) == 1

def test_delete_employee_photo_physically_removes_file_and_excludes(tmp_path):
    c, r = client(tmp_path)
    project = c.post("/api/projects/create", json={"name": "Site"}).json
    emp = c.post("/api/portraits/employee/create", json={
        "project_id": project["project_id"],
        "name": "Le Van C"
    }).json
    eid = emp["employee_id"]

    c.post("/api/portraits/employee/upload", data={
        "project_id": project["project_id"],
        "employee_id": eid,
        "files": (photo(), "test_face.png")
    })
    portraits = c.get("/api/portraits", query_string={"project_id": project["project_id"]}).json
    emp_data = next(e for e in portraits["employees"] if e["employee_id"] == eid)
    assert len(emp_data["images"]) == 1
    image_name = emp_data["images"][0]

    # Verify physical file exists on disk
    photo_file = r.portrait_root / "Site" / image_name
    assert photo_file.exists()

    # Call delete-photo
    del_res = c.post("/api/portraits/employee/delete-photo", json={
        "project_id": project["project_id"],
        "employee_id": eid,
        "filename": image_name
    })
    assert del_res.status_code == 200
    assert del_res.json["success"] is True

    # Physical file must be removed
    assert not photo_file.exists()

    # Should no longer be returned in portraits listing
    portraits_after = c.get("/api/portraits", query_string={"project_id": project["project_id"]}).json
    emp_after = next(e for e in portraits_after["employees"] if e["employee_id"] == eid)
    assert len(emp_after["images"]) == 0

def test_open_portrait_folder_endpoint(tmp_path, monkeypatch):
    c, r = client(tmp_path)
    opened = []
    import sys
    if sys.platform == "win32":
        monkeypatch.setattr("os.startfile", lambda p: opened.append(p))
    project = c.post("/api/projects/create", json={"name": "Site"}).json

    # Project folder
    res1 = c.post("/api/portraits/open-folder", json={"project_id": project["project_id"]})
    assert res1.status_code == 200
    assert res1.json["success"] is True
    assert (r.portrait_root / "Site").exists()

    # Specific employee folder
    emp = c.post("/api/portraits/employee/create", json={
        "project_id": project["project_id"],
        "name": "Dang Van D"
    }).json
    res2 = c.post("/api/portraits/open-folder", json={
        "project_id": project["project_id"],
        "employee_id": emp["employee_id"]
    })
    assert res2.status_code == 200
    assert res2.json["success"] is True
    assert (r.portrait_root / "Site" / emp["employee_id"]).exists()


