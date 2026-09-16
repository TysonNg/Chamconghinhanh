"""Project and portrait management using persistent IDs; legacy photos stay pending."""
from datetime import date
from pathlib import Path
from urllib.parse import urlencode
import io
import re
import uuid
from flask import Blueprint, jsonify, request, send_file
from PIL import Image
from src.daily_photo_routes import safe_component

_EXT = {".jpg",".jpeg",".png",".bmp",".webp"}

def register_identity_routes(app, registry_provider, input_root):
    bp=Blueprint("identities",__name__)

    @bp.errorhandler(ValueError)
    def invalid(exc):
        return jsonify(success=False,error=str(exc)),400

    def registry():
        return registry_provider()

    def project(values):
        p=registry().get_project(values.get("project_id") or values.get("project") or values.get("name") or "")
        if not p["active"]:
            raise ValueError("Dự án đã được lưu trữ")
        return p

    def payload():
        value=request.get_json()
        if not isinstance(value,dict):
            raise ValueError("Yêu cầu phải là JSON object")
        return value

    def employee(values,p):
        eid=values.get("employee_id")
        if not eid:
            raise ValueError("Chọn nhân viên bằng ID; tên không đủ để xác định danh tính")
        if not any(item["employee_id"]==eid for item in registry().list_employees(p["project_id"])):
            raise ValueError("Nhân viên không thuộc dự án")
        return registry().get_employee(eid)

    def display_paths(p,eid):
        r=registry()
        root=r.project_portrait_dir(p["project_id"])
        with r._connect() as c:
            rows=c.execute("""SELECT relative_path FROM legacy_sources WHERE project_id=? AND employee_id=?
                UNION SELECT relative_path FROM portrait_bindings WHERE project_id=? AND employee_id=?""",
                (p["project_id"],eid,p["project_id"],eid)).fetchall()
            exclusions={row[0] for row in c.execute("SELECT relative_path FROM portrait_exclusions WHERE project_id=? AND employee_id=?",
                                                   (p["project_id"],eid))}
        paths=[]
        for row in rows:
            bound=r._bound_path(p["project_id"],row[0])
            candidates=bound.rglob("*") if bound.is_dir() else [bound]
            paths.extend(x for x in candidates if x.is_file() and not x.is_symlink()
                         and x.resolve().is_relative_to(root.resolve()) and x.suffix.lower() in _EXT
                         and x.relative_to(root).as_posix() not in exclusions)
        return sorted(set(paths))

    def discover_projects():
        r=registry()
        for root in (r.portrait_root,Path(input_root())):
            if root.is_dir():
                for p in sorted(root.iterdir()):
                    if p.is_dir() and not p.is_symlink() and not re.fullmatch(r"\d{1,2}",p.name):
                        r.register_project(p.name)
        for p in r.list_projects():
            r.import_legacy(p["project_id"])
        return r

    @bp.post("/api/portraits/internal-codes/generate")
    def generate_internal_codes():
        r=discover_projects()
        result=r.generate_internal_codes()
        return jsonify(success=True,**result)

    @bp.get("/api/projects")
    def projects():
        r=discover_projects()
        items=[]
        for p in r.list_projects():
            r.import_legacy(p["project_id"])
            camera=Path(input_root())/p["storage_dir"]
            photos=[x for x in camera.rglob("*") if x.is_file() and x.suffix.lower() in _EXT] if camera.exists() else []
            items.append({**p,"name":p["storage_dir"],"employee_count":len([e for e in r.list_employees(p["project_id"]) if e["active"]]),
                          "camera_days_count":len({x.parent.name for x in photos}),"camera_total_images":len(photos)})
        return jsonify(success=True,projects=items,default_project=items[0]["name"] if items else "")

    @bp.post("/api/projects/create")
    def create_project():
        p=registry().register_project(payload().get("name"))
        if not p["active"]:
            raise ValueError("Thư mục dự án này đã được lưu trữ")
        registry().project_portrait_dir(p["project_id"]).mkdir(parents=True,exist_ok=True)
        (Path(input_root())/p["storage_dir"]).mkdir(parents=True,exist_ok=True)
        return jsonify(success=True,project=p["storage_dir"],project_id=p["project_id"])

    @bp.post("/api/projects/rename")
    def rename_project():
        data=payload()
        p=registry().get_project(data.get("project_id") or data.get("old_name"))
        registry().rename_project(p["project_id"],data.get("new_name"))
        return jsonify(success=True,name=p["storage_dir"],project_id=p["project_id"],display_name=data["new_name"])

    @bp.post("/api/projects/delete")
    def archive_project():
        p=project(payload())
        registry().archive_project(p["project_id"])
        return jsonify(success=True,message="Đã lưu trữ dự án; giữ nguyên ảnh và lịch sử")

    @bp.get("/api/portraits")
    def portraits():
        p=project(request.args); r=registry()
        r.import_legacy(p["project_id"])
        items=[]
        search=request.args.get("search","").casefold()
        for e in r.list_employees(p["project_id"]):
            if not e["active"]:
                continue
            code=e["memberships"][0]["payroll_code"] if e["memberships"] else ""
            if search and search not in e["display_name"].casefold() and search not in code.casefold():
                continue
            paths=display_paths(p,e["employee_id"])
            root=r.project_portrait_dir(p["project_id"])
            images=[x.relative_to(root).as_posix() for x in paths]
            urls={name:"/api/photos/view/portrait?"+urlencode({"project_id":p["project_id"],
                   "employee_id":e["employee_id"],"filename":name}) for name in images}
            items.append({**e,"name":e["display_name"],"project":p["storage_dir"],"payroll_code":code,
                          "identity_status":e["status"],"images":images,"image_urls":urls,
                          "image_count":len(images),"avatar_url":urls[images[0]] if images else ""})
        return jsonify(success=True,project=p["storage_dir"],project_id=p["project_id"],employees=items,total=len(items))

    @bp.post("/api/portraits/employee/confirm")
    def confirm():
        data=payload(); p=project(data); e=employee(data,p)
        registry().confirm_source(p["project_id"],e["employee_id"],data.get("payroll_code"),data.get("valid_from"),data.get("reviewer"))
        return jsonify(success=True,employee_id=e["employee_id"])

    @bp.post("/api/portraits/employee/create")
    def create_employee():
        from src.identity_registry import _day
        data=payload(); p=project(data); r=registry()
        name=str(data.get("name") or "").strip()
        reviewer=str(data.get("reviewer") or "").strip()
        if not name or not reviewer:
            raise ValueError("Nhập tên nhân viên và người xác nhận")
        eid=uuid.uuid4().hex
        with r._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,name))
            r._assign(c,p["project_id"],eid,data.get("payroll_code"),_day(data.get("valid_from")),None)
        folder=r.project_portrait_dir(p["project_id"])/eid
        folder.mkdir(parents=True,exist_ok=True)
        r.bind_portrait(p["project_id"],eid,eid,reviewer)
        return jsonify(success=True,name=name,employee_id=eid,project_id=p["project_id"])

    @bp.post("/api/portraits/employee/upload")
    def upload():
        p=project(request.form); e=employee(request.form,p); r=registry()
        if not e["active"]:
            raise ValueError("Nhân viên đã được lưu trữ")
        reviewer=str(request.form.get("reviewer") or "").strip()
        if not reviewer:
            raise ValueError("Nhập người xác nhận ảnh chân dung")
        files=request.files.getlist("files") or request.files.getlist("photos")
        if not files:
            raise ValueError("Chọn ảnh chân dung")
        prepared=[]
        for f in files:
            name=safe_component(f.filename)
            if Path(name).suffix.lower() not in _EXT:
                raise ValueError("Định dạng ảnh không hỗ trợ")
            raw=f.read(20*1024*1024+1)
            if len(raw)>20*1024*1024:
                raise ValueError("Ảnh tối đa 20MB")
            try:
                with Image.open(io.BytesIO(raw)) as img:
                    if img.width*img.height>40000000:
                        raise ValueError("Ảnh quá lớn")
                    img.verify()
            except OSError as exc:
                raise ValueError("Ảnh không hợp lệ") from exc
            prepared.append((name,raw))
        memberships=next(row["memberships"] for row in r.list_employees(p["project_id"]) if row["employee_id"]==e["employee_id"])
        if not memberships:
            raise ValueError("Xác nhận mã chấm công và ngày hiệu lực trước khi thêm ảnh")
        folder=r.project_portrait_dir(p["project_id"])/e["employee_id"]
        folder.mkdir(parents=True,exist_ok=True)
        for name,raw in prepared:
            path=folder/(uuid.uuid4().hex[:12]+"_"+name)
            with path.open("xb") as f:
                f.write(raw)
        r.bind_portrait(p["project_id"],e["employee_id"],e["employee_id"],reviewer)
        return jsonify(success=True,saved_count=len(prepared))

    @bp.get("/api/photos/view/portrait")
    def view():
        p=project(request.args); e=employee(request.args,p)
        name=request.args.get("filename","")
        path=registry()._bound_path(p["project_id"],name)
        if path not in display_paths(p,e["employee_id"]):
            return jsonify(error="Không tìm thấy ảnh của nhân viên"),404
        return send_file(path)

    @bp.post("/api/portraits/employee/delete-photo")
    def retire_photo():
        data=payload(); p=project(data); e=employee(data,p); r=registry()
        path=r._bound_path(p["project_id"],data.get("filename",""))
        if path not in display_paths(p,e["employee_id"]):
            raise ValueError("Không tìm thấy ảnh của nhân viên")
        relative=path.relative_to(r.project_portrait_dir(p["project_id"])).as_posix()
        with r._connect() as c:
            c.execute("INSERT OR REPLACE INTO portrait_exclusions VALUES(?,?,?,?)",
                      (p["project_id"],e["employee_id"],relative,date.today().isoformat()))
        return jsonify(success=True,message="Đã ngừng dùng ảnh cho nhận diện mới; giữ ảnh để đối chiếu lịch sử")

    @bp.post("/api/portraits/employee/delete")
    def archive_employee():
        data=payload(); p=project(data); e=employee(data,p)
        registry().archive_employee(e["employee_id"])
        return jsonify(success=True,message="Đã lưu trữ nhân viên; giữ nguyên ảnh và lịch sử")

    @bp.post("/api/portraits/employee/transfer")
    def transfer():
        data=payload(); r=registry()
        source=r.get_project(data.get("source_project_id") or data.get("source_project"))
        target=r.get_project(data.get("target_project_id") or data.get("target_project"))
        e=employee(data,source)
        r.transfer_employee(source["project_id"],target["project_id"],e["employee_id"],
                            data.get("effective_date"),data.get("payroll_code"),data.get("reviewer"))
        return jsonify(success=True,employee_id=e["employee_id"],message="Đã chuyển dự án, giữ lịch sử và ảnh gốc")

    app.register_blueprint(bp)
