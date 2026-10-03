"""Project and portrait management using persistent IDs; legacy photos stay pending."""
from datetime import date
from pathlib import Path
from urllib.parse import urlencode
import io
import os
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
        from src.supplement_evidence import evidence_is_visible
        today_str = date.today().isoformat()
        for p in r.list_projects():
            r.import_legacy(p["project_id"])
            camera=Path(input_root())/p["storage_dir"]
            photos=[x for x in camera.rglob("*") if x.is_file() and x.suffix.lower() in _EXT and evidence_is_visible(x)] if camera.exists() else []
            active_emps = [
                e for e in r.list_employees(p["project_id"])
                if e["active"] and (
                    e["source_paths"] or not e["memberships"]
                    or not e["memberships"][0].get("valid_to")
                    or e["memberships"][0]["valid_to"] > today_str
                )
            ]
            items.append({**p,"name":p["storage_dir"],"employee_count":len(active_emps),
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
        r.generate_internal_codes()
        items=[]
        search=request.args.get("search","").casefold()
        include_history=request.args.get("include_history","").lower() in ("true","1")
        today_str = date.today().isoformat()
        for e in r.list_employees(p["project_id"]):
            if not e["active"] and not include_history:
                continue
            # Bỏ qua nhân viên đã chuyển khỏi dự án này (membership đã kết thúc và không còn trong legacy_sources)
            if not include_history and e["memberships"] and not e["source_paths"]:
                latest = e["memberships"][0]
                if latest.get("valid_to") and latest["valid_to"] <= today_str:
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
        payroll_code = data.get("payroll_code")
        valid_from = data.get("valid_from") or "2000-01-01"
        reviewer = str(data.get("reviewer") or "system").strip()
        registry().confirm_source(p["project_id"],e["employee_id"],payroll_code,valid_from,reviewer)
        return jsonify(success=True,employee_id=e["employee_id"])

    @bp.post("/api/portraits/employee/create")
    def create_employee():
        from src.identity_registry import _day
        data=payload(); p=project(data); r=registry()
        name=str(data.get("name") or "").strip()
        reviewer=str(data.get("reviewer") or "system").strip()
        if not name:
            raise ValueError("Nhập tên nhân viên")
        payroll_code = str(data.get("payroll_code") or "").strip()
        valid_from = data.get("valid_from") or "2000-01-01"
        eid=uuid.uuid4().hex
        with r._connect() as c:
            c.execute("BEGIN IMMEDIATE")
            c.execute("INSERT INTO employees(employee_id,display_name) VALUES(?,?)",(eid,name))
            internal_code, _ = r._ensure_internal_code(c, eid)
            r._assign(c,p["project_id"],eid,payroll_code,_day(valid_from),None)
        folder=r.project_portrait_dir(p["project_id"])/eid
        folder.mkdir(parents=True,exist_ok=True)
        r.bind_portrait(p["project_id"],eid,eid,reviewer)
        return jsonify(success=True,name=name,employee_id=eid,project_id=p["project_id"],payroll_code=payroll_code,internal_code=internal_code)

    @bp.post("/api/portraits/import-file")
    def import_file():
        """Nhập danh sách nhân viên & mã từ file Excel (.xls/.xlsx) hoặc PDF."""
        from src.employee_importer import extract_employees_from_file, sync_employees_to_project
        import tempfile
        r = registry()

        # Xác định dự án đích
        form_project = request.form.get("project_id") or request.form.get("project")
        json_data = request.get_json(silent=True) or {}
        proj_val = form_project or json_data.get("project_id") or json_data.get("project") or ""
        p = registry().get_project(proj_val)
        if not p["active"]:
            raise ValueError("Dự án đã được lưu trữ")

        temp_path = None
        target_path = None

        try:
            if "file" in request.files:
                file = request.files["file"]
                if not file or not file.filename:
                    raise ValueError("Không có file được chọn")
                ext = Path(file.filename).suffix.lower()
                if ext not in (".xls", ".xlsx", ".pdf"):
                    raise ValueError("Chỉ chấp nhận file .xls, .xlsx hoặc .pdf")
                fd, temp_path = tempfile.mkstemp(suffix=ext)
                os.close(fd)
                file.save(temp_path)
                target_path = temp_path
            else:
                raw_path = json_data.get("file_path") or json_data.get("filename")
                if not raw_path:
                    raise ValueError("Thiếu file upload hoặc đường dẫn file")
                # Kiểm tra trong các thư mục uploads nếu chỉ truyền filename
                candidates = [
                    Path(raw_path),
                    Path(r.portrait_root).parent / "excel_uploads" / raw_path,
                    Path(r.portrait_root).parent / "pdf_uploads" / raw_path,
                ]
                for cand in candidates:
                    if cand.exists() and cand.is_file():
                        target_path = str(cand)
                        break
                if not target_path:
                    raise ValueError(f"Không tìm thấy file: {raw_path}")

            employees = extract_employees_from_file(target_path)
            if not employees:
                return jsonify(
                    success=True,
                    project=p["storage_dir"],
                    project_id=p["project_id"],
                    total_found=0,
                    bound_existing=0,
                    created_new=0,
                    updated_code=0,
                    unchanged=0,
                    details=[],
                    message="Không tìm thấy nhân viên nào trong file"
                )

            sync_results = sync_employees_to_project(
                project_id=p["project_id"],
                employees=employees,
                identity_registry=r
            )

            return jsonify(
                success=True,
                project=p["storage_dir"],
                project_id=p["project_id"],
                **sync_results
            )
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    @bp.post("/api/portraits/employee/upload")
    def upload():
        p=project(request.form); e=employee(request.form,p); r=registry()
        if not e["active"]:
            raise ValueError("Nhân viên đã được lưu trữ")
        reviewer=str(request.form.get("reviewer") or "system").strip()
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

    @bp.post("/api/portraits/open-folder")
    def open_portrait_folder():
        import sys
        data = payload()
        p = project(data)
        r = registry()
        root = r.project_portrait_dir(p["project_id"])
        root.mkdir(parents=True, exist_ok=True)

        eid = data.get("employee_id")
        target_folder = root
        if eid:
            e = employee(data, p)
            paths = display_paths(p, e["employee_id"])
            if paths:
                target_folder = paths[0].parent
            else:
                target_folder = root / e["employee_id"]
                target_folder.mkdir(parents=True, exist_ok=True)

        target_folder = target_folder.resolve()
        try:
            if sys.platform == "win32":
                os.startfile(str(target_folder))
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(target_folder)])
            return jsonify(success=True, path=str(target_folder))
        except Exception as exc:
            return jsonify(success=False, error=str(exc)), 500

    @bp.post("/api/portraits/employee/delete-photo")
    def retire_photo():
        data = payload(); p = project(data); e = employee(data, p); r = registry()
        root = r.project_portrait_dir(p["project_id"])
        filename = str(data.get("filename", "")).strip()
        if not filename:
            raise ValueError("Chọn ảnh cần xóa")

        emp_paths = display_paths(p, e["employee_id"])
        matched = None
        for candidate in emp_paths:
            rel = candidate.relative_to(root).as_posix()
            if (rel == filename or 
                rel.replace('\\', '/') == filename.replace('\\', '/') or
                candidate.name == filename or
                candidate.name == Path(filename).name):
                matched = candidate
                break

        if not matched:
            try:
                candidate = r._bound_path(p["project_id"], filename)
                if candidate in emp_paths or candidate.resolve() in [x.resolve() for x in emp_paths]:
                    matched = candidate
            except Exception:
                pass

        if not matched:
            raise ValueError("Không tìm thấy ảnh của nhân viên")

        relative = matched.relative_to(root).as_posix()
        with r._connect() as c:
            c.execute("INSERT OR REPLACE INTO portrait_exclusions VALUES(?,?,?,?)",
                      (p["project_id"], e["employee_id"], relative, date.today().isoformat()))

        # Xóa vĩnh viễn tệp ảnh vật lý trên ổ đĩa
        if matched.exists():
            matched.unlink(missing_ok=True)

        return jsonify(success=True, message="Đã xóa ảnh chân dung thành công", deleted=relative)

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
        payroll_code = str(data.get("payroll_code") or "").strip()
        if not payroll_code:
            with r._connect() as c:
                row = c.execute("SELECT payroll_code FROM memberships WHERE project_id=? AND employee_id=? AND payroll_code<>'' ORDER BY valid_from DESC",
                                (source["project_id"], e["employee_id"])).fetchone()
                if row:
                    payroll_code = row[0]
        r.transfer_employee(source["project_id"],target["project_id"],e["employee_id"],
                            data.get("effective_date") or date.today().isoformat(),
                            payroll_code,
                            data.get("reviewer") or "system")
        return jsonify(success=True,employee_id=e["employee_id"],message="Đã chuyển dự án, giữ lịch sử và ảnh gốc")

    app.register_blueprint(bp)
