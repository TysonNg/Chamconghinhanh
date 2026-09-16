from concurrent.futures import ThreadPoolExecutor
import sqlite3
import pytest
from src.identity_registry import IdentityRegistry

def test_batch_codes_are_unique_stable_and_leave_payroll_untouched(tmp_path):
    r=IdentityRegistry(tmp_path/"ids.db",tmp_path/"portraits")
    p=r.register_project("A")
    people=[r.create_employee("Same") for _ in range(30)]
    r.assign_employee(p["project_id"],people[0]["employee_id"],"001","2026-01-01")
    result=r.generate_internal_codes()
    assert result["created_count"]==30
    with r._connect() as c:
        rows=c.execute("SELECT employee_id,code FROM employee_internal_codes").fetchall()
    assert len({row["code"] for row in rows})==30
    assert all(row["code"].startswith("NV-") for row in rows)
    assert r.resolve_employee(p["project_id"],"001","2026-09-01").employee_id==people[0]["employee_id"]
    assert r.generate_internal_codes()["created_count"]==0
    reopened=IdentityRegistry(r.db_path,r.portrait_root)
    assert reopened.get_employee(people[0]["employee_id"])["internal_code"]==next(row["code"] for row in rows if row["employee_id"]==people[0]["employee_id"])

def test_collision_retries_and_parallel_requests_do_not_replace_codes(tmp_path,monkeypatch):
    r=IdentityRegistry(tmp_path/"ids.db",tmp_path/"portraits")
    people=[r.create_employee("Same") for _ in range(2)]
    values=iter(["aaaaaaaa","aaaaaaaa","bbbbbbbb"])
    monkeypatch.setattr("src.identity_registry.secrets.token_hex",lambda n:next(values))
    with ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(lambda _:r.generate_internal_codes(),range(2)))
    assert sum(row["created_count"] for row in results)==2
    assert {r.get_employee(e["employee_id"])["internal_code"] for e in people}=={"NV-AAAAAAAA","NV-BBBBBBBB"}

def test_endpoint_imports_pending_sources_without_confirming_identity(tmp_path):
    from flask import Flask
    from src.identity_routes import register_identity_routes
    r=IdentityRegistry(tmp_path/"ids.db",tmp_path/"portraits")
    for project in ("A","B"):
        (r.portrait_root/project/"Same Name").mkdir(parents=True)
    app=Flask(__name__);app.testing=True
    register_identity_routes(app,lambda:r,lambda:tmp_path/"camera")
    client=app.test_client()
    response=client.post("/api/portraits/internal-codes/generate",json={})
    assert response.status_code==200
    assert response.json["created_count"]==2
    codes=[]
    for p in r.list_projects():
        emp=client.get("/api/portraits",query_string={"project_id":p["project_id"]}).json["employees"][0]
        codes.append(emp["internal_code"])
        assert emp["payroll_code"]==""
        assert emp["identity_status"]=="needs_confirmation"
    assert len(set(codes))==2

def test_export_does_not_render_employee_codes(tmp_path):
    from docx import Document
    from src.excel_extractor import ExcelToWordExporter
    exporter=ExcelToWordExporter(str(tmp_path/"portraits"),str(tmp_path/"out"))
    path=exporter.export_person({"name":"Employee","id":"PAYROLL-SECRET","internal_code":"NV-SECRET",
                                "month":9,"year":2026,"records":[]})
    doc=Document(path)
    text="\n".join(p.text for p in doc.paragraphs)
    assert "Employee".upper() in text
    assert "PAYROLL-SECRET" not in text
    assert "NV-SECRET" not in text
