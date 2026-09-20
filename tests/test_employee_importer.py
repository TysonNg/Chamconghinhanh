# -*- coding: utf-8 -*-
import os
import shutil
import tempfile
from pathlib import Path
import pytest
from flask import Flask

from src.employee_importer import (
    clean_person_name,
    clean_payroll_code,
    normalize_vietnamese_name,
    extract_employees_from_excel,
    extract_employees_from_pdf,
    extract_employees_from_file,
    sync_employees_to_project,
)
from src.identity_registry import IdentityRegistry
from src.identity_routes import register_identity_routes


def test_name_cleaning_and_normalization():
    # Test cleaning Vietnamese security guard & employee titles
    raw_names = [
        ("BAO VE TTD Bui Tien Giap", "Bui Tien Giap"),
        ("BVNX A7 - LY VAN MANH", "LY VAN MANH"),
        ("BVNX An Vien - Nguyen Van Banh", "Nguyen Van Banh"),
        ("BAO VE SANH C3 - DANG NGUYEN HOANG KHA", "DANG NGUYEN HOANG KHA"),
        ("BAO VE TTD Pham Tai Ngan C1", "Pham Tai Ngan"),
        ("Do Tri Thuc BV151", "Do Tri Thuc"),
        ("Bui Quoc Hien BVNXST", "Bui Quoc Hien"),
        ("Bao ve TTD - Nguyen Viet Mai - 30.7.24", "Nguyen Viet Mai"),
        ("Nguyen Thanh Chan BVSAH4", "Nguyen Thanh Chan"),
        ("BAO VE TTD - Lê Văn Tòng BV", "Lê Văn Tòng"),
    ]
    for raw, expected in raw_names:
        assert clean_person_name(raw) == expected

    # Test smart Vietnamese normalization matching
    assert normalize_vietnamese_name("Nguyễn Văn A") == normalize_vietnamese_name("NGUYEN VAN A")
    assert normalize_vietnamese_name("Bùi Tiến Giáp") == normalize_vietnamese_name("BAO VE TTD Bui Tien Giap")
    assert normalize_vietnamese_name("Lý Văn Mạnh") == normalize_vietnamese_name("BVNX A7 - LY VAN MANH")
    assert normalize_vietnamese_name("Nguyễn Văn Bảnh") == normalize_vietnamese_name("BVNX An Vien - Nguyen Van Banh")


def test_extract_from_real_excel_file():
    excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "excel_uploads", "t8.2026   ehsg.xlsx")
    if os.path.exists(excel_path):
        employees = extract_employees_from_excel(excel_path)
        assert len(employees) >= 80
        # Verify first few records have valid names and codes
        for emp in employees:
            assert len(emp["name"]) >= 2
            assert emp["payroll_code"] != ""


def test_extract_from_real_pdf_file():
    pdf_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "pdf_uploads", "CHAM_CONG_BAO_VE_T2.2026.pdf")
    if os.path.exists(pdf_path):
        employees = extract_employees_from_pdf(pdf_path)
        assert len(employees) >= 60
        # Verify codes are preserved with leading zeros
        codes = [e["payroll_code"] for e in employees]
        assert any(c.startswith("00") for c in codes)


def test_sync_employees_to_project_with_mock_portraits(tmp_path):
    db_path = tmp_path / "test_identity.sqlite3"
    portrait_root = tmp_path / "portraits"
    registry = IdentityRegistry(db_path, portrait_root)

    # Register project
    proj = registry.register_project("Dự án Test")
    pid = proj["project_id"]

    # Create dummy portrait folder for "Bùi Tiến Giáp"
    proj_dir = registry.project_portrait_dir(pid)
    person_dir = proj_dir / "Bùi Tiến Giáp"
    person_dir.mkdir(parents=True)
    dummy_img = person_dir / "face.jpg"
    dummy_img.write_bytes(b"\xff\xd8\xff\xe0" + b"\x00" * 100)

    # Mock list of employees extracted from Excel/PDF
    extracted = [
        {"name": "BAO VE TTD Bui Tien Giap", "payroll_code": "00328"},
        {"name": "Trần Văn Mới", "payroll_code": "00999"},
    ]

    results = sync_employees_to_project(pid, extracted, registry)

    assert results["total_found"] == 2
    assert results["bound_existing"] == 1  # Bui Tien Giap matched existing folder
    assert results["created_new"] == 1     # Tran Van Moi is newly created

    # Verify employees in registry
    emp_list = registry.list_employees(pid)
    assert len(emp_list) == 2

    giap = next(e for e in emp_list if "Giáp" in e["display_name"] or "Giap" in e["display_name"])
    assert giap["memberships"][0]["payroll_code"] == "00328"

    moi = next(e for e in emp_list if "Mới" in e["display_name"])
    assert moi["memberships"][0]["payroll_code"] == "00999"


def test_api_import_file_endpoint(tmp_path):
    app = Flask(__name__)
    db_path = tmp_path / "identity.sqlite3"
    portrait_root = tmp_path / "portraits"
    input_root = tmp_path / "input_images"
    portrait_root.mkdir(parents=True)
    input_root.mkdir(parents=True)

    registry = IdentityRegistry(db_path, portrait_root)
    proj = registry.register_project("Dự án Alpha")
    pid = proj["project_id"]

    register_identity_routes(app, lambda: registry, lambda: str(input_root))

    client = app.test_client()

    # Create dummy Excel file
    excel_file = tmp_path / "test_list.xlsx"
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["STT", "Mã nhân viên", "Tên nhân viên", "Phòng ban"])
    ws.append([1, "00111", "Nguyễn Văn Một", "Bảo vệ"])
    ws.append([2, "00222", "Trần Văn Hai", "Bảo vệ"])
    wb.save(str(excel_file))

    # Test POST /api/portraits/import-file with multipart file upload
    with open(excel_file, "rb") as f:
        res = client.post(
            "/api/portraits/import-file",
            data={"file": f, "project_id": pid},
            content_type="multipart/form-data"
        )

    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["total_found"] == 2
    assert data["created_new"] == 2

    # Verify employees now exist in project
    emp_res = client.get(f"/api/portraits?project_id={pid}")
    emp_data = emp_res.get_json()
    assert emp_data["success"] is True
    assert emp_data["total"] == 2
    assert any(e["payroll_code"] == "00111" for e in emp_data["employees"])
    assert any(e["payroll_code"] == "00222" for e in emp_data["employees"])
