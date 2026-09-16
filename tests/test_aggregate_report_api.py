import pytest

import src.app as app_module
from src.app import ExcelFaceTask, PDFFaceTask, _build_report_options


def test_build_report_options_accepts_iso_range_and_report_name():
    assert _build_report_options(
        {
            "report_project_name": "Chung cư Tân Thuận Đông",
            "from_date": "2026-07-01",
            "to_date": "2026-07-31",
        },
        "Dự án mặc định",
    ) == {
        "project_name": "Chung cư Tân Thuận Đông",
        "from_date": "2026-07-01",
        "to_date": "2026-07-31",
    }


def test_build_report_options_rejects_reversed_range():
    with pytest.raises(ValueError, match="Từ ngày"):
        _build_report_options(
            {"from_date": "2026-07-31", "to_date": "2026-07-01"},
            "Dự án",
        )


@pytest.mark.parametrize("task_class", [ExcelFaceTask, PDFFaceTask])
def test_face_task_status_exposes_aggregate_report(task_class):
    task = task_class("task-1")
    task.aggregate_report = {"name": "GIAI_TRINH.docx", "folder": "dot-1"}
    assert task.to_dict()["aggregate_report"] == {
        "name": "GIAI_TRINH.docx",
        "folder": "dot-1",
    }


def test_aggregate_reports_api_lists_excel_and_pdf_reports(tmp_path, monkeypatch):
    excel_dir = tmp_path / "excel"
    pdf_dir = tmp_path / "pdf"
    (excel_dir / "dot-thang-7").mkdir(parents=True)
    (pdf_dir / "bao-cao-pdf").mkdir(parents=True)

    excel_report = excel_dir / "dot-thang-7" / "GIAI_TRINH_DU_AN_A_01-07-2026_31-07-2026.docx"
    pdf_report = pdf_dir / "bao-cao-pdf" / "GIAI_TRINH_DU_AN_B_01-07-2026_31-07-2026.docx"
    excel_report.write_bytes(b"excel-report")
    pdf_report.write_bytes(b"pdf-report")
    (excel_dir / "dot-thang-7" / "NGUYEN_VAN_A.docx").write_bytes(b"personal-report")

    monkeypatch.setattr(app_module, "EXCEL_FACE_OUTPUT_DIR", str(excel_dir))
    monkeypatch.setattr(app_module, "PDF_FACE_OUTPUT_DIR", str(pdf_dir))

    response = app_module.app.test_client().get("/api/aggregate-reports")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["count"] == 2
    assert {(item["source"], item["folder"], item["name"]) for item in payload["files"]} == {
        ("Excel", "dot-thang-7", excel_report.name),
        ("PDF", "bao-cao-pdf", pdf_report.name),
    }
    assert all(item["download_url"].startswith("/api/") for item in payload["files"])
