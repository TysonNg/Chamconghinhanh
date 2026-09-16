from pathlib import Path

import src.app as app_module


def test_open_excel_output_folder_opens_requested_batch(tmp_path, monkeypatch):
    output_root = tmp_path / "excel_extracted"
    opened = []
    monkeypatch.setattr(app_module, "EXCEL_OUTPUT_DIR", str(output_root))
    monkeypatch.setattr(app_module.os, "startfile", lambda path: opened.append(Path(path)))

    response = app_module.app.test_client().post(
        "/api/open/folder",
        json={"type": "excel_output", "subpath": "thang-08-2026"},
    )

    assert response.status_code == 200
    assert response.get_json()["success"] is True
    assert opened == [output_root / "thang-08-2026"]
    assert opened[0].is_dir()


def test_open_excel_output_folder_rejects_path_traversal(tmp_path, monkeypatch):
    output_root = tmp_path / "excel_extracted"
    opened = []
    monkeypatch.setattr(app_module, "EXCEL_OUTPUT_DIR", str(output_root))
    monkeypatch.setattr(app_module.os, "startfile", lambda path: opened.append(Path(path)))

    response = app_module.app.test_client().post(
        "/api/open/folder",
        json={"type": "excel_output", "subpath": "..\\private"},
    )

    assert response.status_code == 400
    assert response.get_json()["success"] is False
    assert opened == []

