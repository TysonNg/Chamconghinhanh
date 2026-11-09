import threading
from types import SimpleNamespace

import pytest

from src.face_inference import FaceInferenceError
from src.scan_runtime import ScanCancelled, scan_context, check_scan_cancelled, report_scan_activity


def test_scan_context_cancellation_does_not_leak_to_next_person():
    with pytest.raises(ScanCancelled):
        with scan_context(cancel_check=lambda: True):
            pass
    check_scan_cancelled()


def test_excel_analyzer_fails_whole_task_on_native_timeout(tmp_path, monkeypatch):
    import src.excel_face_analyzer as module
    source = tmp_path / 'source'
    source.mkdir()
    (source / 'one.xlsx').touch()
    person = {'id': '001', 'name': 'Person at timeout', 'month': 10, 'year': 2026,
              'records': [{'date': '03/10/2026', 'is_absent': True}]}
    monkeypatch.setattr(module, 'ExcelPersonFileParser', lambda _: SimpleNamespace(parse_person=lambda: person))
    activity = []

    class Exporter:
        def __init__(self, *args, **kwargs):
            pass

        def export_person(self, person, **kwargs):
            report_scan_activity('Đang đối chiếu ảnh', image='002.jpg')
            raise FaceInferenceError('native timeout')

    monkeypatch.setattr(module, 'ExcelToWordExporter', Exporter)
    analyzer = module.ExcelFaceAnalyzer(str(tmp_path / 'portraits'), str(tmp_path / 'photos'), None)
    completed = []
    with pytest.raises(FaceInferenceError, match='Person at timeout'):
        analyzer.analyze_folder(str(source), str(tmp_path / 'output'),
            progress_callback=lambda *args: completed.append(args),
            activity_callback=lambda worker, detail: activity.append(detail))
    assert completed == [(0, 1, '', None)]  # Publish the total before inference; timeout completes nobody.
    assert any(detail and detail.get('image') == '002.jpg' for detail in activity)
    assert activity[-1] is None
