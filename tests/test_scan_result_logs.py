from unittest.mock import patch

import pytest

from src.attendance_matching import match_attendance_record
from src.scan_runtime import ScanCancelled, scan_context


@pytest.mark.parametrize("status", ["matched", "no_match", "no_face", "no_portrait", "error", "not_run"])
def test_result_activity_contains_filter_status(status):
    record = {"date": "03/10/2026"}
    events = []
    def match(record, person, **kwargs):
        record.update(face_match_status=status, face_distance=0.25)
    with patch("src.attendance_matching._match_attendance_record", match):
        with scan_context(events.append):
            match_attendance_record(record, {"id": "001"})
    assert len(events) == 1
    assert events[0]["result_status"] == status
    assert events[0]["day"] == "03/10/2026"
    assert events[0]["distance"] == 0.25


def test_cancelled_match_does_not_emit_result():
    events = []
    with patch("src.attendance_matching._match_attendance_record", side_effect=ScanCancelled()):
        with scan_context(events.append):
            with pytest.raises(ScanCancelled):
                match_attendance_record({}, {})
    assert events == []
