# -*- coding: utf-8 -*-
import pytest
from src.excel_extractor import (
    get_work_duration_minutes,
    is_under_work_duration,
    ExcelChamCongExtractor,
)


def test_work_duration_minutes_normal():
    # 2 hours
    assert get_work_duration_minutes('07:30', '09:30') == 120
    # Exactly 3 hours
    assert get_work_duration_minutes('07:00', '10:00') == 180
    # 3 hours 1 minute
    assert get_work_duration_minutes('07:00', '10:01') == 181
    # 8 hours
    assert get_work_duration_minutes('08:00', '17:00') == 540


def test_work_duration_minutes_overnight():
    # 22:00 to 00:30 (2.5 hours = 150 minutes)
    assert get_work_duration_minutes('22:00', '00:30') == 150
    # 23:00 to 02:00 (3 hours = 180 minutes)
    assert get_work_duration_minutes('23:00', '02:00') == 180
    # 22:00 to 06:00 (8 hours = 480 minutes)
    assert get_work_duration_minutes('22:00', '06:00') == 480


def test_work_duration_invalid():
    assert get_work_duration_minutes('', '09:30') is None
    assert get_work_duration_minutes('07:30', '') is None
    assert get_work_duration_minutes(None, '09:30') is None
    assert get_work_duration_minutes('abc', 'def') is None


def test_is_under_work_duration():
    # <= 3 hours
    assert is_under_work_duration('07:00', '09:30', max_hours=3.0) is True
    assert is_under_work_duration('07:00', '10:00', max_hours=3.0) is True
    assert is_under_work_duration('22:00', '00:30', max_hours=3.0) is True

    # > 3 hours
    assert is_under_work_duration('07:00', '10:01', max_hours=3.0) is False
    assert is_under_work_duration('08:00', '17:00', max_hours=3.0) is False
    assert is_under_work_duration('22:00', '06:00', max_hours=3.0) is False


def test_get_absent_records_under_3h():
    class DummyExtractor:
        get_absent_records = ExcelChamCongExtractor.get_absent_records

    extractor = DummyExtractor()
    person = {
        'name': 'Test Nhân Viên',
        'records': [
            {
                'date': '01/03/2026',
                'gio_vao': '07:00',
                'gio_ra': '17:00',
                'is_absent': False,
                'same_in_out': False,
                'under_3h': False,
                'missing_checkin': False,
                'missing_checkout': False
            },
            {
                'date': '02/03/2026',
                'gio_vao': '07:00',
                'gio_ra': '09:30',
                'is_absent': False,
                'same_in_out': False,
                'under_3h': True,
                'missing_checkin': False,
                'missing_checkout': False
            },
            {
                'date': '03/03/2026',
                'gio_vao': '07:00',
                'gio_ra': '07:02',
                'is_absent': False,
                'same_in_out': True,
                'under_3h': False,
                'missing_checkin': False,
                'missing_checkout': False
            },
            {
                'date': '04/03/2026',
                'gio_vao': '',
                'gio_ra': '',
                'is_absent': True,
                'same_in_out': False,
                'under_3h': False,
                'missing_checkin': False,
                'missing_checkout': False
            }
        ]
    }

    absent = extractor.get_absent_records(person)
    assert len(absent) == 3

    # Check issues
    assert absent[0]['issue'] == 'Làm <3 tiếng (07:00 - 09:30)'
    assert absent[1]['issue'] == 'Trùng giờ vào/ra (07:00 - 07:02)'
    assert absent[2]['issue'] == 'Vắng mặt'
