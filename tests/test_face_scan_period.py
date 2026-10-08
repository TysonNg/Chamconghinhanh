import pytest
from src.face_scan_period import month_limits, filter_scan_people


def test_month_limits_and_leap_year():
    assert month_limits('2024-02')==('2024-02-01','2024-02-29')
    assert month_limits('2026-06')==('2026-06-01','2026-06-30')
    with pytest.raises(ValueError):month_limits('2026-13')


def test_explicit_month_filters_records_before_export_and_preserves_source():
    persons=[{'name':'A','records':[{'date':'31/05/2026'},{'date':'01/06/2026'},{'date':'30/06/2026'},{'date':'01/07/2026'}]},
             {'name':'B','records':[{'date':'01/07/2026'}]}]
    output=filter_scan_people(persons,{'from_date':'2026-06-01','to_date':'2026-06-30'})
    assert len(output)==1 and len(output[0]['records'])==2
    assert len(persons[0]['records'])==4
    assert filter_scan_people(persons,{'project_name':'A'}) is persons
    with pytest.raises(ValueError,match='không có dữ liệu'):
        filter_scan_people(persons,{'from_date':'2026-08-01','to_date':'2026-08-31'})


def test_api_report_options_support_selected_month():
    from src.app import _build_report_options
    assert _build_report_options({'scan_month':'2026-06'},'A')=={
        'project_name':'A','from_date':'2026-06-01','to_date':'2026-06-30'}
