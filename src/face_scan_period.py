"""Explicit calendar-month limits shared by Excel and PDF face scans."""
import calendar
import re
from datetime import date
from src.aggregate_report_exporter import _parse_date


def month_limits(value):
    if not isinstance(value,str) or not re.fullmatch(r'\d{4}-\d{2}',value):
        raise ValueError('Tháng quét phải có định dạng YYYY-MM')
    year,month=map(int,value.split('-'))
    try:
        start=date(year,month,1)
        end=date(year,month,calendar.monthrange(year,month)[1])
    except (ValueError,calendar.IllegalMonthError) as exc:
        raise ValueError('Tháng quét không hợp lệ') from exc
    return start.isoformat(),end.isoformat()


def filter_scan_people(persons, options):
    if not options or not options.get('from_date'):
        return persons
    start,end=_parse_date(options['from_date']),_parse_date(options['to_date'])
    result=[]
    for person in persons:
        records=[r for r in person.get('records',[]) if (day:=_parse_date(r.get('date'))) is not None and start<=day<=end]
        if records:
            result.append({**person,'records':records})
    if not result:
        raise ValueError(f'Bảng đã tách không có dữ liệu chấm công từ {options["from_date"]} đến {options["to_date"]}. Chọn tháng phù hợp hoặc tải bảng tháng cần quét.')
    return result
