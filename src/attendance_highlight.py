"""Shared highlighting rules for individual attendance exports."""
import re
from datetime import date, datetime, time

ATTENDANCE_WARNING_FILL = 'FFF2CC'


def _seconds(value, last=False):
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(value * 86400) if 0 <= value < 1 else None
    matches = list(re.finditer(r'(?<!\d)(\d{1,2}):(\d{2})(?::(\d{2}))?(?!\d)', str(value)))
    if not matches:
        return None
    match = matches[-1] if last else matches[0]
    hour, minute, second = int(match[1]), int(match[2]), int(match[3] or 0)
    if hour > 23 or minute > 59 or second > 59:
        return None
    return hour * 3600 + minute * 60 + second


def should_highlight_attendance(row, columns):
    """Highlight dated rows with missing punches or a duration of at most 3h."""
    if not all(key in columns for key in ('date', 'gio_vao', 'gio_ra')):
        return False

    def cell(key):
        index = columns[key]
        return row[index] if index < len(row) else None

    day = cell('date')
    if not isinstance(day, date):
        text = str(day or '').strip()
        # Exclude totals and repeated headers; accept the supported date strings.
        if not re.fullmatch(r'\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}(?:\s+00:00:00)?', text):
            return False
    start, end = cell('gio_vao'), cell('gio_ra')
    if any(value is None or str(value).strip() == '' for value in (start, end)):
        return True
    start_seconds, end_seconds = _seconds(start), _seconds(end, last=True)
    if start_seconds is None or end_seconds is None:
        return False
    return (end_seconds - start_seconds) % 86400 <= 3 * 3600
