# -*- coding: utf-8 -*-
"""
Module tách Excel chấm công dạng "CHI TIẾT CHẤM CÔNG"
Mỗi nhân viên -> 1 file Excel riêng, giữ nguyên header và cột.
"""
import os
import hashlib
import re
import unicodedata
from datetime import date, datetime, time, timedelta
from typing import Dict, List, Optional, Tuple

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from src.attendance_highlight import ATTENDANCE_WARNING_FILL, should_highlight_attendance


def _normalize_text(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    text = re.sub(r'\s+', ' ', text.lower().strip())
    return text


def attendance_column_map(row):
    """Use the same supported attendance headers for splitting and Word export."""
    columns = {}
    for index, value in enumerate(row):
        norm = _normalize_text(value)
        if not norm:
            continue
        if 'ma nhan vien' in norm or 'ma the' in norm or norm == 'id':
            columns['id'] = index
        elif 'ten nhan vien' in norm or 'ho va ten' in norm or norm == 'ten':
            columns['name'] = index
        elif norm == 'phong ban':
            columns['dept'] = index
        elif norm == 'ngay' or 'ngay' in norm:
            columns['date'] = index
        elif norm == 'thu':
            columns['weekday'] = index
        elif 'gio vao' in norm or 'check in' in norm or 'time in' in norm:
            columns['gio_vao'] = index
        elif 'gio ra' in norm or 'check out' in norm or 'time out' in norm:
            columns['gio_ra'] = index
    return columns


def convert_xls_cell(workbook, cell):
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ''
    if cell.ctype == xlrd.XL_CELL_DATE:
        # Elapsed formats such as [h]:mm represent durations, even above 24h.
        xf = workbook.xf_list[cell.xf_index]
        fmt = workbook.format_map[xf.format_key].format_str
        fmt = re.sub(r'"[^\"]*"|\\.', '', fmt)
        if re.search(r'\[(?:h+|m+|s+)\]', fmt, flags=re.IGNORECASE):
            return timedelta(seconds=round(float(cell.value) * 86400))
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(
            cell.value, workbook.datemode)
        if year == 0 and month == 0 and day == 0:
            return time(hour, minute, second)
        if hour or minute or second:
            return datetime(year, month, day, hour, minute, second)
        return date(year, month, day)
    return cell.value


class _ExcelReader:
    def __init__(self, path: str):
        self.path = path
        self.ext = os.path.splitext(path)[1].lower()
        self.rows: List[List] = []
        self._load()

    def _load(self):
        if self.ext == '.xlsx':
            wb = load_workbook(self.path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                self.rows.append([v if v is not None else '' for v in row])
        else:
            wb = xlrd.open_workbook(self.path, formatting_info=True)
            sheet = wb.sheet_by_index(0)
            for r in range(sheet.nrows):
                row = [self._convert_xls_cell(wb, sheet.cell(r, c)) for c in range(sheet.ncols)]
                self.rows.append(row)

    def _convert_xls_cell(self, workbook, cell):
        return convert_xls_cell(workbook, cell)


class ExcelAttendanceSplitter:
    """
    Tách Excel chấm công dạng bảng list có header:
    STT | Mã nhân viên | Tên nhân viên | Phòng ban | Ngày | Thứ | Giờ vào | Giờ ra | ...
    """

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.reader = _ExcelReader(excel_path)
        self.header_row_idx = None
        self.title_row_idx = None
        self.col_map = {}
        self._detect_header()

    def _detect_header(self):
        rows = self.reader.rows
        best_row = None
        best_score = -1
        best_map = None

        for r in range(min(40, len(rows))):
            row = rows[r]
            col_map = attendance_column_map(row)
            score = 0

            for k in ('id', 'name', 'date'):
                if k in col_map:
                    score += 1

            if score > best_score:
                best_score = score
                best_row = r
                best_map = col_map

        if best_row is None or best_score < 2:
            return

        # Detect title row above (contains "CHI TIẾT CHẤM CÔNG")
        title_row = None
        for r in range(max(0, best_row - 3), best_row + 1):
            row_text = ' '.join(_normalize_text(v) for v in rows[r] if v)
            if 'chi tiet cham cong' in row_text:
                title_row = r
                break

        self.header_row_idx = best_row
        self.title_row_idx = title_row
        self.col_map = best_map

    def _iter_data_rows(self) -> List[List]:
        rows = self.reader.rows
        if self.header_row_idx is None:
            return []
        return rows[self.header_row_idx + 1 :]

    def _row_has_data(self, row: List) -> bool:
        if not row:
            return False
        for v in row:
            if str(v).strip() != '':
                return True
        return False

    def _get_cell(self, row: List, key: str) -> str:
        idx = self.col_map.get(key)
        if idx is None or idx >= len(row):
            return ''
        return str(row[idx]).strip()

    def split(self, output_dir: str) -> Tuple[List[str], List[Dict]]:
        if self.header_row_idx is None:
            raise ValueError("Không nhận diện được header của bảng chấm công.")

        output_dir = os.path.normpath(output_dir)
        drive, rest = os.path.splitdrive(output_dir)
        parts = [p.rstrip('. ') for p in rest.split(os.sep) if p]
        clean_dir = (drive + os.sep + os.sep.join(parts)) if drive else (os.sep.join(parts) or output_dir)
        os.makedirs(clean_dir, exist_ok=True)
        output_dir = clean_dir

        rows = self._iter_data_rows()
        groups: Dict[Tuple[str, str], List[List]] = {}

        for row_index, row in enumerate(rows):
            if not self._row_has_data(row):
                continue
            name = self._get_cell(row, 'name')
            emp_id = self._get_cell(row, 'id')
            if not name and not emp_id:
                continue
            # Unknown codes remain separate source rows; never infer identity from names.
            key = (str(emp_id).strip() if emp_id else f"pending:{row_index}", str(emp_id).strip())
            groups.setdefault(key, []).append(row)

        selected = {key: (emp_id, data_rows, None)
                    for (key, emp_id), data_rows in groups.items()}

        header_rows = []
        if self.title_row_idx is not None and self.title_row_idx < self.header_row_idx:
            for r in range(self.title_row_idx, self.header_row_idx + 1):
                header_rows.append(self.reader.rows[r])
        else:
            header_rows.append(self.reader.rows[self.header_row_idx])

        output_files = []
        summaries = []
        for norm_name, (emp_id, data_rows, score) in selected.items():
            display_name = next((self._get_cell(row, 'name') for row in data_rows
                                 if self._get_cell(row, 'name')), emp_id)
            # Fill missing display names only within the same explicit source code.
            if 'name' in self.col_map:
                named_rows = []
                for row in data_rows:
                    values = list(row)
                    if not self._get_cell(row, 'name'):
                        values[self.col_map['name']] = display_name
                    named_rows.append(values)
                data_rows = named_rows
            safe_name = re.sub(r'[<>:"/\\\\|?*\x00-\x1f]', '_', str(display_name).strip()).rstrip('. ')
            if not safe_name:
                safe_name = "nhan_vien"
            identity_suffix = hashlib.sha256((emp_id or norm_name).encode("utf-8")).hexdigest()[:16]
            filename = f"{safe_name}_{identity_suffix}.xlsx"
            out_path = os.path.join(output_dir, filename)

            wb = Workbook()
            ws = wb.active
            # Write header rows
            for hr in header_rows:
                ws.append(list(hr))
            # Write data rows for this person
            for r in data_rows:
                ws.append(list(r))
                if should_highlight_attendance(r, self.col_map):
                    for cell in ws[ws.max_row]:
                        cell.fill = PatternFill('solid', fgColor=ATTENDANCE_WARNING_FILL)
            wb.save(out_path)

            output_files.append(out_path)
            summaries.append({
                'name': str(display_name),
                'id': str(emp_id),
                'rows': len(data_rows),
                'present_rows': sum(bool(self._get_cell(row, 'gio_vao') or self._get_cell(row, 'gio_ra')) for row in data_rows),
            })

        return output_files, summaries

