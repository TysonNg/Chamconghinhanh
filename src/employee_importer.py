# -*- coding: utf-8 -*-
"""
Module trích xuất danh sách nhân viên (Tên + Mã chấm công) từ Excel/PDF
và đồng bộ tự động vào cơ sở dữ liệu dự án / ảnh chân dung.
"""

import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import openpyxl
import xlrd

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


def clean_person_name(name: str) -> str:
    """Làm sạch tên người, loại bỏ tiền tố chức vụ/dự án/vị trí và hậu tố ngày tháng, mã ca."""
    if not name:
        return ""
    s = str(name).strip()
    # Loại bỏ ngày tháng ở đuôi: - 30.7.24 hoặc - 03.08.2024
    s = re.sub(r"\s*[-–]\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}.*$", "", s, flags=re.I)
    # Loại bỏ tiền tố chức vụ / dự án / vị trí / tổ đội
    prefix_pat = r"^(?:b[aả]o\s*v[eệ]|bvnx|bv|nv\s*v[eệ]\s*sinh)\s*(?:ttd|s[aả]nh\s*[a-z0-9]*|nx\s*[a-z0-9]*|nxav|an\s*kh[aá]ng|an\s*vi[eê]n)?\s*[-–:]*\s*"
    s = re.sub(prefix_pat, "", s, flags=re.I)
    # Loại bỏ tiền tố chung cư / sảnh / cổng nếu có
    # Consume location tokens only; never swallow the following person's name.
    s = re.sub(r"^(?:nx[a-z0-9]*|s[aả]nh(?:\s+[a-z]\d+)?|a\d+|c\d+)\s+(?:[-–:]\s*)?", "", s, flags=re.I)
    s = re.sub(r"^(?:cc|chung\s*c[uư])\s+[^-–:]+\s*[-–:]\s*", "", s, flags=re.I)
    # Loại bỏ các hậu tố vị trí / mã ca ở cuối như BVSAH4, BVNXST, BV151, C1, A3, NXAV, BV, vv.
    suffix_pat = r"\s+(?:bvsah\d*|bvnxst|bvnxav|bvsak|bv151|bvnx|bv\s*[-–]?\s*[a-z0-9]+|bv|nxav|av|c\d|a\d)\s*$"
    s = re.sub(suffix_pat, "", s, flags=re.I)
    # Loại bỏ dấu gạch nối hoặc ký tự thừa đầu/cuối
    s = s.strip(" -–:")
    s = re.sub(r'[<>:"/\\|?*\r\n\t]', " ", s)
    return re.sub(r"\s+", " ", s).strip()


def normalize_vietnamese_name(name: str) -> str:
    """
    Chuẩn hóa tên tiếng Việt để so khớp:
    - Loại bỏ tiền tố/hậu tố chức vụ, dự án qua clean_person_name
    - Chuyển chữ thường
    - Bỏ dấu tiếng Việt (NFD)
    - Thay thế đ/Đ -> d
    - Bỏ ký tự đặc biệt, chỉ giữ chữ cái và khoảng trắng
    - Chuẩn hóa khoảng trắng
    """
    if not name:
        return ""
    cleaned = clean_person_name(name)
    s = cleaned.lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    # Giữ lại chữ và số, khoảng trắng
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_display_name(name: str) -> str:
    """Làm sạch tên hiển thị (bỏ tiền/hậu tố rác, bỏ ký tự cấm của Windows, chuẩn hóa khoảng trắng)."""
    cleaned = clean_person_name(name)
    return cleaned or str(name).strip()


def clean_payroll_code(code: str) -> str:
    """Chuẩn hóa mã chấm công, giữ nguyên số 0 ở đầu nếu có."""
    if code is None:
        return ""
    s = str(code).strip()
    # Nếu là float dạng "123.0" từ Excel
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    return s


# ==============================================================================
# Trích xuất từ Excel
# ==============================================================================

def extract_employees_from_excel(file_path: str) -> List[Dict[str, str]]:
    """
    Trích xuất danh sách nhân viên { 'name': ..., 'payroll_code': ... } từ file Excel (.xls, .xlsx).
    Hỗ trợ cả:
    1. Bảng danh sách chuẩn (Cột: Mã NV, Tên NV, ...)
    2. Bảng dạng Pivot / Chi tiết ca (Khối: 'ID:xxx Tên:yyy')
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    sheets = []
    if ext == ".xlsx":
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        try:
            sheets = [list(sheet.iter_rows(values_only=True)) for sheet in wb.worksheets]
        finally:
            wb.close()
    elif ext == ".xls":
        wb = xlrd.open_workbook(file_path)
        sheets = [[sheet.row_values(i) for i in range(sheet.nrows)] for sheet in wb.sheets()]
    else:
        raise ValueError("Chỉ chấp nhận .xls hoặc .xlsx")
    found = {}
    for rows in sheets:
        for employee in (_extract_pivot_format(rows) or _extract_table_format(rows)):
            # Retain different names sharing a code so the import can report conflicts.
            key = (employee['payroll_code'], unicodedata.normalize('NFC', employee['name']).casefold())
            found.setdefault(key, employee)
    return list(found.values())


def _extract_pivot_format(rows: List[List]) -> List[Dict[str, str]]:
    """Tìm các ô dạng: ID: 001 Tên: Nguyễn Văn A hoặc tương tự."""
    found_dict = {}
    ten_pat = r"t[eèéẻẽẹêềếểễệ]n"
    id_pat = r"(?:id|m[aàáảãạăằắẳẵặâầấẩẫậ]|\bma\b)"
    info_pat = re.compile(
        rf"{id_pat}[:\s]*([^\s]+).*?{ten_pat}[:\s]*(.+?)(?:ph[oòóỏõọôồốổỗộơờớởỡợ]ng|ca|$)",
        re.IGNORECASE
    )

    for row in rows:
        for val in row:
            if not val:
                continue
            text = str(val).strip()
            match = info_pat.search(text)
            if match:
                emp_id = clean_payroll_code(match.group(1))
                emp_name = clean_display_name(match.group(2))
                if emp_name and len(emp_name) >= 2:
                    key = (unicodedata.normalize('NFC', emp_name).casefold(), emp_id)
                    if key not in found_dict:
                        found_dict[key] = {
                            'name': emp_name,
                            'payroll_code': emp_id
                        }

    return list(found_dict.values())


def _extract_table_format(rows: List[List]) -> List[Dict[str, str]]:
    """Nhận diện header bảng có cột Mã NV và Tên NV."""
    col_map = {}
    header_idx = -1
    keywords = {
        'id': ['ma nhan vien', 'ma the', 'ma nv', 'manv', 'mathe', 'id', 'ma so'],
        'name': ['ten nhan vien', 'ho va ten', 'ho ten', 'ten', 'ho va ten nv', 'hoten', 'tennv']
    }

    for r_idx in range(min(40, len(rows))):
        row = rows[r_idx]
        current_map = {}
        for c_idx, val in enumerate(row):
            norm = normalize_vietnamese_name(str(val or ""))
            if not norm:
                continue
            for key, kws in keywords.items():
                if key not in current_map:
                    for kw in kws:
                        if kw == norm or (len(kw) > 3 and kw in norm):
                            current_map[key] = c_idx
                            break
        if 'id' in current_map and 'name' in current_map:
            col_map = current_map
            header_idx = r_idx
            break

    if header_idx == -1 or 'name' not in col_map:
        # Fallback: nếu không tìm thấy header rõ ràng, thử tìm hàng có độ tương đồng cao nhất
        for r_idx in range(min(40, len(rows))):
            row = rows[r_idx]
            for c_idx, val in enumerate(row):
                norm = normalize_vietnamese_name(str(val or ""))
                if any(kw in norm for kw in keywords['name']) and 'name' not in col_map:
                    col_map['name'] = c_idx
                    header_idx = r_idx
                elif any(kw in norm for kw in keywords['id']) and 'id' not in col_map:
                    col_map['id'] = c_idx

        if 'name' not in col_map:
            return []

    found_dict = {}
    id_col = col_map.get('id')
    name_col = col_map.get('name')

    for r_idx in range(header_idx + 1, len(rows)):
        row = rows[r_idx]
        if not row or name_col >= len(row):
            continue
        if not isinstance(row[name_col], str) or not row[name_col].strip():
            continue
        raw_name = clean_display_name(row[name_col])
        if not raw_name or len(raw_name) < 2:
            continue
        # Bỏ qua các dòng tiêu đề phụ lặp lại
        if any(kw in normalize_vietnamese_name(raw_name) for kw in keywords['name']):
            continue

        raw_id = clean_payroll_code(row[id_col]) if id_col is not None and id_col < len(row) else ""
        key = (unicodedata.normalize('NFC', raw_name).casefold(), raw_id)
        if key not in found_dict:
            found_dict[key] = {
                'name': raw_name,
                'payroll_code': raw_id
            }

    return list(found_dict.values())


# ==============================================================================
# Trích xuất từ PDF
# ==============================================================================

def extract_employees_from_pdf(file_path: str) -> List[Dict[str, str]]:
    """
    Trích xuất danh sách nhân viên { 'name': ..., 'payroll_code': ... } từ file PDF.
    Đọc nhanh từng trang qua PyMuPDF (fitz) và đối chiếu regex tên/mã.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    if not FITZ_AVAILABLE:
        raise RuntimeError("Thư viện PyMuPDF (fitz) chưa được cài đặt để đọc file PDF.")

    doc = fitz.open(file_path)
    found_dict = {}

    id_patterns = [
        r"M[aã]\s*nh[aâ]n\s*vi[eê]n\s*[:\s]*([^\s\n\r]+)",
        r"Ma\s*nhan\s*vien\s*[:\s]*([^\s\n\r]+)",
        r"M[aã]\s*th[eẻ]\s*[:\s]*([^\s\n\r]+)",
        r"ID[:\s]*([^\s\n\r]+)"
    ]
    name_patterns = [
        r"T[eê]n\s*nh[aâ]n\s*vi[eê]n\s*[:\s]*([^\n\r]+?)(?:Ph[oò]ng|Ca|$)",
        r"Ten\s*nhan\s*vien\s*[:\s]*([^\n\r]+?)(?:Phong|Ca|$)",
        r"H[oọ]\s*v[aà]\s*t[eê]n\s*[:\s]*([^\n\r]+?)(?:Ph[oò]ng|Ca|$)",
        r"H[oọ]\s*t[eê]n\s*[:\s]*([^\n\r]+?)(?:Ph[oò]ng|Ca|$)"
    ]

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text = page.get_text()
        if not text:
            continue

        # Thử tìm mã
        emp_id = ""
        for pat in id_patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                emp_id = clean_payroll_code(m.group(1))
                break

        # Thử tìm tên
        emp_name = ""
        for pat in name_patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                emp_name = clean_display_name(m.group(1))
                break

        if emp_name and len(emp_name) >= 2:
            key = (unicodedata.normalize('NFC', emp_name).casefold(), emp_id)
            if key not in found_dict:
                found_dict[key] = {
                    'name': emp_name,
                    'payroll_code': emp_id
                }

    doc.close()
    return list(found_dict.values())


# ==============================================================================
# Trích xuất tổng hợp từ File (Excel hoặc PDF)
# ==============================================================================

def extract_employees_from_file(file_path: str) -> List[Dict[str, str]]:
    """Tự động phát hiện loại file và trích xuất danh sách nhân viên."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".xls", ".xlsx"):
        return extract_employees_from_excel(file_path)
    elif ext == ".pdf":
        return extract_employees_from_pdf(file_path)
    else:
        raise ValueError(f"Loại file không hỗ trợ: {ext}. Vui lòng chọn file .xls, .xlsx hoặc .pdf")


# ==============================================================================
# Đồng bộ danh sách nhân viên vào Dự án & Thư mục Chân dung
# ==============================================================================

def sync_employees_to_project(
    project_id: str,
    employees: List[Dict[str, str]],
    identity_registry,
    default_valid_from: str = "2000-01-01",
    default_reviewer: str = "system"
) -> Dict:
    """
    Đồng bộ danh sách nhân viên và mã chấm công vào dự án:
    1. Quét nhân viên hiện có trong DB và thư mục ảnh chân dung của dự án.
    2. Khớp thông minh tên tiếng Việt (có dấu / không dấu / hoa thường).
    3. Nếu đã có thư mục ảnh cũ / legacy: Tự động gán mã và xác nhận liên kết ảnh.
    4. Nếu chưa có ảnh: Vẫn tạo nhân viên kèm mã sẵn, tạo thư mục trống để sẵn sàng thêm ảnh.
    5. Nếu đã có hồ sơ nhân viên nhưng chưa có mã hoặc mã cũ: Cập nhật membership với mã mới.
    """
    project = identity_registry.get_project(project_id)
    project_dir = identity_registry.project_portrait_dir(project_id)
    project_dir.mkdir(parents=True, exist_ok=True)

    # Đảm bảo nạp legacy sources hiện có vào registry trước
    identity_registry.import_legacy(project_id)

    # Lấy danh sách nhân viên hiện có trong DB của dự án
    existing_employees = identity_registry.list_employees(project_id)

    # Tạo map tra cứu nhân viên hiện có theo normalized name
    # norm_name -> employee dict
    existing_by_norm = {}
    for emp in existing_employees:
        norm = normalize_vietnamese_name(emp["display_name"])
        if norm:
            existing_by_norm[norm] = emp

    # Tạo map tra cứu các thư mục ảnh chân dung có sẵn trên ổ đĩa
    # norm_folder_name -> folder_path
    disk_folders_by_norm = {}
    if project_dir.exists():
        for item in project_dir.iterdir():
            if item.is_dir() and not item.is_symlink():
                norm = normalize_vietnamese_name(item.name)
                if norm:
                    disk_folders_by_norm[norm] = item

    results = {
        'total_found': len(employees),
        'bound_existing': 0,
        'created_new': 0,
        'updated_code': 0,
        'unchanged': 0,
        'internal_codes_created': 0,
        'details': []
    }

    with identity_registry._connect() as c:
        c.execute("BEGIN IMMEDIATE")
        used_codes = {r[0].upper() for r in c.execute("SELECT code FROM employee_internal_codes")}
        used_codes.update(r[0].upper() for r in c.execute("SELECT payroll_code FROM memberships"))

        for emp_item in employees:
            raw_name = clean_display_name(emp_item.get('name', ''))
            payroll_code = clean_payroll_code(emp_item.get('payroll_code', ''))
            if not raw_name or re.fullmatch(r"^[0-9a-fA-F]{32}$|^[0-9a-fA-F-]{36}$", raw_name):
                continue

            norm_name = normalize_vietnamese_name(raw_name)

            # Trường hợp 1: Đã có hồ sơ nhân viên trong DB
            if norm_name in existing_by_norm:
                emp = existing_by_norm[norm_name]
                eid = emp["employee_id"]
                current_memberships = emp.get("memberships", [])

                # Đảm bảo nhân viên có mã nội bộ
                internal_code, is_new_code = identity_registry._ensure_internal_code(c, eid, used_codes)
                if is_new_code:
                    results['internal_codes_created'] += 1

                # Kiểm tra xem đã có mã chấm công trùng chưa
                has_code = any(m.get("payroll_code") == payroll_code for m in current_memberships)
                
                # Nếu chưa có mã chấm công hoặc cần cập nhật mã
                if payroll_code and not has_code:
                    try:
                        identity_registry._assign(
                            c, project_id, eid, payroll_code, default_valid_from, None
                        )
                        results['updated_code'] += 1
                    except Exception as assign_err:
                        # Có thể đã có membership trùng khoảng ngày, bỏ qua
                        pass

                # Nếu là legacy source đang chờ xác nhận: tự động bind và xác nhận
                sources = emp.get("source_paths", [])
                if sources and emp.get("status") == "needs_confirmation":
                    for src_rel in sources:
                        try:
                            identity_registry._bind(c, project_id, eid, src_rel, default_reviewer)
                        except Exception:
                            pass
                    results['bound_existing'] += 1
                    status_note = "Đã ghép ảnh và gán mã"
                else:
                    results['unchanged'] += 1
                    status_note = "Đã có sẵn"

                results['details'].append({
                    'name': emp["display_name"],
                    'payroll_code': payroll_code,
                    'status': status_note,
                    'employee_id': eid,
                    'internal_code': internal_code
                })

            # Trường hợp 2: Chưa có hồ sơ nhân viên trong DB
            else:
                # Tạo nhân viên mới trong DB
                import uuid
                eid = uuid.uuid4().hex
                c.execute("INSERT INTO employees(employee_id, display_name) VALUES(?,?)", (eid, raw_name))

                # Đảm bảo nhân viên có mã nội bộ
                internal_code, is_new_code = identity_registry._ensure_internal_code(c, eid, used_codes)
                if is_new_code:
                    results['internal_codes_created'] += 1

                # Gán mã chấm công nếu có
                if payroll_code:
                    try:
                        identity_registry._assign(
                            c, project_id, eid, payroll_code, default_valid_from, None
                        )
                    except Exception:
                        pass

                # Kiểm tra xem trên ổ đĩa có thư mục ảnh trùng tên sẵn không
                bound_photo = False
                if norm_name in disk_folders_by_norm:
                    folder = disk_folders_by_norm[norm_name]
                    rel = folder.relative_to(project_dir).as_posix()
                    identity_registry._bind(c, project_id, eid, rel, default_reviewer)
                    bound_photo = True
                    results['bound_existing'] += 1
                else:
                    # Tạo thư mục chân dung rỗng theo employee_id để sẵn sàng upload ảnh
                    emp_folder = project_dir / eid
                    emp_folder.mkdir(parents=True, exist_ok=True)
                    identity_registry._bind(c, project_id, eid, eid, default_reviewer)
                    results['created_new'] += 1

                # Cập nhật lại cache tạm thời
                existing_by_norm[norm_name] = {
                    "employee_id": eid,
                    "display_name": raw_name,
                    "memberships": [{"payroll_code": payroll_code}] if payroll_code else [],
                    "source_paths": [],
                    "status": "confirmed",
                    "internal_code": internal_code
                }

                results['details'].append({
                    'name': raw_name,
                    'payroll_code': payroll_code,
                    'status': "Đã ghép ảnh có sẵn" if bound_photo else "Tạo mới (chưa có ảnh)",
                    'employee_id': eid,
                    'internal_code': internal_code
                })

    return results
