# -*- coding: utf-8 -*-
"""
Phân tích khuôn mặt từ các file Word đã tách ra từ PDF.
Chuẩn hóa dữ liệu PDF về cùng format person/records như luồng Excel.
"""

import os
from src.attendance_records import merge_attendance_people
import re
import unicodedata
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document

from src.aggregate_report_exporter import export_aggregate_report
from src.excel_extractor import ExcelToWordExporter, is_same_or_close_time, is_under_work_duration


def _normalize_text(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = re.sub(r"\s+", " ", text.lower().strip())
    return text


def _clean_text(text: str) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).strip())


def _parse_date(text: str) -> Optional[date]:
    text = _clean_text(text)
    if not text:
        return None

    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass
    return None


def _extract_with_patterns(text: str, patterns: List[str]) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return _clean_text(match.group(1))
    return ""


def _cleanup_person_name(name: str) -> str:
    name = _clean_text(name)
    if not name:
        return ""

    # Loại suffix ngày bị dính vào tên khi convert từ PDF.
    name = re.sub(r"\s*[-–]\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*$", "", name)
    name = re.sub(r"\s+", " ", name).strip(" -")
    return name


class PDFPersonFileParser:
    """Parse một file docx đã tách từ PDF thành dữ liệu person/records."""

    def __init__(self, path: str):
        self.path = path
        self.doc = Document(path)
        self.rows = self._load_rows()
        self.lines = self._load_lines()

    def _load_rows(self) -> List[List[str]]:
        rows: List[List[str]] = []
        for table in self.doc.tables:
            for row in table.rows:
                values = [_clean_text(cell.text) for cell in row.cells]
                if any(values):
                    rows.append(values)
        return rows

    def _load_lines(self) -> List[str]:
        lines = []
        for para in self.doc.paragraphs:
            text = _clean_text(para.text)
            if text:
                lines.append(text)
        for row in self.rows[:6]:
            text = _clean_text(" ".join(v for v in row if v))
            if text:
                lines.append(text)
        return lines

    def _extract_metadata(self) -> Tuple[str, str, str]:
        person_name = ""
        person_id = ""
        department = ""

        for line in self.lines:
            if not person_id:
                person_id = _extract_with_patterns(line, [
                    r"Mã\s*nhân\s*viên\s*:\s*([^\s]+)",
                    r"Ma\s*nhan\s*vien\s*:\s*([^\s]+)",
                ])

            if not person_name:
                person_name = _extract_with_patterns(line, [
                    r"Tên\s*nhân\s*viên\s*:\s*(.*?)\s*Phòng\s*ban\s*:",
                    r"Ten\s*nhan\s*vien\s*:\s*(.*?)\s*Phong\s*ban\s*:",
                    r"Tên\s*nhân\s*viên\s*:\s*(.+)$",
                    r"Ten\s*nhan\s*vien\s*:\s*(.+)$",
                ])

            if not department:
                department = _extract_with_patterns(line, [
                    r"Phòng\s*ban\s*:\s*(.+)$",
                    r"Phong\s*ban\s*:\s*(.+)$",
                ])

            if person_name and person_id:
                break

        if not person_name:
            person_name = os.path.splitext(os.path.basename(self.path))[0]

        return (
            _cleanup_person_name(person_name),
            _clean_text(person_id),
            _clean_text(department),
        )

    def _find_detail_rows(self) -> Tuple[Optional[int], Optional[int]]:
        for idx in range(len(self.rows) - 1):
            row_norm = [_normalize_text(v) for v in self.rows[idx]]
            joined = " ".join(row_norm)
            has_date = any(v == "ngay" or "ngay" in v for v in row_norm)
            has_weekday = any(v == "thu" for v in row_norm)
            has_symbol = "ky hieu" in joined or "kyhieu" in joined
            if has_date and has_weekday and has_symbol:
                return idx, idx + 1
        return None, None

    def _find_column(self, row: List[str], keywords: List[str]) -> Optional[int]:
        for idx, value in enumerate(row):
            norm = _normalize_text(value)
            for keyword in keywords:
                if norm == keyword or keyword in norm:
                    return idx
        return None

    def _build_shift_pairs(self, header_row: List[str], subheader_row: List[str]) -> List[Tuple[int, int]]:
        header_norm = [_normalize_text(v) for v in header_row]
        sub_norm = [_normalize_text(v) for v in subheader_row]
        pairs: List[Tuple[int, int]] = []

        for idx, label in enumerate(header_norm):
            if label not in {"1", "2", "3"}:
                continue

            in_idx = None
            out_idx = None
            for pos in range(idx, min(len(sub_norm), idx + 3)):
                if in_idx is None and "vao" in sub_norm[pos]:
                    in_idx = pos
                elif in_idx is not None and "ra" in sub_norm[pos]:
                    out_idx = pos
                    break

            if in_idx is not None and out_idx is not None:
                pairs.append((in_idx, out_idx))

        if pairs:
            return pairs

        in_cols = [i for i, value in enumerate(header_norm) if "vao" in value]
        out_cols = [i for i, value in enumerate(header_norm) if "ra" in value]
        for in_idx, out_idx in zip(in_cols, out_cols):
            pairs.append((in_idx, out_idx))
        return pairs

    def _get_cell(self, row: List[str], idx: Optional[int]) -> str:
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        return _clean_text(row[idx])

    def parse_person(self) -> Optional[Dict]:
        detail_header_idx, detail_subheader_idx = self._find_detail_rows()
        if detail_header_idx is None or detail_subheader_idx is None:
            return None

        header_row = self.rows[detail_header_idx]
        subheader_row = self.rows[detail_subheader_idx]
        date_idx = self._find_column(header_row, ["ngay"])
        weekday_idx = self._find_column(header_row, ["thu"])
        shift_pairs = self._build_shift_pairs(header_row, subheader_row)

        if date_idx is None or not shift_pairs:
            return None

        person_name, person_id, department = self._extract_metadata()
        records = []
        month = None
        year = None

        for row in self.rows[detail_subheader_idx + 1:]:
            rec_date = _parse_date(self._get_cell(row, date_idx))
            if not rec_date:
                continue

            in_values: List[str] = []
            out_values: List[str] = []
            missing_checkin = False
            missing_checkout = False

            for in_idx, out_idx in shift_pairs:
                in_val = self._get_cell(row, in_idx)
                out_val = self._get_cell(row, out_idx)

                if in_val:
                    in_values.append(in_val)
                if out_val:
                    out_values.append(out_val)
                if not in_val and out_val:
                    missing_checkin = True
                if in_val and not out_val:
                    missing_checkout = True

            gio_vao = in_values[0] if in_values else ""
            gio_ra = out_values[-1] if out_values else ""
            same_in_out = False
            under_3h = False
            if gio_vao and gio_ra:
                same_in_out = is_same_or_close_time(gio_vao, gio_ra, max_diff_minutes=5)
                if not same_in_out:
                    under_3h = is_under_work_duration(gio_vao, gio_ra, max_hours=3.0)
            is_absent = not in_values and not out_values

            month = rec_date.month
            year = rec_date.year

            records.append({
                "day": rec_date.day,
                "date": rec_date.strftime("%d/%m/%Y"),
                "weekday": self._get_cell(row, weekday_idx),
                "gio_vao": gio_vao,
                "gio_ra": gio_ra,
                "is_absent": is_absent,
                "same_in_out": same_in_out,
                "under_3h": under_3h,
                "missing_checkout": missing_checkout and not is_absent and not same_in_out and not under_3h,
                "missing_checkin": missing_checkin and not is_absent and not same_in_out and not under_3h,
                "department": department,
            })

        if not person_name or not records:
            return None

        return {
            "id": person_id,
            "name": person_name,
            "month": month or datetime.now().month,
            "year": year or datetime.now().year,
            "records": records,
        }


class PDFFaceAnalyzer:
    def __init__(
        self,
        portrait_dir: str,
        input_images_dir: str,
        matcher,
        accuracy_mode: bool = True,
        match_distance_threshold: Optional[float] = None,
        log_detail: bool = False,
        project_id: str = "",
        identity_registry=None
    ):
        self.project_id = project_id
        self.identity_registry = identity_registry
        self.portrait_dir = portrait_dir
        self.input_images_dir = input_images_dir
        self.matcher = matcher
        self.accuracy_mode = accuracy_mode
        self.match_distance_threshold = match_distance_threshold
        self.log_detail = log_detail

    def _normalize_key(self, text: str) -> str:
        return _normalize_text(text)

    def _resolve_person_name(self, raw_name: str) -> str:
        return raw_name

    def analyze_folder(
        self,
        input_dir: str,
        output_dir: str,
        log_callback=None,
        progress_callback=None,
        report_options=None,
        cancel_check=None,
    ) -> List[str]:
        os.makedirs(output_dir, exist_ok=True)

        persons = []
        for f in os.listdir(input_dir):
            if not f.lower().endswith(".docx"):
                continue
            path = os.path.join(input_dir, f)
            parser = PDFPersonFileParser(path)
            person = parser.parse_person()
            if person:
                person["source_id"] = os.path.abspath(path)
                persons.append(person)
            elif log_callback:
                log_callback(f"⚠️ Không đọc được dữ liệu từ {f}", "warning")

        final_persons = merge_attendance_people(
            persons, project_id=self.project_id, registry=self.identity_registry)
        total_persons = len(final_persons)
        if log_callback:
            log_callback(f"📌 Tổng số người sẽ xử lý từ PDF: {total_persons}", "info")

        if total_persons == 0:
            return []

        exporter = ExcelToWordExporter(
            self.portrait_dir,
            output_dir,
            input_images_dir=self.input_images_dir,
            face_matcher=self.matcher,
            accuracy_mode=self.accuracy_mode,
            match_distance_threshold=self.match_distance_threshold,
            log_detail=self.log_detail,
            project_id=self.project_id,
            identity_registry=self.identity_registry
        )

        # Tự động tính số luồng tối ưu (50% - 75% CPU)
        cpu_cores = os.cpu_count() or 4
        workers = max(2, min(6, int(cpu_cores * 0.75)))
        if log_callback:
            log_callback(f"🚀 Kích hoạt quét song song đa luồng ({workers} workers / {cpu_cores} nhân CPU)", "info")

        results = []
        completed_count = 0
        lock = threading.Lock()

        def _process_one_person(person):
            nonlocal completed_count
            if cancel_check and cancel_check():
                return None
            name = person["name"]
            issue_days = sum(
                1 for r in person["records"]
                if r.get("is_absent") or r.get("missing_checkout") or r.get("missing_checkin") or r.get("same_in_out") or r.get("under_3h")
            )
            try:
                out_path = exporter.export_person(person, log_callback=None)
                with lock:
                    completed_count += 1
                    pct = int((completed_count / total_persons) * 100)
                    if log_callback:
                        log_callback(f"👤 [{completed_count}/{total_persons} - {pct}%] Xong: {name} ({issue_days} ngày đối soát)", "default")
                    if progress_callback:
                        progress_callback(completed_count, total_persons, name, out_path)
                return out_path
            except Exception as e:
                with lock:
                    completed_count += 1
                    pct = int((completed_count / total_persons) * 100)
                    if log_callback:
                        log_callback(f"❌ [{completed_count}/{total_persons}] Lỗi xuất {name}: {e}", "error")
                    if progress_callback:
                        progress_callback(completed_count, total_persons, name, None)
                return None

        # Chạy song song qua ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_person = {}
            for p in final_persons:
                if cancel_check and cancel_check():
                    break
                future_to_person[executor.submit(_process_one_person, p)] = p

            for future in as_completed(future_to_person):
                p_path = future.result()
                if p_path:
                    results.append(p_path)
                if cancel_check and cancel_check():
                    for f in future_to_person:
                        f.cancel()
                    break

        # Lưu cache vector khuôn mặt xuống ổ đĩa
        if self.matcher and hasattr(self.matcher, "save_cache"):
            self.matcher.save_cache()

        # Nếu có yêu cầu hủy, dừng lại và bảo toàn các file đã tạo, không xuất báo cáo tổng hợp dở dang
        if cancel_check and cancel_check():
            if log_callback:
                log_callback(f"⚠️ Đã dừng quét theo yêu cầu! Đã giữ lại {len(results)} file Word cá nhân hoàn tất.", "warning")
            return results

        if report_options is not None:
            report_path = export_aggregate_report(
                final_persons,
                output_dir=output_dir,
                **report_options,
            )
            results.append(report_path)
            if log_callback:
                log_callback(
                    f"📄 Đã tạo giải trình tổng hợp: {os.path.basename(report_path)}",
                    "success",
                )

        return results
