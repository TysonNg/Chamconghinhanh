from datetime import date
import pytest
from src.excel_extractor import ExcelChamCongExtractor, ExcelToWordExporter

def person(code, name="Same Name", day="01/09/2026", incoming="08:00", source=""):
    return {"id":code,"name":name,"month":9,"year":2026,"source_id":source,
            "records":[{"date":day,"weekday":"Thứ ba","gio_vao":incoming,"gio_ra":"","missing_checkout":True,"is_absent":False}]}

def test_upstream_dedup_does_not_drop_same_name_different_codes():
    e=object.__new__(ExcelChamCongExtractor)
    e._persons_data=[person("001"),person("002")]
    e._dedupe_persons_by_name()
    assert len(e._persons_data) == 2

def test_merge_preserves_dates_and_flags_conflict():
    from src.attendance_records import merge_attendance_people
    values=[person("001",source="one"),person("001",day="02/09/2026",source="two"),
            person("001",incoming="09:00",source="three"),person("002",source="four")]
    result=merge_attendance_people(values)
    assert len(result) == 2
    first=next(p for p in result if p["id"]=="001")
    assert len(first["records"]) == 2
    conflict=next(r for r in first["records"] if r["date"]=="01/09/2026")
    assert conflict["record_conflict"] is True
    assert len(conflict["conflicting_records"]) == 2

def test_missing_codes_are_not_merged():
    from src.attendance_records import merge_attendance_people
    assert len(merge_attendance_people([person("",source="a"),person("",source="b")])) == 2

def test_person_reports_with_same_name_have_distinct_paths(tmp_path):
    ex=ExcelToWordExporter(str(tmp_path/"portraits"),str(tmp_path/"out"))
    one=ex.export_person(person("001"))
    two=ex.export_person(person("002"))
    assert one != two

def test_unknown_identity_never_uses_name_match(tmp_path):
    from src.attendance_matching import match_attendance_record
    class UnsafeMatcher:
        def match_face_in_images(self,*args,**kwargs):
            raise AssertionError("Name-only matching must not be used")
    rec=person("")["records"][0]
    match_attendance_record(rec,person(""),project_id="",registry=None,matcher=UnsafeMatcher(),
                            input_images_dir=tmp_path,threshold=None,accuracy_mode=True)
    assert rec["identity_status"] == "needs_confirmation"
    assert rec["matched_image_path"] is None


def test_splitter_keeps_both_employee_codes(tmp_path):
    from openpyxl import Workbook
    from src.excel_splitter import ExcelAttendanceSplitter
    wb=Workbook(); ws=wb.active
    ws.append(["STT","Mã nhân viên","Tên nhân viên","Phòng ban","Ngày","Thứ","Giờ vào","Giờ ra"])
    ws.append([1,"001","Same Name","A","01/09/2026","Ba","08:00","17:00"])
    ws.append([2,"002","Same Name","A","01/09/2026","Ba","08:01","17:01"])
    source=tmp_path/"input.xlsx"; wb.save(source)
    paths,summaries=ExcelAttendanceSplitter(str(source)).split(str(tmp_path/"split"))
    assert len(paths)==2
    assert {p["id"] for p in summaries} == {"001","002"}
