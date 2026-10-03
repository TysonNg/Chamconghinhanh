import pytest
from supplement_test_support import environment
import io
from PIL import Image

def make_record(tmp_path, captured, requested="2026-09-01"):
    service, client, ids = environment(tmp_path)
    raw = io.BytesIO()
    exif = Image.Exif()
    if captured:
        exif[36867] = captured + " 08:00:00"
    Image.new("RGB", (20,20), "blue").save(raw,"JPEG",exif=exif)
    item = service.create_fake_photo(raw.getvalue(), "Site", "Employee", requested, "08:00:00",
                                     "original.jpg", replace_timestamp=False, modify_exif=False, **ids)
    return service, item, raw.getvalue()

def test_apply_original_to_full_date_and_keep_source(tmp_path):
    service, item, raw = make_record(tmp_path,"2026:09:01")
    wrong=service.input_images_dir/"Site"/"2026-08-01"
    wrong.mkdir(parents=True)
    result=service.apply_to_attendance([item["id"]])
    assert result["success"]
    dest=service.input_images_dir/"Site"/"2026-09-01"
    assert dest.is_dir()
    assert next(dest.glob("*.jpg")).read_bytes() == raw
    assert (service.raw_dir/item["original_file_name"]).read_bytes() == raw
    assert not list(wrong.iterdir())
    again=service.apply_to_attendance([item["id"]])
    assert again["success"]
    assert len(list(dest.glob("*.jpg"))) == 1

def test_mismatch_never_enters_attendance(tmp_path):
    service,item,raw=make_record(tmp_path,"2026:08:01")
    result=service.apply_to_attendance([item["id"]])
    assert result["success"] is False
    assert result["rejected"][0]["date_status"] == "mismatch"
    assert not list(service.input_images_dir.rglob("*.jpg"))

def test_unknown_date_is_not_verified_using_requested_date(tmp_path):
    service,item,raw=make_record(tmp_path,None)
    result=service.apply_to_attendance([item["id"]])
    assert result["success"] is False
    assert result["rejected"][0]["date_status"] == "unknown"

def test_empty_selection_does_not_apply_everything(tmp_path):
    service,item,raw=make_record(tmp_path,"2026:09:01")
    with pytest.raises(ValueError):
        service.apply_to_attendance([])
    assert service.get_staging_image_path(item["id"]).read_bytes() == raw
    assert not list(service.input_images_dir.rglob("*.jpg"))

def test_cleanup_preserves_applied_original_and_audit(tmp_path):
    service,item,raw=make_record(tmp_path,"2026:09:01")
    assert service.apply_to_attendance([item["id"]],delete_after=False)["success"]
    with pytest.raises(RuntimeError):
        service.delete_staging_photos([item["id"]])
    with pytest.raises(ValueError):
        service.delete_staging_photos([])
    with service._connect() as c:
        assert c.execute("SELECT action FROM supplement_history WHERE record_id=?", (item["id"],)).fetchall()
    assert (service.raw_dir/item["original_file_name"]).read_bytes() == raw
    assert service.apply_to_attendance([item["id"]])["success"]
