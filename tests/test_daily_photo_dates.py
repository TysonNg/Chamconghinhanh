import io
from pathlib import Path
import pytest
from PIL import Image
from flask import Flask

def make_client(tmp_path):
    from src.daily_photo_routes import register_daily_photo_routes
    app = Flask(__name__)
    app.testing = True
    register_daily_photo_routes(app, lambda: tmp_path)
    return app.test_client()

def photo():
    f = io.BytesIO()
    Image.new("RGB",(12,12), "navy").save(f,"PNG")
    f.seek(0)
    return f

def test_upload_view_delete_are_scoped_to_full_date(tmp_path):
    c = make_client(tmp_path)
    for day in ("2026-08-01","2026-09-01"):
        r = c.post("/api/photos/daily/upload", data={"project":"Site","date":day,"files":(photo(),"same.png")})
        assert r.status_code == 200
    r = c.get("/api/photos/daily/2026-09-01?project=Site")
    assert len(r.json["photos"]) == 1
    assert c.get(r.json["photos"][0]["url"]).status_code == 200
    stats = c.get("/api/photos/daily?project=Site&period=2026-09").json
    assert len(stats["days"]) == 30
    assert stats["days"][0]["image_count"] == 1
    assert c.post("/api/photos/daily/delete",json={"project":"Site","date":"2026-09-01","delete_all":True}).status_code == 200
    assert (tmp_path/"Site"/"2026-08-01"/"same.png").exists()
    assert not (tmp_path/"Site"/"2026-09-01"/"same.png").exists()

@pytest.mark.parametrize("day", ["01","2026-02-30","2026-9-1","../../outside"])
def test_incomplete_or_invalid_date_never_writes(tmp_path, day):
    c = make_client(tmp_path)
    r = c.post("/api/photos/daily/upload",data={"project":"Site","date":day,"files":(photo(),"photo.png")})
    assert r.status_code == 400
    assert not list(tmp_path.rglob("photo.png"))

def test_legacy_mapping_requires_safety_confirmation_but_not_reviewer_name(tmp_path):
    c = make_client(tmp_path)
    old=tmp_path/"Site"/"01"
    old.mkdir(parents=True)
    (old/"original.png").write_bytes(photo().getvalue())
    assert c.get("/api/photos/daily/2026-09-01?project=Site").status_code == 409
    r=c.post("/api/photos/daily/legacy-map",json={"project":"Site","folder":"01","date":"2026-09-01"})
    assert r.status_code == 400
    r=c.post("/api/photos/daily/legacy-map",json={"project":"Site","folder":"01","date":"2026-09-01","confirm_single_period":True})
    assert r.status_code == 200
    assert r.json["mapping"]["confirmed_by"] == "system"
    assert len(c.get("/api/photos/daily/2026-09-01?project=Site").json["photos"]) == 1
    assert len(c.get("/api/photos/daily/2026-08-01?project=Site").json["photos"]) == 0

def test_upload_does_not_overwrite_existing_photo(tmp_path):
    c=make_client(tmp_path)
    for _ in range(2):
        assert c.post("/api/photos/daily/upload",data={"project":"Site","date":"2026-09-01","files":(photo(),"same.png")}).status_code == 200
    assert len(list((tmp_path/"Site"/"2026-09-01").glob("*.png"))) == 2

def test_invalid_project_and_filename_cannot_escape(tmp_path):
    c=make_client(tmp_path)
    assert c.post("/api/photos/daily/upload",data={"project":"..","date":"2026-09-01","files":(photo(),"a.png")}).status_code == 400
    assert c.post("/api/photos/daily/delete",json={"project":"Site","date":"2026-09-01","filename":"../outside"}).status_code == 400

def test_delete_all_days_scoped_to_period(tmp_path):
    c = make_client(tmp_path)
    # Upload to August and September
    c.post("/api/photos/daily/upload", data={"project":"Site","date":"2026-08-15","files":(photo(),"aug.png")})
    c.post("/api/photos/daily/upload", data={"project":"Site","date":"2026-09-01","files":(photo(),"sep1.png")})
    c.post("/api/photos/daily/upload", data={"project":"Site","date":"2026-09-02","files":(photo(),"sep2.png")})

    # Delete all days in September 2026
    r = c.post("/api/photos/daily/delete", json={"project":"Site", "period":"2026-09", "delete_all_days":True, "scope":"period"})
    assert r.status_code == 200
    assert r.json["deleted_count"] == 2

    # Verify September photos are gone and August remains
    assert not (tmp_path / "Site" / "2026-09-01" / "sep1.png").exists()
    assert not (tmp_path / "Site" / "2026-09-02" / "sep2.png").exists()
    assert (tmp_path / "Site" / "2026-08-15" / "aug.png").exists()

def test_delete_all_days_entire_project(tmp_path):
    c = make_client(tmp_path)
    c.post("/api/photos/daily/upload", data={"project":"Site","date":"2026-08-15","files":(photo(),"aug.png")})
    c.post("/api/photos/daily/upload", data={"project":"Site","date":"2026-09-01","files":(photo(),"sep1.png")})

    r = c.post("/api/photos/daily/delete", json={"project":"Site", "delete_all_days":True, "scope":"all"})
    assert r.status_code == 200
    assert r.json["deleted_count"] == 2
    assert not (tmp_path / "Site" / "2026-08-15" / "aug.png").exists()
    assert not (tmp_path / "Site" / "2026-09-01" / "sep1.png").exists()



def test_listing_exposes_shift_without_losing_legacy_or_bad_metadata(tmp_path):
    import json
    c = make_client(tmp_path)
    folder = tmp_path / "Site" / "2026-10-01"
    folder.mkdir(parents=True)
    for name in ("morning.png", "afternoon.png", "legacy.png", "broken.png"):
        (folder / name).write_bytes(photo().getvalue())
    for shift, time in (("morning", "05:00:00"), ("afternoon", "16:00:00")):
        (folder / (shift + ".png.json")).write_text(json.dumps({"shift": shift, "send_time": time}), encoding="utf-8")
    (folder / "broken.png.json").write_text("broken", encoding="utf-8")
    response = c.get("/api/photos/daily/2026-10-01?project=Site")
    assert response.status_code == 200
    photos = {p["filename"]: p for p in response.json["photos"]}
    assert len(photos) == 4
    assert photos["morning.png"]["shift"] == "morning"
    assert photos["afternoon.png"]["send_time"] == "16:00:00"
    assert photos["legacy.png"]["shift"] == "unknown"
    assert photos["broken.png"]["shift"] == "unknown"
