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

def test_legacy_mapping_requires_explicit_confirmation(tmp_path):
    c = make_client(tmp_path)
    old=tmp_path/"Site"/"01"
    old.mkdir(parents=True)
    (old/"original.png").write_bytes(photo().getvalue())
    assert c.get("/api/photos/daily/2026-09-01?project=Site").status_code == 409
    r=c.post("/api/photos/daily/legacy-map",json={"project":"Site","folder":"01","date":"2026-09-01","reviewer":"User"})
    assert r.status_code == 400
    r=c.post("/api/photos/daily/legacy-map",json={"project":"Site","folder":"01","date":"2026-09-01","reviewer":"User","confirm_single_period":True})
    assert r.status_code == 200
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
