"""Temporary identity and deterministic image fixtures; no live AI or face models."""
import io
from unittest.mock import Mock
from flask import Flask
from PIL import Image
from src.identity_registry import IdentityRegistry
from src.face_matcher import MatchResult
from src.fake_photo_service import FakePhotoService
from src.supplement_batches import register_batches


def image_bytes(captured="2026:09:28 05:59:30", color="blue"):
    stream = io.BytesIO()
    exif = Image.Exif()
    if captured:
        exif[36867] = captured
    Image.new("RGB", (64, 64), color).save(stream, "JPEG", exif=exif)
    return stream.getvalue()


class Matcher:
    model_name = "test"
    detector_backend = "test"
    distance_metric = "cosine"
    def match_employee_in_photo(self, **kwargs):
        return MatchResult("matched", kwargs["image_path"], .1,
                           kwargs["project_id"], kwargs["employee_id"], "matched")
    def _get_default_threshold(self):
        return .37


def environment(tmp_path, project_name="Site", employee_name="Employee"):
    registry = IdentityRegistry(tmp_path / "identity.sqlite3", tmp_path / "portraits")
    project = registry.register_project(project_name)
    employee = registry.create_employee(employee_name)
    registry.assign_employee(project["project_id"], employee["employee_id"], "001", "2026-01-01", "2027-01-01")
    folder = registry.project_portrait_dir(project["project_id"])
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "portrait.jpg").write_bytes(image_bytes())
    registry.bind_portrait(project["project_id"], employee["employee_id"], "portrait.jpg", "test-reviewer")
    matcher = Matcher()
    service = FakePhotoService(str(tmp_path / "supplement_data"), lambda: registry, lambda: matcher)
    service.ai_service.inspect_photo = Mock(return_value={"status": "unconfigured", "visible_date": None})
    service.ai_service.analyze_watermark_block = Mock(return_value={"status": "verification_failed"})
    service.smart_replacer.replace_timestamp = Mock(side_effect=AssertionError("Unexpected live fallback"))
    app = Flask(__name__)
    app.testing = True
    register_batches(app, service.data_dir, registry_provider=lambda: registry,
                     matcher_provider=lambda: matcher, service=service)
    ids = {"project_id": project["project_id"], "employee_id": employee["employee_id"]}
    return service, app.test_client(), ids
