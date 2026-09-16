import io
import json
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from flask import Flask

from src.fake_photo_service import FakePhotoService
from src.supplement_batches import register_batches


def image_bytes(fmt="PNG", color="navy"):
    stream = io.BytesIO()
    Image.new("RGB", (80, 120), color).save(stream, fmt)
    return stream.getvalue()


def pending_analysis():
    return {
        "status": "needs_confirmation",
        "suggested_lines": ["20 Th12, 2025 07:59:46", "Đường Trần Trọng Cung", "Quận 7"],
        "confirmed_lines": [],
        "new_timestamp": "20 Th12, 2026 07:59:46",
        "block_box_2d": [700, 400, 990, 990],
        "provider_results": {"gemini": {"confidence": 0.7}},
    }


def test_create_preserves_original_bytes_when_automatic_generation_fails(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=pending_analysis()), \
         patch.object(service.smart_replacer, "replace_timestamp", return_value=False):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "ảnh gốc.png"
        )

    assert item["watermark_status"] == "verification_failed"
    original_path = service.raw_dir / item["original_file_name"]
    assert original_path.read_bytes() == raw
    with service._connect() as connection:
        columns = {row[1] for row in connection.execute("PRAGMA table_info(staging_photos)")}
    assert {
        "original_file_name", "watermark_status", "watermark_ocr_json",
        "confirmed_watermark_json", "generation_meta_json",
    }.issubset(columns)


def test_create_rejects_empty_block_ocr_without_timestamp_fallback(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    empty_ocr = {
        "status": "needs_confirmation", "confirmed_lines": [],
        "suggested_lines": [], "provider_results": {}, "block_box_2d": None,
    }

    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=empty_ocr), \
         patch.object(service.smart_replacer, "replace_timestamp") as fallback:
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-06-17", "08:15:40", "photo.png"
        )

    assert item["watermark_status"] == "verification_failed"
    assert item["generation_meta"]["error"] == "full_watermark_ocr_required"
    fallback.assert_not_called()


def test_regenerate_refuses_staging_fallback_when_original_is_missing(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=pending_analysis()):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )
    original_path = service.raw_dir / item["original_file_name"]
    original_path.unlink()
    staging = service.get_staging_image_path(item["id"])
    before = staging.read_bytes()

    result = service.regenerate_staging_photo(item["id"])

    assert result["status"] == "original_missing"
    assert staging.read_bytes() == before


def test_regenerate_uses_confirmed_lines_and_only_replaces_verified_candidate(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    analysis = pending_analysis()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=analysis):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )
    staging = service.get_staging_image_path(item["id"])
    before = staging.read_bytes()
    confirmed = ["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 7"]
    generated = image_bytes("JPEG", "green")

    with patch.object(service.ai_service, "generate_verified_watermark_crop", return_value={
        "status": "completed", "image_bytes": generated, "provider": "openai",
        "model": "gpt-image-2", "attempts": 1, "crop_box": [1, 2, 3, 4],
    }) as generate:
        result = service.regenerate_staging_photo(item["id"], confirmed_lines=confirmed)

    assert result["status"] == "completed"
    assert staging.read_bytes() != before
    block = generate.call_args.args[1]
    assert block["confirmed_lines"] == confirmed
    with service._connect() as connection:
        row = connection.execute(
            "SELECT confirmed_watermark_json, generation_meta_json FROM staging_photos WHERE id=?",
            (item["id"],),
        ).fetchone()
    assert json.loads(row[0])["lines"] == confirmed
    assert json.loads(row[1])["provider"] == "openai"


def test_regenerate_uses_ocr_suggestions_without_human_confirmation(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=pending_analysis()):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )

    generated = image_bytes("JPEG", "green")
    with patch.object(service.ai_service, "generate_verified_watermark_crop", return_value={
        "status": "completed", "image_bytes": generated, "provider": "gemini",
        "model": "gemini-image", "attempts": 1, "crop_box": [1, 2, 3, 4],
    }) as generate:
        result = service.regenerate_staging_photo(item["id"])

    assert result["status"] == "completed"
    assert generate.call_args.args[1]["confirmed_lines"] == [
        "20 Th12, 2026 07:59:46",
        "\u0110\u01b0\u1eddng Tr\u1ea7n Tr\u1ecdng Cung",
        "Qu\u1eadn 7",
    ]


def test_regenerate_reports_verification_failure_when_ocr_has_zero_lines(tmp_path):
    """With the >= 1 threshold, only 0 lines triggers verification_failed."""
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    empty = pending_analysis()
    empty["suggested_lines"] = []
    empty["block_box_2d"] = None
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=empty), \
         patch.object(service.smart_replacer, "replace_timestamp", return_value=False):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )

    with patch.object(service.smart_replacer, "replace_timestamp", return_value=False):
        result = service.regenerate_staging_photo(item["id"])

    assert result["status"] == "verification_failed"
    assert result["generation_meta"]["error"] == "full_watermark_ocr_required"


def test_regenerate_rejects_empty_saved_ocr_without_timestamp_fallback(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    empty_ocr = pending_analysis()
    empty_ocr.update({"suggested_lines": [], "provider_results": {}, "block_box_2d": None})
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=empty_ocr), \
         patch.object(service.smart_replacer, "replace_timestamp", return_value=False):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-06-17", "08:15:40", "photo.png"
        )

    with patch.object(service.smart_replacer, "replace_timestamp") as fallback:
        result = service.regenerate_staging_photo(item["id"])

    assert result["status"] == "verification_failed"
    assert result["generation_meta"]["error"] == "full_watermark_ocr_required"
    fallback.assert_not_called()


def test_failed_regeneration_keeps_current_result(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    analysis = pending_analysis()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=analysis):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )
    staging = service.get_staging_image_path(item["id"])
    before = staging.read_bytes()

    with patch.object(service.ai_service, "generate_verified_watermark_crop", return_value={
        "status": "verification_failed", "attempts": 4,
    }):
        result = service.regenerate_staging_photo(
            item["id"],
            confirmed_lines=["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 7"],
        )

    assert result["status"] == "verification_failed"
    assert staging.read_bytes() == before


def test_regenerate_api_maps_confirmation_verification_and_missing_original(tmp_path):
    app = Flask(__name__)
    register_batches(app, tmp_path)
    client = app.test_client()
    confirmed = ["20 Th12, 2026 07:59:46", "Quận 7"]

    with patch.object(FakePhotoService, "regenerate_staging_photo", return_value={
        "status": "needs_confirmation", "ocr": {"suggested_lines": ["Quận 7"]},
    }) as regenerate:
        response = client.post("/api/supplement/staging/photo-1/regenerate", json={
            "confirmed_lines": confirmed,
        })
    assert response.status_code == 409
    assert response.json["status"] == "needs_confirmation"
    regenerate.assert_called_once_with("photo-1", confirmed_lines=confirmed)

    with patch.object(FakePhotoService, "regenerate_staging_photo", return_value={
        "status": "verification_failed",
    }):
        assert client.post("/api/supplement/staging/photo-1/regenerate", json={}).status_code == 422

    with patch.object(FakePhotoService, "regenerate_staging_photo", return_value={
        "status": "original_missing",
    }):
        assert client.post("/api/supplement/staging/photo-1/regenerate", json={}).status_code == 404


def test_delete_staging_removes_immutable_original(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = image_bytes()
    with patch.object(service.ai_service, "is_configured", return_value=True), \
         patch.object(service.ai_service, "analyze_watermark_block", return_value=pending_analysis()):
        item = service.create_fake_photo(
            raw, "Site", "Employee", "2026-12-20", "07:59:46", "photo.png"
        )
    original = service.raw_dir / item["original_file_name"]
    assert original.exists()

    service.delete_staging_photos([item["id"]])

    assert not original.exists()


def test_easyocr_fallback_when_ai_quota_exceeded(tmp_path):
    from src.ai_timestamp_service import AITimestampService
    service = AITimestampService(str(tmp_path))
    mock_ocr = {
        "timestamp_line": {"text": "20:54 11 Tháng 2, 2026", "new_text": "08:12 23 Tháng 9, 2026", "box_2d": [750, 10, 850, 500]},
        "address_lines": [{"text": "Long An", "box_2d": [860, 20, 900, 400]}],
        "alignment": "left",
        "confidence": 0.95
    }
    with patch.object(service, "_available_providers", return_value=[("gemini", "key", "model")]), \
         patch.object(service, "_call_provider_watermark_ocr", return_value=None), \
         patch.object(service, "_detect_watermark_with_easyocr", return_value=mock_ocr) as mock_detect:
        result = service.analyze_watermark_block(image_bytes(), "2026-09-23", "08:12:00")
        assert mock_detect.called
        assert result["status"] == "confirmed"
        assert result["selected_provider"] == "easyocr"
        assert len(result["confirmed_lines"]) >= 2


def test_timemark_weekday_replacement(tmp_path):
    from src.smart_watermark_replacer import SmartWatermarkReplacer
    replacer = SmartWatermarkReplacer()
    
    # Target date: 2026-09-23 is Wednesday -> Thứ Tư
    orig = "20:54 11 Tháng 2, 2026 Thứ Tư"
    res = replacer._format_new_datetime(orig, "DD_THG_MM_YYYY", "2026-09-23", "08:12:00")
    assert "23" in res
    assert "08:12" in res
    assert "Thứ Tư" in res

    # Target date: 2026-12-01 is Tuesday -> Thứ Ba
    orig2 = "21:00 12 Tháng 2, 2026 Thứ Năm"
    res2 = replacer._format_new_datetime(orig2, "DD_THG_MM_YYYY", "2026-12-01", "08:00:00")
    assert "01" in res2 or "1" in res2
    assert "08:00" in res2
    assert "Thứ Ba" in res2

