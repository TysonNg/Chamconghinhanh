# -*- coding: utf-8 -*-
"""
Unit tests for AITimestampService and AI-integrated SmartWatermarkReplacer
"""

import json
import io
import unicodedata
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PIL import Image

from src.ai_timestamp_service import AITimestampService
from src.smart_watermark_replacer import SmartWatermarkReplacer, hex_to_rgb


def test_hex_to_rgb():
    assert hex_to_rgb("#FFFFFF") == (255, 255, 255)
    assert hex_to_rgb("000000") == (0, 0, 0)
    assert hex_to_rgb("#FF0000") == (255, 0, 0)
    assert hex_to_rgb("#FFF") == (255, 255, 255)
    assert hex_to_rgb("invalid", default=(10, 20, 30)) == (10, 20, 30)


def test_ai_service_config(tmp_path):
    svc = AITimestampService(str(tmp_path))
    assert not svc.is_configured()

    # Save Gemini config
    success = svc.save_config({
        "provider": "gemini",
        "gemini_api_key": "AIzaSyTest1234567890",
        "gemini_model": "gemini-3.6-flash",
        "prompt_template": "Custom prompt {original_text}"
    })
    assert success
    assert svc.is_configured()

    pub = svc.get_public_config()
    assert pub["provider"] == "gemini"
    assert pub["has_gemini_key"] is True
    assert "AIza" in pub["gemini_api_key"]
    assert "..." in pub["gemini_api_key"]
    assert pub["is_active"] is True
    assert pub["prompt_template"] == "Custom prompt {original_text}"


def test_ai_service_parse_json():
    # Plain json
    raw1 = '{"text_color": "#FFFF00", "has_shadow": true}'
    assert AITimestampService._parse_json_result(raw1) == {"text_color": "#FFFF00", "has_shadow": True}

    # Markdown codeblock json
    raw2 = '```json\n{"text_color": "#FFFFFF", "stroke_width": 2}\n```'
    assert AITimestampService._parse_json_result(raw2) == {"text_color": "#FFFFFF", "stroke_width": 2}


def test_ai_test_empty_key(tmp_path):
    svc = AITimestampService(str(tmp_path))
    ok, msg = svc.test_connection("gemini", "", "gemini-3.6-flash")
    assert not ok
    assert "không được để trống" in msg


def test_smart_replacer_with_ai_mock(tmp_path):
    # Mock AI service
    ai_mock = MagicMock(spec=AITimestampService)
    ai_mock.is_configured.return_value = True
    ai_mock.analyze_timestamp_style.return_value = {
        "text_color": "#FFFFFF",
        "has_shadow": True,
        "shadow_color": "#000000",
        "shadow_offset": [1, 1],
        "has_stroke": True,
        "stroke_color": "#111111",
        "stroke_width": 1
    }

    replacer = SmartWatermarkReplacer(ai_service=ai_mock)
    assert replacer.ai_service is not None

    # Test datetime format helper
    original = "14 Th9, 2026 20:52:10"
    res = replacer._format_new_datetime(original, "DD_THG_MM_YYYY", "2026-10-25", "08:15:00")
    assert "25" in res
    assert "10" in res
    assert "08:15:00" in res


def test_image_edit_prompt_uses_detected_timestamp_geometry_and_style(tmp_path):
    svc = AITimestampService(str(tmp_path))
    svc.save_config({
        "provider": "gemini",
        "gemini_api_key": "test-key",
        "gemini_model": "gemini-3.6-flash",
    })
    stream = io.BytesIO()
    Image.new("RGB", (800, 1200), "white").save(stream, "JPEG")
    timestamp_info = {
        "box_2d": [850, 610, 885, 995],
        "alignment": "right",
        "right_margin_x": 995,
        "text_color": "#C2C2C2",
        "font_weight": "regular",
        "has_shadow": True,
        "shadow_color": "#202020",
        "shadow_offset": [1, 2],
    }

    with patch("src.ai_timestamp_service.requests.post") as post:
        post.return_value.status_code = 503
        svc.edit_timestamp_image(
            stream.getvalue(),
            "2026-06-30",
            "08:13:24",
            original_text="30 Th6, 2025 08.13.24",
            new_text="30 Th6, 2026 08.13.24",
            timestamp_info=timestamp_info,
        )

    prompt = post.call_args.kwargs["json"]["contents"][0]["parts"][0]["text"]
    assert "[850, 610, 885, 995]" in prompt
    assert "#C2C2C2" in prompt
    assert "regular" in prompt
    assert "[1, 2]" in prompt
    assert "sao chép trực tiếp từ chính dòng timestamp cũ" in prompt
    assert "không lấy kích thước chữ từ các dòng địa chỉ" in prompt


def test_replacer_passes_detected_timestamp_info_to_image_editor(tmp_path):
    source = tmp_path / "source.jpg"
    output = tmp_path / "output.jpg"
    Image.new("RGB", (320, 480), "gray").save(source, "JPEG")
    raw = source.read_bytes()
    timestamp_info = {
        "box_2d": [850, 610, 885, 995],
        "original_text": "30 Th6, 2025 08.13.24",
        "new_text": "30 Th6, 2026 08.13.24",
        "text_color": "#C2C2C2",
    }
    ai_mock = MagicMock(spec=AITimestampService)
    ai_mock.is_configured.return_value = True
    ai_mock.detect_and_analyze_watermark.return_value = timestamp_info
    ai_mock.edit_timestamp_image.return_value = raw
    ai_mock.get_full_config.return_value = {"anti_ai_enabled": False}

    result = SmartWatermarkReplacer(ai_mock).replace_timestamp(
        str(source), str(output), "2026-06-30", "08:13:24"
    )

    assert result is True
    assert ai_mock.edit_timestamp_image.call_args.kwargs["timestamp_info"] == timestamp_info


def test_watermark_consensus_normalizes_unicode_but_keeps_case_and_punctuation(tmp_path):
    svc = AITimestampService(str(tmp_path))
    decomposed = unicodedata.normalize("NFD", "Đường Trần Trọng Cung")
    results = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46", "box_2d": [800, 500, 840, 990]},
            "address_lines": [
                {"text": decomposed, "box_2d": [840, 500, 890, 990]},
                {"text": "Quận 7", "box_2d": [890, 700, 930, 990]},
            ],
            "confidence": 0.97,
        },
        "openai": {
            "timestamp_line": {"text": "20  Th12, 2025 07:59:46", "box_2d": [801, 501, 841, 991]},
            "address_lines": [
                {"text": "Đường Trần Trọng Cung", "box_2d": [841, 501, 891, 991]},
                {"text": "Quận 7", "box_2d": [891, 701, 931, 991]},
            ],
            "confidence": 0.96,
        },
    }

    consensus = svc.resolve_watermark_consensus(results)

    assert consensus["status"] == "confirmed"
    assert consensus["confirmed_lines"][1] == "Đường Trần Trọng Cung"
    assert consensus["block_box_2d"] == [800, 500, 931, 991]

    results["openai"]["address_lines"][1]["text"] = "quận 7"
    resolved = svc.resolve_watermark_consensus(results)
    assert resolved["status"] == "confirmed"
    assert resolved["confirmed_lines"][2] == "Qu\u1eadn 7"


def test_crop_box_is_padded_clamped_and_composite_preserves_outside_pixels(tmp_path):
    svc = AITimestampService(str(tmp_path))
    source = Image.new("RGB", (100, 80), (10, 20, 30))
    block = {"block_box_2d": [750, 800, 1000, 1000]}

    crop, box = svc.extract_watermark_crop(source, block, padding_ratio=0.25)
    assert box == (75, 55, 100, 80)
    assert crop.size == (25, 25)

    edited = Image.new("RGB", crop.size, (200, 210, 220))
    result = svc.composite_watermark_crop(source, edited, box)
    assert result.getpixel((74, 54)) == (10, 20, 30)
    assert result.getpixel((99, 79)) == (200, 210, 220)


def test_generated_watermark_requires_exact_verified_lines_and_falls_back_provider(tmp_path):
    svc = AITimestampService(str(tmp_path))
    svc.save_config({
        "provider": "gemini",
        "gemini_api_key": "gem-key",
        "openai_api_key": "oa-key",
    })
    source = Image.new("RGB", (160, 120), (50, 60, 70))
    raw = io.BytesIO()
    source.save(raw, "PNG")
    block = {
        "block_box_2d": [700, 500, 980, 980],
        "confirmed_lines": ["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 7"],
        "alignment": "right",
        "style": {"text_color": "#FFFFFF"},
    }

    def fake_edit(provider, *_args, **_kwargs):
        crop_size = _kwargs["crop_size"]
        image = Image.new("RGB", crop_size, "white" if provider == "gemini" else "black")
        buf = io.BytesIO()
        image.save(buf, "PNG")
        return buf.getvalue()

    with patch.object(svc, "_edit_watermark_crop_with_provider", side_effect=fake_edit), \
         patch.object(svc, "verify_watermark_crop", side_effect=[False, False, True]) as verify:
        result = svc.generate_verified_watermark_crop(raw.getvalue(), block, max_attempts=2)

    assert result["status"] == "completed"
    assert result["provider"] == "openai"
    assert result["attempts"] == 3
    assert verify.call_count == 3
