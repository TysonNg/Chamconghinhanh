import io
import json
import unicodedata
from pathlib import Path
from unittest.mock import patch, Mock

import numpy as np
import pytest
from PIL import Image
from flask import Flask

from src.ai_timestamp_service import AITimestampService
from src.fake_photo_service import FakePhotoService
from src.supplement_batches import register_batches


def create_test_image(size=(100, 120), color=(10, 20, 30)):
    img = Image.new("RGB", size, color)
    stream = io.BytesIO()
    img.save(stream, "PNG")
    return stream.getvalue()


# ==================== 1. TDD: OCR CONSENSUS & DISAGREEMENT ====================

def test_ocr_consensus_two_providers_agreed_locks_content(tmp_path):
    service = AITimestampService(str(tmp_path))
    provider_results = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46", "new_text": "20 Th12, 2026 07:59:46", "box_2d": [750, 400, 800, 950]},
            "address_lines": [{"text": "Đường Trần Trọng Cung", "box_2d": [810, 400, 860, 950]}, {"text": "Quận 7", "box_2d": [870, 400, 920, 950]}],
            "confidence": 0.95,
            "alignment": "right",
            "style": {"text_color": "#FFFFFF"}
        },
        "openai": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46", "new_text": "20 Th12, 2026 07:59:46", "box_2d": [755, 395, 805, 955]},
            "address_lines": [{"text": "Đường Trần Trọng Cung", "box_2d": [815, 395, 865, 955]}, {"text": "Quận 7", "box_2d": [875, 395, 925, 955]}],
            "confidence": 0.90,
            "alignment": "right",
            "style": {"text_color": "#FFFFFF"}
        }
    }

    consensus = service.resolve_watermark_consensus(provider_results)

    assert consensus["status"] == "confirmed"
    assert consensus["confirmed_lines"] == ["20 Th12, 2025 07:59:46", "Đường Trần Trọng Cung", "Quận 7"]
    assert consensus["new_timestamp"] == "20 Th12, 2026 07:59:46"
    assert consensus["timestamp_line"]["new_text"] == "20 Th12, 2026 07:59:46"
    assert len(consensus["address_lines"]) == 2
    assert consensus["confidence"] == 0.95
    assert consensus["block_box_2d"] == [750, 395, 925, 955]


def test_ocr_disagreement_automatically_uses_highest_confidence_provider(tmp_path):
    service = AITimestampService(str(tmp_path))
    provider_results = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46", "box_2d": [750, 400, 800, 950]},
            "address_lines": [{"text": "Đường Trần Trọng Cung", "box_2d": [810, 400, 860, 950]}, {"text": "Quận 7", "box_2d": [870, 400, 920, 950]}],
            "confidence": 0.95
        },
        "openai": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46", "box_2d": [750, 400, 800, 950]},
            "address_lines": [{"text": "Đường Trần Trọng Cung", "box_2d": [810, 400, 860, 950]}, {"text": "Quận 1", "box_2d": [870, 400, 920, 950]}],  # Disagreement here!
            "confidence": 0.90
        }
    }

    consensus = service.resolve_watermark_consensus(provider_results)

    assert consensus["status"] == "confirmed"
    assert consensus["confirmed_lines"] == [
        "20 Th12, 2025 07:59:46",
        "\u0110\u01b0\u1eddng Tr\u1ea7n Tr\u1ecdng Cung",
        "Qu\u1eadn 7",
    ]
    assert consensus["confidence"] == 0.95


def test_complete_low_confidence_ocr_is_accepted_without_confirmation(tmp_path):
    service = AITimestampService(str(tmp_path))
    provider_results = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46"},
            "address_lines": [{"text": "Quận 7"}],
            "confidence": 0.50  # Low confidence
        },
        "openai": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46"},
            "address_lines": [{"text": "Quận 7"}],
            "confidence": 0.50
        }
    }

    consensus = service.resolve_watermark_consensus(provider_results)
    assert consensus["status"] == "confirmed"
    assert consensus["confirmed_lines"] == ["20 Th12, 2025 07:59:46", "Qu\u1eadn 7"]


def test_single_timestamp_line_is_confirmed(tmp_path):
    """Watermark chỉ có 1 dòng timestamp (không có address) vẫn được confirmed."""
    service = AITimestampService(str(tmp_path))
    provider_results = {
        "gemini": {
            "timestamp_line": {"text": "23 Dec 2025 at 21:02:37", "new_text": "28 Sep 2026 at 08:01:29", "box_2d": [38, 415, 69, 953]},
            "address_lines": [],
            "confidence": 0.98,
            "alignment": "right",
        }
    }

    consensus = service.resolve_watermark_consensus(provider_results)
    assert consensus["status"] == "confirmed"
    assert consensus["confirmed_lines"] == ["23 Dec 2025 at 21:02:37"]
    assert consensus["new_timestamp"] == "28 Sep 2026 at 08:01:29"


def test_zero_lines_is_needs_confirmation(tmp_path):
    """Watermark không nhận diện được dòng nào thì trả needs_confirmation."""
    service = AITimestampService(str(tmp_path))
    provider_results = {
        "gemini": {
            "timestamp_line": {},
            "address_lines": [],
            "confidence": 0.0,
        }
    }

    consensus = service.resolve_watermark_consensus(provider_results)
    assert consensus["status"] == "needs_confirmation"
    assert consensus["confirmed_lines"] == []


# ==================== 2. TDD: UNICODE NFC & VIETNAMESE DIACRITICS ====================

def test_unicode_normalization_preserves_diacritics_and_handles_nfd_nfc():
    # Composed vs Decomposed forms of Vietnamese text
    composed = "Đường Trần Trọng Cung, Phường Tân Thuận Đông, Quận 7"
    decomposed = unicodedata.normalize("NFD", composed)
    assert composed != decomposed  # byte-level difference in representation

    norm1 = AITimestampService.normalize_watermark_text(composed)
    norm2 = AITimestampService.normalize_watermark_text(decomposed)

    assert norm1 == norm2
    assert norm1 == composed

    # Different accents must NEVER match
    accent_diff = AITimestampService.normalize_watermark_text("Đuong Tran Trong Cung, Phuong Tan Thuan Dong, Quan 7")
    assert norm1 != accent_diff


def test_missing_lines_rejected_but_provider_disagreement_is_auto_resolved(tmp_path):
    service = AITimestampService(str(tmp_path))

    # Only 1 line: now sufficient with >= 1 threshold
    single_line = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46"},
            "confidence": 0.99
        }
    }
    assert service.resolve_watermark_consensus(single_line)["status"] == "confirmed"

    # Provider disagreement no longer blocks generation; the best result wins.
    reversed_order = {
        "gemini": {
            "timestamp_line": {"text": "20 Th12, 2025 07:59:46"},
            "address_lines": [{"text": "Quận 7"}],
            "confidence": 0.95
        },
        "openai": {
            "timestamp_line": {"text": "Quận 7"},
            "address_lines": [{"text": "20 Th12, 2025 07:59:46"}],
            "confidence": 0.95
        }
    }
    resolved = service.resolve_watermark_consensus(reversed_order)
    assert resolved["status"] == "confirmed"
    assert resolved["confirmed_lines"] == ["20 Th12, 2025 07:59:46", "Qu\u1eadn 7"]


# ==================== 3. TDD: CROP UNION & PIXEL IDENTICALITY ====================

def test_crop_union_and_padding_bounds():
    image = Image.new("RGB", (200, 300), "blue")
    block = {
        "block_box_2d": [750, 400, 950, 950]  # normalized ymin, xmin, ymax, xmax
    }
    crop, pixel_box = AITimestampService.extract_watermark_crop(image, block, padding_ratio=0.20)
    left, top, right, bottom = pixel_box

    assert 0 <= left < right <= 200
    assert 0 <= top < bottom <= 300
    assert crop.size == (right - left, bottom - top)


def test_composite_watermark_crop_preserves_pixels_outside_crop():
    np.random.seed(42)
    orig_arr = np.random.randint(0, 256, (100, 100, 3), dtype=np.uint8)
    original = Image.fromarray(orig_arr, mode="RGB")

    pixel_box = (30, 40, 70, 80)
    crop_w = pixel_box[2] - pixel_box[0]
    crop_h = pixel_box[3] - pixel_box[1]
    edited_crop = Image.new("RGB", (crop_w, crop_h), (255, 0, 0))

    composited = AITimestampService.composite_watermark_crop(original, edited_crop, pixel_box)
    comp_arr = np.array(composited)

    # Every pixel strictly outside pixel_box must be IDENTICAL
    mask_outside = np.ones((100, 100), dtype=bool)
    mask_outside[pixel_box[1]:pixel_box[3], pixel_box[0]:pixel_box[2]] = False

    np.testing.assert_array_equal(comp_arr[mask_outside], orig_arr[mask_outside])


# ==================== 4. TDD: GENERATION PROMPT & VERIFICATION ====================

def test_generation_prompt_contains_exact_numbered_lines(tmp_path):
    service = AITimestampService(str(tmp_path))
    block = {
        "confirmed_lines": ["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 7"]
    }
    prompt = service._build_watermark_generation_prompt(block)

    assert "1. 20 Th12, 2026 07:59:46" in prompt
    assert "2. Đường Trần Trọng Cung" in prompt
    assert "3. Quận 7" in prompt
    assert "Preserve Vietnamese diacritics exactly" in prompt


def test_verify_watermark_crop_rejects_single_character_or_line_error(tmp_path):
    service = AITimestampService(str(tmp_path))
    expected = ["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 7"]

    # Exact match -> True
    with patch.object(service, "_ocr_watermark_lines", return_value={"gemini": expected}):
        assert service.verify_watermark_crop(b"fake_crop", expected) is True

    # 1 character wrong (e.g. Quận 1 instead of Quận 7) -> False
    bad_char = ["20 Th12, 2026 07:59:46", "Đường Trần Trọng Cung", "Quận 1"]
    with patch.object(service, "_ocr_watermark_lines", return_value={"gemini": bad_char}):
        assert service.verify_watermark_crop(b"fake_crop", expected) is False

    # Diacritic missing (Tran instead of Trần) -> False
    missing_accent = ["20 Th12, 2026 07:59:46", "Đường Tran Trọng Cung", "Quận 7"]
    with patch.object(service, "_ocr_watermark_lines", return_value={"gemini": missing_accent}):
        assert service.verify_watermark_crop(b"fake_crop", expected) is False

    # Swapped line order -> False
    swapped = ["Đường Trần Trọng Cung", "20 Th12, 2026 07:59:46", "Quận 7"]
    with patch.object(service, "_ocr_watermark_lines", return_value={"gemini": swapped}):
        assert service.verify_watermark_crop(b"fake_crop", expected) is False


# ==================== 5. TDD: RETRY & PROVIDER FALLBACK LIMITS ====================

def test_generate_verified_crop_retries_and_falls_back(tmp_path):
    service = AITimestampService(str(tmp_path))
    service.save_config({
        "provider": "gemini",
        "gemini_api_key": "key-gemini",
        "openai_api_key": "key-openai",
    })

    raw = create_test_image()
    block = {
        "confirmed_lines": ["20 Th12, 2026 07:59:46", "Quận 7"],
        "block_box_2d": [750, 400, 950, 950],
    }

    call_history = []

    def mock_edit(provider, key, model, crop_bytes, prompt, crop_size):
        call_history.append((provider, model))
        return create_test_image(size=crop_size)

    # All verification attempts fail:
    with patch.object(service, "_edit_watermark_crop_with_provider", side_effect=mock_edit), \
         patch.object(service, "verify_watermark_crop", return_value=False):
        result = service.generate_verified_watermark_crop(raw, block, max_attempts=2)

    assert result["status"] == "verification_failed"
    assert result["attempts"] == 4  # 2 attempts for gemini + 2 attempts for openai
    assert [c[0] for c in call_history] == ["gemini", "gemini", "openai", "openai"]


def test_generate_verified_crop_succeeds_on_fallback(tmp_path):
    service = AITimestampService(str(tmp_path))
    service.save_config({
        "provider": "gemini",
        "gemini_api_key": "key-gemini",
        "openai_api_key": "key-openai",
    })

    raw = create_test_image()
    block = {
        "confirmed_lines": ["20 Th12, 2026 07:59:46", "Quận 7"],
        "block_box_2d": [750, 400, 950, 950],
    }

    call_history = []
    def mock_edit(provider, key, model, crop_bytes, prompt, crop_size):
        call_history.append((provider, model))
        return create_test_image(size=crop_size)

    # Gemini fails on 2 calls, OpenAI succeeds on third call
    def mock_verify(crop_bytes, expected):
        return len(call_history) >= 3

    with patch.object(service, "_edit_watermark_crop_with_provider", side_effect=mock_edit), \
         patch.object(service, "verify_watermark_crop", side_effect=mock_verify):
        result = service.generate_verified_watermark_crop(raw, block, max_attempts=2)

    assert result["status"] == "completed"
    assert result["provider"] == "openai"
    assert result["attempts"] == 3


# ==================== 6. TDD: IMMUTABLE ORIGINAL & REGENERATE RESILIENCE ====================

def test_crop_watermark_fallback_when_no_block_box(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = create_test_image()
    item = service.create_fake_photo(raw, "Site", "Emp", "2026-12-20", "07:59:46", "sample.png", replace_timestamp=False)

    crop = service.get_watermark_crop_bytes(item["id"])
    assert crop is not None
    with Image.open(io.BytesIO(crop)) as img:
        assert img.width > 0 and img.height > 0
