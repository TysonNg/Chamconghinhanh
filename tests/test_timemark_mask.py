import cv2
import numpy as np

from src.smart_watermark_replacer import SmartWatermarkReplacer, TimestampRegion


def _timemark_region():
    return TimestampRegion(
        bbox=[[28, 18], [122, 18], [122, 82], [28, 82]],
        text="07:46 | 08 Thang 6, 2026 Thu Hai",
        confidence=0.99,
        estimated_font_size=48,
        is_timemark=True,
        timemark_info={"time_box": [[32, 24], [78, 24], [78, 72], [32, 72]]},
    )


def _shadowed_timemark_image():
    image = np.full((110, 180, 3), 210, dtype=np.uint8)

    # A simplified bright glyph and its detached dark drop shadow.
    image[30:68, 38:47] = (250, 250, 250)
    image[33:71, 50:55] = (45, 45, 45)

    # Timemark's gold separator.
    image[25:76, 82:85] = (40, 170, 220)

    # Unrelated dark scene detail inside the OCR rectangle.
    image[22:78, 112:116] = (35, 35, 35)
    return image


def test_timemark_mask_covers_detached_shadow_without_masking_scene_detail():
    replacer = SmartWatermarkReplacer()
    image = _shadowed_timemark_image()

    mask = replacer._build_timemark_mask(image, _timemark_region())

    assert mask.shape == image.shape[:2]
    assert mask.dtype == np.uint8
    assert np.count_nonzero(mask[35:66, 50:55]) >= 120
    assert np.count_nonzero(mask[24:76, 112:116]) == 0


def test_timemark_inpaint_preserves_pixels_outside_adaptive_mask():
    replacer = SmartWatermarkReplacer()
    image = _shadowed_timemark_image()
    mask = replacer._build_timemark_mask(image, _timemark_region())

    result = cv2.inpaint(
        image,
        mask,
        inpaintRadius=replacer._timemark_inpaint_radius(mask),
        flags=cv2.INPAINT_NS,
    )

    outside = mask == 0
    np.testing.assert_array_equal(result[outside], image[outside])


def test_timemark_mask_stays_empty_when_no_bright_or_gold_core_exists():
    replacer = SmartWatermarkReplacer()
    image = np.full((110, 180, 3), 90, dtype=np.uint8)
    image[30:70, 50:55] = (25, 25, 25)

    mask = replacer._build_timemark_mask(image, _timemark_region())

    assert np.count_nonzero(mask) == 0


def test_generic_mask_detects_dark_blue_watermark_from_detected_text_color():
    replacer = SmartWatermarkReplacer()
    image = np.full((100, 220, 3), (105, 115, 125), dtype=np.uint8)
    image[25:70, 18:205] = (90, 105, 120)

    # Dark-blue glyph on a textured roof-like background (BGR input).
    image[34:61, 42:55] = (63, 31, 0)
    image[34:61, 64:77] = (63, 31, 0)
    # A dark red scene detail inside the OCR box must not be selected.
    image[30:68, 175:181] = (30, 30, 95)

    region = TimestampRegion(
        bbox=[[34, 28], [188, 28], [188, 68], [34, 68]],
        text="23 Dec 2025 at 21:02:37",
        confidence=0.99,
        detected_color=(0, 31, 63),
        estimated_font_size=28,
    )

    mask = replacer._build_text_mask(image, region.bbox, detected_color=region.detected_color)

    assert np.count_nonzero(mask[36:59, 42:77]) >= 500
    assert np.count_nonzero(mask[32:66, 175:181]) == 0
