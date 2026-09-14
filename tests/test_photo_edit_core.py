import random

import src.photo_supplement as supplement


def test_normalizer_derives_vietnamese_weekday_and_seconds():
    normalizer = supplement.EditRequestNormalizer()

    result = normalizer.normalize_item("2026-08-29", "08:30")

    assert result == {
        "target_date": "2026-08-29",
        "target_time": "08:30:00",
        "weekday": "Thứ Bảy",
    }


def test_timestamp_analyzer_ignores_compass_and_address_text():
    analyzer = supplement.TimestampAnalyzer(ocr_engine=None)
    lines = [
        {"text": "77° E", "box": [1150, 760, 1270, 810], "score": 0.99},
        {"text": "Bến Lức Tây Ninh", "box": [1000, 820, 1270, 940], "score": 0.98},
        {"text": "14 Th9, 2026 20:52:10.931", "box": [810, 790, 1270, 835], "score": 0.93},
    ]

    result = analyzer.select_candidate(lines, image_size=(1280, 960))

    assert result["text"] == "14 Th9, 2026 20:52:10.931"
    assert result["status"] == "quick_ready"


def test_image_source_samples_at_most_31_without_duplicates(tmp_path):
    employee_dir = tmp_path / "Dự án A" / "Nguyễn Văn A"
    employee_dir.mkdir(parents=True)
    for index in range(40):
        (employee_dir / f"portrait-{index}.jpg").write_bytes(b"image")

    source = supplement.ImageSourceService(tmp_path, rng=random.Random(7))
    selected = source.sample_employee_images("Dự án A", "Nguyễn Văn A", count=99)

    assert len(selected) == 31
    assert len(set(selected)) == 31
    assert all(str(employee_dir) in path for path in selected)
