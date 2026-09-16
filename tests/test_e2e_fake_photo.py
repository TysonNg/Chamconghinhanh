import io
import json
import sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, str(Path(__file__).parent.parent))

from PIL import Image
from src.fake_photo_service import FakePhotoService

def test_fake_photo_generation():
    # Use real staging image or generate dummy image
    staging_files = list(Path('supplement_data/supplement_staging').glob('*.jpg'))
    if staging_files:
        raw_bytes = staging_files[0].read_bytes()
    else:
        img = Image.new('RGB', (400, 300), color=(200, 200, 200))
        buf = io.BytesIO()
        img.save(buf, format='JPEG')
        raw_bytes = buf.getvalue()

    service = FakePhotoService('supplement_data')
    
    # Test random time generation
    morning_time = service.generate_random_time("morning")
    afternoon_time = service.generate_random_time("afternoon")
    print(f"Generated morning time: {morning_time}")
    print(f"Generated afternoon time: {afternoon_time}")
    
    assert 5 <= int(morning_time.split(':')[0]) <= 6
    assert 17 <= int(afternoon_time.split(':')[0]) <= 18

    target_date = "2026-09-28"
    from unittest.mock import patch
    confirmed_analysis = {
        "status": "confirmed",
        "suggested_lines": ["28 Th9, 2026 06:15:00", "Đường Trần Trọng Cung", "Quận 7"],
        "confirmed_lines": ["28 Th9, 2026 06:15:00", "Đường Trần Trọng Cung", "Quận 7"],
        "new_timestamp": f"28 Th9, 2026 {morning_time}",
        "block_box_2d": [750, 400, 990, 990],
        "provider_results": {"gemini": {"confidence": 0.95}},
    }
    with patch.object(service.ai_service, "analyze_watermark_block", return_value=confirmed_analysis), \
         patch.object(service.ai_service, "generate_verified_watermark_crop", return_value={
             "status": "completed", "image_bytes": raw_bytes, "provider": "gemini",
             "model": "gemini-3.1-flash-image", "attempts": 1, "crop_box": [10, 10, 50, 50],
         }):
        item = service.create_fake_photo(
            raw_bytes=raw_bytes,
            project="Chung cư Tân Thuận Đông",
            employee="Dang Van Tau",
            target_date=target_date,
            target_time=morning_time,
            original_name="test_tau.jpg",
            replace_timestamp=True,
            modify_exif=True
        )

    print("Created Item:", json.dumps(item, ensure_ascii=False, indent=2))
    assert item['target_date'] == target_date
    assert item['target_time'] == morning_time
    assert item['inspection']['status'] == 'ok'
    assert item['inspection']['date_mismatch'] is False

    # Check generated file
    file_path = service.get_staging_image_path(item['id'])
    assert file_path is not None
    assert file_path.exists()
    
    # Verify EXIF
    with Image.open(file_path) as img:
        exif = img.getexif()
        captured = exif.get(306)
        print("Updated EXIF in file:", captured)
        assert captured.startswith("2026:09:28")

    # Clean up test artifact from staging
    service.delete_staging_photos([item['id']])

    print("\nALL ASSERTIONS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_fake_photo_generation()
