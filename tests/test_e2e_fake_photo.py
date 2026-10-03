import hashlib
from unittest.mock import patch
from PIL import Image
from supplement_test_support import environment, image_bytes


def test_fake_photo_generation(tmp_path):
    service, client, ids = environment(tmp_path)
    raw = image_bytes('2026:09:27 06:15:00')
    target_time = '06:15:00'
    analysis = {'status': 'confirmed',
                'confirmed_lines': ['27 Th9, 2026 06:15:00', 'Test address'],
                'new_timestamp': '28 Th9, 2026 06:15:00'}
    with patch.object(service.ai_service, 'is_configured', return_value=True), \
         patch.object(service.ai_service, 'analyze_watermark_block', return_value=analysis), \
         patch.object(service.ai_service, 'generate_verified_watermark_crop', return_value={
             'status': 'completed', 'image_bytes': raw, 'provider': 'test',
             'model': 'test', 'attempts': 1}) as generate:
        item = service.create_supplement_photo(raw, 'Site', 'Employee', '2026-09-28',
                    'employee.jpg', target_time=target_time, **ids)
    generate.assert_called_once()
    assert generate.call_args.args[0] == raw
    assert generate.call_args.args[1]['confirmed_lines'] == ['28 Th9, 2026 06:15:00', 'Test address']
    assert item['target_time'] == target_time
    assert item['processing'] == {'watermark': 'completed', 'exif': 'completed'}
    assert item['inspection']['exif_datetime'] == '2026-09-27 06:15:00'
    assert item['inspection']['date_mismatch'] is True
    assert item['date_status'] == 'mismatch' and item['can_apply'] is False
    assert item['source_sha256'] == hashlib.sha256(raw).hexdigest()
    assert (service.raw_dir / item['original_file_name']).read_bytes() == raw
    path = service.get_staging_image_path(item['id'])
    with Image.open(path) as image:
        assert image.getexif().get_ifd(34665)[36867] == '2026:09:28 06:15:00'
    assert service.apply_to_attendance([item['id']])['success'] is False
    assert service.delete_staging_photos([item['id']]) == 1
    assert service.get_staging_image_path(item['id']) is None
