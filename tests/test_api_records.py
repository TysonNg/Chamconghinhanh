import hashlib
import io
import json
from supplement_test_support import environment, image_bytes


def test_api_supplement_records(tmp_path):
    service, client, ids = environment(tmp_path)
    raw = image_bytes()
    res = client.post('/api/supplement/records', data={
        **ids, 'configs': json.dumps([{'target_date': '2026-09-28', 'target_time': '05:59:30'}]),
        'replace_timestamp': 'false', 'modify_exif': 'false',
        'photos': (io.BytesIO(raw), 'employee.jpg')})
    assert res.status_code == 201
    data = res.get_json()
    assert data['success'] is True and data['count'] == 1
    item = data['items'][0]
    assert item['target_date'] == '2026-09-28'
    assert item['target_time'] == '05:59:30'
    assert item['inspection']['date_mismatch'] is False
    assert item['inspection']['exif_datetime'] == '2026-09-28 05:59:30'
    assert item['face_status'] == 'matched' and item['can_apply'] is True
    assert item['source_sha256'] == hashlib.sha256(raw).hexdigest()
    assert client.get(item['url']).data == raw
    assert (service.raw_dir / item['original_file_name']).read_bytes() == raw
