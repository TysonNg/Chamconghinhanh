import io
import json
from unittest.mock import patch

from PIL import Image
from flask import Flask

from src.ai_timestamp_service import AITimestampService
from src.fake_photo_service import FakePhotoService
from src.supplement_batches import register_batches


def photo_bytes():
    stream = io.BytesIO()
    exif = Image.Exif()
    exif[306] = '2026:01:11 17:11:51'
    Image.new('RGB', (20, 20), 'white').save(stream, 'JPEG', exif=exif)
    return stream.getvalue()


def test_supplement_preserves_bytes_and_reports_duplicate(tmp_path):
    service = FakePhotoService(str(tmp_path))
    raw = photo_bytes()
    with patch.object(service.ai_service, 'inspect_photo', return_value={'status': 'unconfigured', 'message': 'No key'}):
        first = service.create_supplement_photo(raw, 'Site', 'Employee', '2026-09-16', 'photo.jpg')
        second = service.create_supplement_photo(raw, 'Site', 'Employee', '2026-09-17', 'photo.jpg')
    assert service.get_staging_image_path(first['id']).read_bytes() == raw
    assert first['inspection']['exif_datetime'] == '2026:01:11 17:11:51'
    assert first['inspection']['date_mismatch'] is True
    assert second['inspection']['duplicate_ids'] == [first['id']]
    assert service.get_staging_photos()[0]['inspection']['status'] == 'unconfigured'


def test_inspection_uses_available_provider_and_reports_failure(tmp_path):
    service = AITimestampService(str(tmp_path))
    service.save_config({'provider': 'openai', 'gemini_api_key': 'test-key'})
    with patch('src.ai_timestamp_service.requests.post', side_effect=TimeoutError):
        result = service.inspect_photo(photo_bytes())
    assert result['status'] == 'failed'
    assert result['provider'] == 'gemini'


def test_inspection_prompt_and_date_validation(tmp_path):
    service = AITimestampService(str(tmp_path))
    service.save_config({'gemini_api_key': 'test-key'})
    with patch('src.ai_timestamp_service.requests.post') as post:
        post.return_value.status_code = 200
        post.return_value.json.return_value = {'candidates': [{'content': {'parts': [{'text': '{"visible_text":"11 Th1, 2026", "visible_date":"2026-01-11", "visible_time":null}'}]}}]}
        result = service.inspect_photo(photo_bytes())
    assert result['status'] == 'ok'
    assert result['visible_date'] == '2026-01-11'
    assert '2026-09-16' not in str(post.call_args)


def test_records_endpoint_preserves_bytes_and_validates_all_dates(tmp_path):
    app = Flask(__name__)
    register_batches(app, tmp_path)
    client = app.test_client()
    raw = photo_bytes()
    with patch.object(AITimestampService, 'inspect_photo', return_value={'status': 'failed', 'message': 'AI unavailable'}):
        response = client.post('/api/supplement/records', data={
            'project': 'Site', 'employee': 'Employee',
            'configs': json.dumps([{'target_date': '2026-09-16'}]),
            'photos': (io.BytesIO(raw), 'original.jpg')})
    assert response.status_code == 201
    item = response.json['items'][0]
    assert item['target_date'] == '2026-09-16'
    assert item['target_time'] == ''
    assert item['inspection']['status'] == 'failed'
    assert client.get(item['url']).data == raw
    assert client.get('/api/supplement/download-staging/' + item['id']).data == raw
    response = client.post('/api/supplement/records', data={
        'project': 'Site', 'employee': 'Employee',
        'configs': json.dumps([{'target_date': '2026-09-16'}, {'target_date': 'invalid'}]),
        'photos': [(io.BytesIO(raw), 'one.jpg'), (io.BytesIO(raw), 'two.jpg')]})
    assert response.status_code == 400
    assert client.get('/api/supplement/staging').json['count'] == 1


def test_style_prompt_with_json_braces_does_not_fail(tmp_path):
    import numpy as np
    service = AITimestampService(str(tmp_path))
    service.save_config({'gemini_api_key': 'test-key'})
    with patch.object(service, '_call_gemini_vision', return_value={'text_color': '#FFFFFF'}) as call:
        result = service.analyze_timestamp_style(np.zeros((10, 10, 3), dtype=np.uint8), 'sample', '2026-09-16')
    assert result['text_color'] == '#FFFFFF'
    prompt = call.call_args.args[3]
    assert '"text_color"' in prompt
    assert '2026-09-16' in prompt


def test_invalid_ai_date_is_explicit_failure(tmp_path):
    service = AITimestampService(str(tmp_path))
    service.save_config({'gemini_api_key': 'test-key'})
    with patch('src.ai_timestamp_service.requests.post') as post:
        post.return_value.status_code = 200
        post.return_value.json.return_value = {'candidates': [{'content': {'parts': [{'text': '{"visible_text":"unclear", "visible_date":"2026-02-31"}'}]}}]}
        result = service.inspect_photo(photo_bytes())
    assert result['status'] == 'failed'


def test_provider_fallback_keeps_original_failure_visible(tmp_path):
    from unittest.mock import Mock
    service = AITimestampService(str(tmp_path))
    service.save_config({'gemini_api_key': 'first-key', 'openai_api_key': 'second-key'})
    failed = Mock(status_code=503)
    success = Mock(status_code=200)
    success.json.return_value = {'choices': [{'message': {'content': '{"visible_text":"2026-01-11", "visible_date":"2026-01-11"}'}}]}
    with patch('src.ai_timestamp_service.requests.post', side_effect=[failed, success]):
        result = service.inspect_photo(photo_bytes())
    assert result['provider'] == 'openai'
    assert result['status'] == 'ok'
    assert result['provider_failures'] == ['gemini: HTTP 503']
