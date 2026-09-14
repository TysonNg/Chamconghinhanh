import io
from concurrent.futures import ThreadPoolExecutor

import pytest
from flask import Flask
from PIL import Image

from src.supplement_batches import register_batches


def client(root):
    app = Flask(__name__)
    app.testing = True
    register_batches(app, root)
    return app.test_client()


def create(c, count=1):
    raw = io.BytesIO()
    Image.new('RGB', (16, 16), 'blue').save(raw, 'PNG')
    return c.post('/api/supplement/batches', data={
        'employee': 'A', 'supplement_date': '2026-08-29',
        'photos': [(io.BytesIO(raw.getvalue()), '../../sample.jpg') for _ in range(count)]})


def test_31_images_persist_and_archive_paths_are_generated(tmp_path):
    c = client(tmp_path)
    response = create(c, 31)
    assert response.status_code == 201
    b = response.json['batch']
    assert len({i['archive_name'] for i in b['items']}) == 31
    assert all('/' not in i['archive_name'] and i['archive_name'].endswith('.png') for i in b['items'])
    assert client(tmp_path).get('/api/supplement/batches/' + b['id']).json['batch'] == b
    assert create(c, 32).status_code == 400


@pytest.mark.parametrize('payload', [[], 'bad', {'supplement_date': None}, {'supplement_date': '2026-02-30'}])
def test_invalid_json_dates_are_client_errors(tmp_path, payload):
    c = client(tmp_path)
    b = create(c).json['batch']
    assert c.patch(f"/api/supplement/batches/{b['id']}/items/{b['items'][0]['id']}", json=payload).status_code == 400


def test_invalid_image_and_tamper_block_approval(tmp_path):
    c = client(tmp_path)
    assert c.post('/api/supplement/batches', data={
        'employee': 'A', 'supplement_date': '2026-08-29',
        'photos': (io.BytesIO(b'not an image'), 'fake.jpg')}).status_code == 400
    b = create(c).json['batch']
    url = '/api/supplement/batches/' + b['id']
    assert c.post(url + '/approve', json={}).status_code == 400
    assert c.post(url + '/approve', json={'reviewer': 'A'}).status_code == 200
    path = tmp_path / 'supplement_output' / b['id'] / b['items'][0]['archive_name']
    path.write_bytes(b'tampered')
    assert c.post(url + '/approve', json={'reviewer': 'A'}).status_code == 409
    assert c.get(url + '/download').status_code == 409


def test_concurrent_updates_keep_both_audit_entries(tmp_path):
    c = client(tmp_path)
    b = create(c, 2).json['batch']
    def update(item):
        return c.application.test_client().patch(
            f"/api/supplement/batches/{b['id']}/items/{item['id']}",
            json={'supplement_date': '2026-08-30', 'note': item['id']}).status_code
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert list(pool.map(update, b['items'])) == [200, 200]
    loaded = c.get('/api/supplement/batches/' + b['id']).json['batch']
    assert len(loaded['history']) == 3
    assert all(i['note'] == i['id'] for i in loaded['items'])
