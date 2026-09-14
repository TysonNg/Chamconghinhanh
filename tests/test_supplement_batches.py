import io
import json
import zipfile
from PIL import Image
from flask import Flask
from src.supplement_batches import register_batches

def test_review_export_preserves_bytes_and_edit_revokes_approval(tmp_path):
    app = Flask(__name__)
    register_batches(app, tmp_path / 'data')
    c = app.test_client()
    image = io.BytesIO()
    exif = Image.Exif()
    exif[306] = '2020:01:02 03:04:05'
    Image.new('RGB', (32, 32), 'red').save(image, format='JPEG', exif=exif)
    original = image.getvalue()
    r = c.post('/api/supplement/batches', data={'employee': 'A', 'supplement_date': '2026-08-29', 'photos': (io.BytesIO(original), 'a.jpg')})
    b = r.get_json()['batch']
    url = '/api/supplement/batches/' + b['id']
    assert c.get(url + '/download').status_code == 409
    assert c.post(url + '/approve', json={'reviewer': 'A'}).status_code == 200
    archive = zipfile.ZipFile(io.BytesIO(c.get(url + '/download').data))
    manifest = json.loads(archive.read('manifest.json'))
    assert archive.read(manifest['items'][0]['archive_name']) == original
    assert manifest['items'][0]['supplement_date'] == '2026-08-29'
    assert c.patch(url + '/items/' + b['items'][0]['id'], json={'supplement_date': '2026-08-30', 'note': 'Bổ sung'}).status_code == 200
    assert c.get(url + '/download').status_code == 409
    assert c.get(url).get_json()['batch']['history'][-1]['action'] == 'item_updated'
