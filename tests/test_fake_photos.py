import io
import json
from pathlib import Path
from PIL import Image
from flask import Flask
from src.supplement_batches import register_batches

def test_fake_photos_flow(tmp_path):
    app = Flask(__name__)
    register_batches(app, tmp_path / 'data')
    c = app.test_client()

    # Create dummy images with matching EXIF dates
    exif1 = Image.Exif()
    exif1[36867] = "2026:08:01 08:15:00"
    img1 = io.BytesIO()
    Image.new('RGB', (100, 100), 'blue').save(img1, format='JPEG', exif=exif1)

    exif2 = Image.Exif()
    exif2[36867] = "2026:08:03 17:30:00"
    img2 = io.BytesIO()
    Image.new('RGB', (100, 100), 'green').save(img2, format='JPEG', exif=exif2)

    configs = [
        {"target_date": "2026-08-01", "target_time": "08:15:00"},
        {"target_date": "2026-08-03", "target_time": "17:30:00"}
    ]

    # Test POST fake-photos
    resp = c.post('/api/supplement/fake-photos', data={
        'project': 'TestProject',
        'employee': 'NguyenVanA',
        'configs': json.dumps(configs),
        'photos': [
            (io.BytesIO(img1.getvalue()), 'p1.jpg'),
            (io.BytesIO(img2.getvalue()), 'p2.jpg')
        ]
    })
    assert resp.status_code == 201
    data = resp.get_json()
    assert data['success'] is True
    assert data['count'] == 2
    items = data['items']
    assert items[0]['target_date'] == '2026-08-01'
    assert items[1]['target_date'] == '2026-08-03'

    # Test GET staging
    resp_stg = c.get('/api/supplement/staging?project=TestProject')
    assert resp_stg.status_code == 200
    stg_items = resp_stg.get_json()['items']
    assert len(stg_items) == 2

    # Test GET staging image
    p0_id = stg_items[0]['id']
    resp_img = c.get(f'/api/supplement/staging/{p0_id}/image')
    assert resp_img.status_code == 200

    # Test apply to attendance
    resp_apply = c.post('/api/supplement/apply-to-attendance', json={
        'ids': [p0_id],
        'delete_after': True
    })
    assert resp_apply.status_code == 200
    apply_data = resp_apply.get_json()
    assert apply_data['count'] == 1

    # Verify p0 was deleted from staging
    resp_stg2 = c.get('/api/supplement/staging?project=TestProject')
    assert len(resp_stg2.get_json()['items']) == 1

    # Verify input_images file exists
    dest_path = Path(apply_data['applied'][0]['dest_path'])
    assert dest_path.exists()
