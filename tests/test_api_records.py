import io
import json
import sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask
from src.supplement_batches import register_batches

def test_api_supplement_records(tmp_path):
    app = Flask(__name__)
    register_batches(app, tmp_path)
    client = app.test_client()

    # Read real staging image or generate dummy image
    staging_files = list(Path('supplement_data/supplement_staging').glob('*.jpg'))
    if staging_files:
        raw_bytes = staging_files[0].read_bytes()
    else:
        img = Image.new('RGB', (400, 300), color=(200, 200, 200))
        buf = io.BytesIO()
        img.save(buf, format='JPEG')
        raw_bytes = buf.getvalue()

    res = client.post('/api/supplement/records', data={
        'project': 'Chung cư Tân Thuận Đông',
        'employee': 'Dang Van Tau',
        'shift': 'morning',
        'configs': json.dumps([{
            'target_date': '2026-09-28',
            'target_time': '05:59:30'
        }]),
        'photos': (io.BytesIO(raw_bytes), 'dang_van_tau.jpg')
    })

    print("Status code:", res.status_code)
    data = res.get_json()
    print("Response data:", json.dumps(data, ensure_ascii=False, indent=2))

    assert res.status_code == 201
    assert data['success'] is True
    assert data['count'] == 1
    item = data['items'][0]
    assert item['target_date'] == '2026-09-28'
    assert item['target_time'] == '05:59:30'
    assert item['inspection']['date_mismatch'] is False

    print("\nAPI ENDPOINT TEST PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    from pathlib import Path
    test_api_supplement_records(Path('supplement_data'))
