from pathlib import Path

from flask import Flask, render_template

from src.supplement_batches import register_batches


def test_template_uses_supplement_records_not_timestamp_rewriting(tmp_path):
    app = Flask(__name__, template_folder=str(Path(__file__).resolve().parents[1] / 'templates'))
    register_batches(app, tmp_path)
    with app.test_request_context('/'):
        html = render_template('index.html')
    assert html.count('id="tab-supplement"') == 1
    assert 'id="batch-create"' in html
    assert 'supplement-batches.js' in html
    assert 'id="supp-modify-exif"' not in html
    assert 'id="supplement-upload-form"' not in html
    c = app.test_client()
    for route in ('upload', 'process'):
        assert c.post('/api/supplement/' + route).status_code == 410
