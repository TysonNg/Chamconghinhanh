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
    assert 'Lưu ảnh bổ sung và kiểm tra' in html
    assert 'Ngày bạn chọn là ngày đề nghị bổ sung công' in html
    assert 'id="chk-replace-timestamp"' not in html
    assert 'id="chk-modify-exif"' not in html
    from html.parser import HTMLParser
    class Inputs(HTMLParser):
        def handle_starttag(self, tag, attrs):
            if tag == 'input' and dict(attrs).get('id') == 'supp-random-count':
                self.count_input = dict(attrs)
    parser = Inputs()
    parser.feed(html)
    assert parser.count_input['type'] == 'text'
    assert not any(key in parser.count_input for key in ('required', 'min', 'max', 'pattern'))
    assert 'supplement-batches.js' in html
    assert 'id="supp-modify-exif"' not in html
    assert 'id="supplement-upload-form"' not in html
    c = app.test_client()
    for route in ('upload', 'process'):
        assert c.post('/api/supplement/' + route).status_code == 410


def test_full_watermark_generation_never_requests_confirmation_and_keeps_sources():
    root = Path(__file__).resolve().parents[1]
    html = (root / 'templates' / 'index.html').read_text(encoding='utf-8')
    script = (root / 'static' / 'js' / 'supplement-batches.js').read_text(encoding='utf-8')

    assert 'id="modal-watermark-confirm"' in html
    assert script.count('showWatermarkConfirmation(') == 1
    assert "res.status === 409" not in script
    # Stored uploads report per-image processing and validation rather than unconditional success.
    assert 'item.processing?.watermark' in script
    assert 'item.processing?.exif' in script
    for field in ('date_status', 'face_status', 'integrity_status', 'storage_status', 'can_apply'):
        assert 'item.' + field in script
    assert 'verification_failed' in script
    assert 'supp-staging-status-tag' in script


def test_failed_processing_displays_error_status():
    import subprocess
    root = Path(__file__).resolve().parents[1]
    # Exercise the existing UI in a VM; the Node agent owns all JS and .cjs files.
    program = r"""
const fs = require('node:fs'), vm = require('node:vm'), assert = require('node:assert/strict');
let source = fs.readFileSync('static/js/supplement-batches.js', 'utf8');
source = source.slice(0, source.indexOf('  // --- INITIALIZE ---')) + '\n window.regenerate=regenerateStagingPhoto; })();';
const message = {};
const context = vm.createContext({window:{}, document:{getElementById:id=>id==='batch-message'?message:null}, console,
 fetch:async()=>({ok:true,status:200,json:async()=>({success:true,processing:{watermark:{status:'failed'}},status:'verification_failed'})})});
vm.runInContext(source, context);
(async()=>{assert.equal(await context.window.regenerate('one', {innerHTML:'Generate'}), false);
assert.equal(message.className.includes('error'), true);
assert.equal(message.className.includes('success'), false);
assert.ok(message.textContent);})().catch(error=>{console.error(error);process.exitCode=1;});
"""
    result = subprocess.run(['node', '-e', program], cwd=root, capture_output=True, text=True)
    assert result.returncode == 0, result.stdout + result.stderr
