const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('browser progress counts downloaded images once and exposes partial results', () => {
    const source = fs.readFileSync(path.join(__dirname, '../../static/js/app.js'), 'utf8');
    const start = source.indexOf('function updateZaloProgressUI(');
    const end = source.indexOf('// Append log to terminal', start);
    const elements = new Map();
    const document = {getElementById: id => {
        if (!elements.has(id)) elements.set(id, {textContent: '', style: {}, classList: {add() {}, remove() {}}});
        return elements.get(id);
    }};
    const context = vm.createContext({document});
    vm.runInContext(source.slice(start, end), context);
    context.updateZaloProgressUI({counterVersion: 2, status: 'done', total: 10, current: 10, downloaded: 5, skipped: 2, failed: 1, unknownDate: 2, lowQuality: 3});
    assert.equal(elements.get('zalo-stat-downloaded').textContent, 5);
    assert.equal(elements.get('zalo-stat-low-quality').textContent, 3);
    assert.equal(elements.get('zalo-stat-unknown-date').textContent, 2);
    assert.equal(elements.get('zalo-progress-percent').textContent, '100%');
    assert.doesNotMatch(elements.get('zalo-current-file-text').textContent, /tất cả ảnh.*thành công/);
});
