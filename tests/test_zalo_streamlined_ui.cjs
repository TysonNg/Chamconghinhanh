const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

test('Zalo UI keeps date downloads and removes background sync controls', () => {
    const html = fs.readFileSync(path.resolve(__dirname, '..', 'templates', 'index.html'), 'utf8');
    for (const id of ['btn-zalo-auto-sync-modal', 'modal-zalo-auto-sync', 'zalo-header-sync-status-badge', 'zalo-smart-gap-alert', 'zalo-smart-gap-ok', 'btn-zalo-backfill-now']) {
        assert.ok(!html.includes(`id="${id}"`), `${id} must be removed`);
    }
    assert.ok(!html.includes('addCurrentSelectionToAutoSync()'));
    for (const id of ['zalo-target-project', 'zalo-date-from', 'zalo-date-to', 'btn-cancel-zalo-download']) {
        assert.ok(html.includes(`id="${id}"`), `${id} must remain`);
    }
    assert.ok(html.includes('startZaloDownload()'));
});

test('Zalo frontend no longer initializes background sync', () => {
    const js = fs.readFileSync(path.resolve(__dirname, '..', 'static', 'js', 'app.js'), 'utf8');
    assert.ok(!js.includes('initZaloAutoSync'));
    assert.ok(!js.includes('refreshZaloTimelineGaps'));
    assert.ok(!js.includes('openZaloAutoSyncModal'));
    assert.ok(js.includes('async function startZaloDownload('));
});
