const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

test('Zalo streamlined UI removes big card and provides modal + smart alert', () => {
    const htmlPath = path.resolve(__dirname, '..', 'templates', 'index.html');
    const html = fs.readFileSync(htmlPath, 'utf8');

    // 1. Verify old bulky card is gone
    assert.ok(!html.includes('id="zalo-auto-sync-card"'), 'Old bulky #zalo-auto-sync-card must be removed');

    // 2. Verify modal button in header ribbon
    assert.ok(html.includes('id="btn-zalo-auto-sync-modal"'), '#btn-zalo-auto-sync-modal must be in header');
    assert.ok(html.includes('openZaloAutoSyncModal()'), 'Header button must invoke openZaloAutoSyncModal()');
    assert.ok(html.includes('id="zalo-header-sync-status-badge"'), '#zalo-header-sync-status-badge must exist in header');

    // 3. Verify smart alert in left column
    assert.ok(html.includes('id="zalo-smart-gap-alert"'), '#zalo-smart-gap-alert must exist');
    assert.ok(html.includes('id="zalo-smart-gap-ok"'), '#zalo-smart-gap-ok must exist');
    assert.ok(html.includes('id="btn-zalo-backfill-now"'), '#btn-zalo-backfill-now must exist inside smart gap alert');
    assert.ok(html.includes('id="zalo-timeline-details-collapse"'), 'Collapsible timeline details container must exist');

    // 4. Verify modal exists
    assert.ok(html.includes('id="modal-zalo-auto-sync"'), '#modal-zalo-auto-sync must exist');
    assert.ok(html.includes('id="modal-zalo-sync-enabled"'), '#modal-zalo-sync-enabled must exist');
    assert.ok(html.includes('id="modal-zalo-sync-time"'), '#modal-zalo-sync-time must exist');
    assert.ok(html.includes('id="modal-zalo-sync-lookback"'), '#modal-zalo-sync-lookback must exist');
    assert.ok(html.includes('id="modal-zalo-sync-startup"'), '#modal-zalo-sync-startup must exist');
    assert.ok(html.includes('modal-header-btn-close'), 'Modal must have a dedicated close button with modal-header-btn-close class');
    assert.ok(html.includes('clearAllZaloAutoSyncMappings()'), 'Modal must have clear all button');
});

test('app.js defines necessary functions for streamlined Zalo UI', () => {
    const jsPath = path.resolve(__dirname, '..', 'static', 'js', 'app.js');
    const js = fs.readFileSync(jsPath, 'utf8');

    assert.ok(js.includes('function openZaloAutoSyncModal('), 'openZaloAutoSyncModal must be defined');
    assert.ok(js.includes('function closeZaloAutoSyncModal('), 'closeZaloAutoSyncModal must be defined');
    assert.ok(js.includes('function saveZaloSyncModalSettings('), 'saveZaloSyncModalSettings must be defined');
    assert.ok(js.includes('function toggleZaloTimelineDetails('), 'toggleZaloTimelineDetails must be defined');
    assert.ok(js.includes('function updateZaloHeaderSyncBadge('), 'updateZaloHeaderSyncBadge must be defined');
    assert.ok(js.includes('function renderZaloAutoSyncModalMappings('), 'renderZaloAutoSyncModalMappings must be defined');
    assert.ok(js.includes('function clearAllZaloAutoSyncMappings('), 'clearAllZaloAutoSyncMappings must be defined');
});
