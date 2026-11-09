const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
const start = source.indexOf('async function exportDailyPhotosWord(');
const end = source.indexOf('async function handleDailyPhotosUpload(', start);

function setup(response, elements = null) {
  const calls = [], messages = [];
  const timers = [];
  const link = {click() {calls.push(['click', this.download]);}, remove() {calls.push(['remove']);}};
  const context = vm.createContext({
    currentProjectName: 'Dự án A', selectedPhotoDate: () => '2026-09-18', selectedPhotoPeriod: () => '2026-09',
    URLSearchParams, decodeURIComponent, Uint8Array,
    withLoading: async (btn, fn) => {btn.disabled = true; try {await fn();} finally {btn.disabled = false;}},
    fetch: async url => {calls.push(['fetch', url]); return response;},
    parseJsonResponse: async () => ({error: 'Ngày đang chọn không có ảnh để xuất Word'}),
    URL: {createObjectURL: () => 'blob:report', revokeObjectURL: url => calls.push(['revoke', url])},
    document: {createElement: () => link, body: {appendChild() {}}, getElementById: id => elements?.[id]},
    setTimeout: fn => elements ? timers.push(fn) : fn(), clearTimeout: id => calls.push(['clearTimer', id]),
    showToast: (text, type) => messages.push({text, type}),
  });
  vm.runInContext(source.slice(start, end), context);
  return {context, calls, messages, timers};
}

test('downloads Word for captured project/date with server Unicode filename', async () => {
  const filename = 'Anh_Dự án A_2026-09-18.docx';
  const {context, calls, messages} = setup({ok: true, blob: async () => new Blob([new Uint8Array([0x50, 0x4b, 3, 4])]),
    headers: {get: key => key === 'Content-Type'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`}});
  const btn = {disabled: false};
  await context.exportDailyPhotosWord(btn);
  const query = new URL(calls[0][1], 'http://localhost').searchParams;
  assert.equal(query.get('project'), 'Dự án A');
  assert.equal(query.get('date'), '2026-09-18');
  assert.deepEqual(calls[1], ['click', filename]);
  assert.ok(calls.some(c => c[0] === 'revoke'));
  assert.equal(messages[0].type, 'success');
  assert.equal(btn.disabled, false);
});

test('server error shows message and restores button without downloading', async () => {
  const {context, calls, messages} = setup({ok: false});
  const btn = {disabled: false};
  await context.exportDailyPhotosWord(btn);
  assert.equal(messages[0].type, 'error');
  assert.match(messages[0].text, /không có ảnh/);
  assert.equal(calls.length, 1);
  assert.equal(btn.disabled, false);
});

test('monthly button downloads a ZIP of daily Word files for the selected project and month', async () => {
  const filename = 'Anh_Dự án A_2026-09.zip';
  const {context, calls, messages} = setup({ok: true,
    blob: async () => new Blob([new Uint8Array([0x50, 0x4b, 3, 4])]),
    headers: {get: key => key === 'Content-Type'
      ? 'application/zip'
      : `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`}});
  const btn = {disabled: false};
  await context.exportMonthlyPhotosWord(btn);
  const url = new URL(calls[0][1], 'http://localhost');
  assert.equal(url.pathname, '/api/photos/daily/export-month-word');
  assert.equal(url.searchParams.get('project'), 'Dự án A');
  assert.equal(url.searchParams.get('period'), '2026-09');
  assert.equal(url.searchParams.has('date'), false);
  assert.deepEqual(calls[1], ['click', filename]);
  assert.equal(messages[0].type, 'success');
  assert.equal(btn.disabled, false);
});

test('monthly export errors restore button without downloading', async () => {
  const {context, calls, messages} = setup({ok: false});
  const btn = {disabled: false};
  await context.exportMonthlyPhotosWord(btn);
  assert.equal(calls.length, 1);
  assert.equal(messages[0].type, 'error');
  assert.equal(btn.disabled, false);
});

test('old server returning a combined monthly DOCX asks to restart instead of saving a fake ZIP', async () => {
  const {context, calls, messages} = setup({ok: true,
    headers: {get: () => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},
    blob: async () => new Blob([new Uint8Array([0x50, 0x4b, 3, 4])])});
  await context.exportMonthlyPhotosWord({disabled: false});
  assert.equal(messages[0].type, 'error');
  assert.match(messages[0].text, /khởi động lại/);
  assert(!calls.some(call => call[0] === 'click'));
});

test('monthly export shows live day progress and keeps a manual download link', async () => {
  const elements = Object.fromEntries(['status', 'message', 'progress', 'download'].map(name =>
    [`monthly-word-export-${name}`, {hidden: true, dataset: {}, removeAttribute() {}}]));
  const response = {ok: true, blob: async () => new Blob([new Uint8Array([0x50, 0x4b, 3, 4])]),
    headers: {get: key => key === 'Content-Type' ? 'application/zip' : ''}};
  const {context, calls, timers} = setup(response, elements);
  let finish;
  context.fetch = async url => url.includes('export-progress')
    ? {ok: true, json: async () => ({success: true, status: 'running', completed_days: 4, total_days: 30})}
    : new Promise(resolve => {finish = resolve;});
  const button = {disabled: false, textContent: 'Xuất Word cả tháng'};
  const pending = context.exportMonthlyPhotosWord(button);
  assert.equal(button.disabled, true);
  assert.equal(button.textContent, 'Đang tạo Word…');
  assert.equal(elements['monthly-word-export-status'].hidden, false);
  await timers.shift()();
  assert.match(elements['monthly-word-export-message'].textContent, /4\/30 ngày/);
  assert.equal(elements['monthly-word-export-progress'].value, 4);
  finish(response);
  await pending;
  assert.equal(button.disabled, false);
  assert.equal(button.textContent, 'Xuất Word cả tháng');
  assert.equal(elements['monthly-word-export-download'].hidden, false);
  assert.equal(elements['monthly-word-export-download'].href, 'blob:report');
  assert.match(elements['monthly-word-export-download'].download, /2026-09.zip$/);
  assert.match(elements['monthly-word-export-message'].textContent, /mỗi ngày một file Word/);
  assert.equal(elements['monthly-word-export-status'].dataset.state, 'completed');
  assert(!calls.some(call => call[0] === 'revoke'), 'Keep the blob until the user can manually download it');
  assert(calls.some(call => call[0] === 'clearTimer'));
});

test('monthly failure remains visible after toast closes and restores the button', async () => {
  const elements = Object.fromEntries(['status', 'message', 'progress', 'download'].map(name =>
    [`monthly-word-export-${name}`, {hidden: true, dataset: {}, removeAttribute() {}}]));
  const {context, calls} = setup({ok: false}, elements);
  const button = {disabled: false, textContent: 'Xuất Word cả tháng'};
  await context.exportMonthlyPhotosWord(button);
  assert.equal(elements['monthly-word-export-status'].dataset.state, 'error');
  assert.match(elements['monthly-word-export-message'].textContent, /Không xuất được Word/);
  assert.equal(elements['monthly-word-export-download'].hidden, true);
  assert.equal(button.disabled, false);
  assert.equal(button.textContent, 'Xuất Word cả tháng');
  assert(!calls.some(call => call[0] === 'click'));
});

test('disabled export button ignores repeated click', async () => {
  const {context, calls} = setup({ok: true});
  await context.exportDailyPhotosWord({disabled: true});
  assert.equal(calls.length, 0);
});

test('stale server returning successful gallery JSON cannot download a fake docx', async () => {
  const {context, calls, messages} = setup({ok: true,
    headers: {get: () => 'application/json'},
    blob: async () => new Blob(['{"success":true,"photos":[]}'])});
  const btn = {disabled: false};
  await context.exportDailyPhotosWord(btn);
  assert.equal(calls.length, 1);
  assert.equal(messages[0].type, 'error');
  assert.match(messages[0].text, /khởi động lại/);
  assert.equal(btn.disabled, false);
});

test('invalid payload is rejected even when response claims Word content type', async () => {
  const {context, calls, messages} = setup({ok: true,
    headers: {get: () => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},
    blob: async () => new Blob(['{"success":true}'])});
  await context.exportDailyPhotosWord({disabled: false});
  assert.equal(calls.length, 1);
  assert.match(messages[0].text, /không hợp lệ/);
});
