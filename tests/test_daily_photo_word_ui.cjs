const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
const start = source.indexOf('async function exportDailyPhotosWord(');
const end = source.indexOf('async function handleDailyPhotosUpload(', start);

function setup(response) {
  const calls = [], messages = [];
  const link = {click() {calls.push(['click', this.download]);}, remove() {calls.push(['remove']);}};
  const context = vm.createContext({
    currentProjectName: 'Dự án A', selectedPhotoDate: () => '2026-09-18',
    URLSearchParams, decodeURIComponent, Uint8Array,
    withLoading: async (btn, fn) => {btn.disabled = true; try {await fn();} finally {btn.disabled = false;}},
    fetch: async url => {calls.push(['fetch', url]); return response;},
    parseJsonResponse: async () => ({error: 'Ngày đang chọn không có ảnh để xuất Word'}),
    URL: {createObjectURL: () => 'blob:report', revokeObjectURL: url => calls.push(['revoke', url])},
    document: {createElement: () => link, body: {appendChild() {}}},
    setTimeout: fn => fn(), showToast: (text, type) => messages.push({text, type}),
  });
  vm.runInContext(source.slice(start, end), context);
  return {context, calls, messages};
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
