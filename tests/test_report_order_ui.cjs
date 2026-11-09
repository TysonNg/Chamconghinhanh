const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const elements = Object.fromEntries(['report-order-list', 'report-order-project', 'report-order-state', 'report-order-save'].map(id => [id, {value: '', innerHTML: ''}]));
const people = ['a', 'b', 'c'].map((id, index) => ({employee_id: id, display_name: index === 0 ? '<img onerror=x>' : id,
  memberships: [{payroll_code: id}], active: true}));
let response = {success: true, employees: people};
let posted;
let confirm = false;
const context = vm.createContext({console, encodeURIComponent,
  document: {getElementById: id => elements[id], querySelector: () => null, addEventListener() {}},
  escapeHtml: value => String(value).replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;'),
  apiGet: async () => response,
  apiPost: async (url, body) => {posted = {url, body}; return response;},
  showToast() {}, confirmAction: async () => confirm,
});
vm.runInContext(fs.readFileSync('static/js/report-order.js', 'utf8'), context);
(async () => {
  elements['report-order-project'].value = 'p1';
  await context.loadReportOrder();
  assert.match(elements['report-order-list'].innerHTML, /&lt;img onerror=x&gt;/);
  context.moveReportEmployee(2, 0);
  assert.equal(elements['report-order-save'].disabled, false);
  response = {success: true, project_id: 'p1', employee_id: 'b', scan_paused: true};
  await context.setEmployeeScanPaused('b', true, 'p1', {});
  assert.equal(posted.url, '/api/portraits/employee/scan-pause');
  assert.equal(posted.body.paused, true);
  assert.match(elements['report-order-list'].innerHTML, /Tạm nghỉ · Bỏ qua quét và báo cáo/);
  assert.equal(elements['report-order-save'].disabled, false, 'Toggling leave must preserve the unsaved order');
  response = {success: false, error: 'offline'};
  await context.setEmployeeScanPaused('b', false, 'p1', {});
  assert.match(elements['report-order-list'].innerHTML, /Tạm nghỉ · Bỏ qua quét và báo cáo/, 'Failed resume keeps the saved leave status');
  response = {success: true, project_id: 'p1', employee_id: 'b', scan_paused: false};
  await context.setEmployeeScanPaused('b', false, 'p1', {});
  assert(!elements['report-order-list'].innerHTML.includes('Tạm nghỉ · Bỏ qua quét và báo cáo'));
  // Switching projects with an unsaved edit must retain the draft when cancelled.
  elements['report-order-project'].value = 'p2';
  await context.loadReportOrder();
  assert.equal(elements['report-order-project'].value, 'p1');
  response = {success: false, error: 'offline'};
  await context.saveReportOrder();
  assert.deepEqual(Array.from(posted.body.employee_ids), ['c', 'a', 'b']);
  assert.equal(posted.body.project_id, 'p1');
  assert.equal(elements['report-order-save'].disabled, false, 'Failed save preserves the draft for retry');
  response = {success: true, employees: [people[2], people[0], people[1]]};
  await context.saveReportOrder();
  assert.equal(elements['report-order-save'].disabled, true);
  assert.match(elements['report-order-state'].textContent, /Đã lưu/);
  context.moveReportEmployee(-1, 1);
  context.moveReportEmployee(0, 99);
  assert.equal(elements['report-order-save'].disabled, true);
  // A delayed response from a previous project cannot replace the current list.
  let resolveOld;
  context.apiGet = url => url.includes('p1') ? new Promise(resolve => {resolveOld = resolve;}) : Promise.resolve({success: true, employees: [{...people[1], display_name: 'Project two'}]});
  elements['report-order-project'].value = 'p1';
  const oldLoad = context.loadReportOrder();
  elements['report-order-project'].value = 'p2';
  await context.loadReportOrder();
  resolveOld({success: true, employees: people});
  await oldLoad;
  assert.match(elements['report-order-list'].innerHTML, /Project two/);
  assert(!elements['report-order-list'].innerHTML.includes('onerror'));
  console.log('Report order and employee leave: moving, project scoping, drafts, pause/resume, retry, escaping and stale responses passed.');
})().catch(error => {console.error(error); process.exitCode = 1;});
