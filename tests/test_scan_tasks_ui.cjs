const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync('static/js/scan-tasks.js', 'utf8');
const elements = Object.fromEntries(['scan-task-active', 'scan-task-history', 'scan-task-summary', 'scan-task-error'].map(id => [id, {}]));
const escapeHtml = value => String(value || '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;'}[c]));
let response = {success: true, counts: {running: 1, queued: 2, cancelling: 0}, max_running: 1,
    active: [{task_id: 'a', task_type: 'excel_face', status: 'running', project_name: '<img onerror=x>', progress: 2, total: 4}], history: []};
let posted;
const notifications = [];
const bodyClasses = new Set();
const drawerEvents = {};
elements['scan-status-button'] = {dataset: {}, setAttribute(name, value) {this[name] = value;}};
elements['scan-status-text'] = {};
elements['scan-task-history-details'] = {};
elements['scan-drawer'] = {open: false, showModal() {this.open = true;},
    close() {this.open = false; drawerEvents.close?.();},
    addEventListener(name, callback) {drawerEvents[name] = callback;}};
const context = {Set, Date, Number, Math, Error, encodeURIComponent, escapeHtml,
    document: {getElementById: id => elements[id], querySelectorAll: () => [], addEventListener() {},
        body: {classList: {add: name => bodyClasses.add(name), remove: name => bodyClasses.delete(name)}}},
    apiGet: async () => response, apiPost: async url => {posted = url; return {success: true, status: 'cancelled'};},
    showToast(...args) {notifications.push(args);}};
vm.createContext(context);
vm.runInContext(source, context);
(async () => {
    assert.match(context.scanQueueMessage({ahead_count: 0}), /chuẩn bị/);
    assert.match(context.scanQueueMessage({ahead_count: 3}), /3 lượt/);
    await context.refreshScanTasks();
    assert.match(elements['scan-task-summary'].textContent, /1 đang quét.*2 đang chờ/);
    assert.match(elements['scan-task-active'].innerHTML, /50%/);
    assert(!elements['scan-task-active'].innerHTML.includes('<img'));
    assert.match(elements['scan-task-active'].innerHTML, /Dừng quét/);
    assert.equal(elements['scan-status-text'].textContent, 'Quét mặt · Đang quét 50% · 2 chờ');
    assert.equal(notifications.length, 0, 'Initial state must not replay old notifications');
    context.renderScanStatus({counts: {running: 1, cancelling: 0, queued: 0},
        active: [{status: 'running', progress: 0, total: 0}]});
    assert.equal(elements['scan-status-text'].textContent, 'Quét mặt · Đang quét 0%');
    assert.match(context.scanTaskRow({task_id: 'starting', status: 'running', total: 0, progress: 0}), /0%/);
    context.renderScanStatus({counts: {running: 1, cancelling: 0, queued: 0},
        active: [{status: 'running', progress: 0, total: 14}]});
    assert.equal(elements['scan-status-text'].textContent, 'Quét mặt · Đang quét 0%');
    context.initializeScanDrawer();
    context.openScanDrawer();
    assert.equal(elements['scan-drawer'].open, true);
    assert.equal(elements['scan-status-button']['aria-expanded'], 'true');
    assert(bodyClasses.has('scan-drawer-open'));
    drawerEvents.click({target: {}});
    assert.equal(elements['scan-drawer'].open, true, 'Clicks inside the panel keep it open');
    drawerEvents.click({target: elements['scan-drawer']});
    assert.equal(elements['scan-drawer'].open, false);
    assert.equal(elements['scan-status-button']['aria-expanded'], 'false');
    assert(!bodyClasses.has('scan-drawer-open'));
    const activityRow = context.scanTaskRow({task_id: 'a', current: 'Last done',
        active_workers: [{person: '<active>', day: '2026-10-03', image: '002.jpg', image_index: 2, image_total: 20, stage: 'Reading', idle_seconds: 40}]});
    assert.match(activityRow, /Đã xong gần nhất: Last done/);
    assert.match(activityRow, /&lt;active&gt;/);
    assert.match(activityRow, /002.jpg/);
    assert.match(activityRow, /40 giây/);
    assert.match(context.scanTaskRow({task_id: 'b', status: 'queued', ahead_count: 1}), /Hủy chờ/);
    assert.match(context.scanTaskRow({task_id: 'b', status: 'failed'}), /Quét lại/);
    assert(!context.scanTaskRow({task_id: 'b', status: 'cancelling'}).includes('data-scan-action'));
    await context.manageScanTask('cancel', 'a/b');
    assert.equal(posted, '/api/task/cancel/a%2Fb');
    response = {success: false, error: 'offline'};
    await context.refreshScanTasks();
    assert.equal(elements['scan-task-error'].hidden, false);
    assert.match(elements['scan-task-error'].textContent, /có thể đã cũ/);
    assert.equal(elements['scan-status-button'].dataset.state, 'error');
    assert.match(elements['scan-status-text'].textContent, /Mất kết nối/);
    response = {success: true, counts: {running: 0, queued: 0, cancelling: 0}, max_running: 1,
        active: [], history: [{task_id: 'a', status: 'completed', project_name: 'Site'}]};
    await context.refreshScanTasks();
    assert.equal(elements['scan-status-text'].textContent, 'Quét mặt · Đang nghỉ');
    assert.equal(elements['scan-status-button'].dataset.state, 'idle');
    assert.equal(elements['scan-task-error'].hidden, true);
    const completion = notifications.filter(n => n[2] === 'Xem chi tiết');
    assert.equal(completion.length, 1);
    assert.equal(completion[0][1], 'success');
    assert.equal(elements['scan-drawer'].open, false, 'Completion must not force the drawer open');
    await context.refreshScanTasks();
    assert.equal(notifications.filter(n => n[2] === 'Xem chi tiết').length, 1, 'No repeated completion toast');
    completion[0][3]();
    assert.equal(elements['scan-drawer'].open, true);
    assert.equal(elements['scan-task-history-details'].open, true);
    context.closeScanDrawer();
    context.renderScanStatus({counts: {running: 0, cancelling: 1, queued: 2}, active: []});
    assert.match(elements['scan-status-text'].textContent, /Đang dừng · 2 chờ/);
    context.renderScanStatus({counts: {running: 0, cancelling: 0, queued: 2}, active: []});
    assert.match(elements['scan-status-text'].textContent, /Đang chờ · 2 chờ/);
    context.notifyScanTaskChanges({active: [{task_id: 'b', status: 'running'}], history: []});
    context.notifyScanTaskChanges({active: [], history: [{task_id: 'b', status: 'failed'}]});
    assert.equal(notifications.at(-1)[1], 'error');
    elements['scan-ai-mode'] = {value:'', dataset:{}};
    elements['scan-ai-save'] = {};
    elements['scan-ai-hardware'] = {};
    elements['scan-ai-state'] = {};
    const config = {hardware_max:2, max_workers:2, mode:'auto', pending_mode:null,
        cpu_count:4, total_ram_gb:7.9, available_ram_gb:1.4, threads_per_worker:1,
        active:1, waiting:2, memory_waiting:1};
    context.renderScanAIConfig(config);
    assert.match(elements['scan-ai-mode'].innerHTML,/tối đa 2 AI/);
    assert.match(elements['scan-ai-state'].textContent,/1\/2/);
    assert.match(elements['scan-ai-state'].textContent,/chờ đủ RAM/);
    elements['scan-ai-mode'].value='1';
    context.renderScanAIConfig(config);
    assert.equal(elements['scan-ai-mode'].value,'1','Polling must not overwrite an unsaved choice');
    context.apiPost = async(url,data)=>{posted=url; assert.equal(data.mode,1); return {success:true,inference:{...config,pending_mode:1,pending_max_workers:1}};};
    await context.saveScanAIConfig();
    assert.equal(posted,'/api/tasks/ai-config');
    assert.equal(elements['scan-ai-save'].disabled,false);
    assert.match(elements['scan-ai-state'].textContent,/Sẽ đổi sang 1 AI/);
    console.log('Scan monitor rendering, queue messages, cancellation, escaping and stale data passed.');
})().catch(error => {console.error(error); process.exitCode = 1;});
