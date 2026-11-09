const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync('static/js/scan-logs.js', 'utf8');
const elements = {};
const listeners = {};
const storage = new Map();
const streams = [];
const escapeHtml = value => String(value || '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;'}[c]));
for (const id of ['scan-log-output','scan-log-task','scan-log-level','scan-log-search','scan-log-count',
    'scan-log-result','scan-log-autoscroll','scan-log-pause','scan-log-clear','scan-log-download','scan-log-connection']) {
    elements[id] = {value: '', checked: true, dataset: {}, options: [], innerHTML: '',
        addEventListener(type, callback) {listeners[`${id}:${type}`] = callback;}};
}
const context = {escapeHtml, console, setTimeout,
    localStorage: {getItem: key => storage.get(key), setItem: (key, value) => storage.set(key,value)},
    document: {getElementById: id => elements[id], addEventListener() {}},
    addEventListener() {},
    EventSource: function(url) {this.url=url; this.close=()=>{}; streams.push(this);}};
vm.createContext(context);
vm.runInContext(source, context);
const store = new context.ScanLogStore();
const task = {task_id: 'scan-1', project_name: 'Long An', target_folder: 'month', status: 'running', progress:19, total:79,
    current:'Last done', active_workers:[{person:'Still scanning', employee_code:'001', day:'2026-10-03',
        image:'002.jpg', image_index:2, image_total:33, stage:'Đang nhận diện ảnh', updated_at:'2026-10-16T16:00:00', idle_seconds:5}]};
store.observe({active:[task], history:[]});
const count = store.entries.length;
store.observe({active:[task], history:[]});
assert.equal(store.entries.length,count,'Identical polls must not duplicate log lines');
task.active_workers[0].idle_seconds=95;
store.observe({active:[task], history:[]});
store.observe({active:[task], history:[]});
assert.equal(store.filtered({level:'warning'}).length,1,'Warn once for a slow inference');
assert.equal(store.filtered({query:'002.jpg', task:'scan-1'}).length,2);
assert.equal(store.filtered({query:'Long An'}).length,store.entries.length);
store.add('Doanminhtriet đang quét', 'info', task, '2020-01-01T00:00:00Z');
assert.equal(store.filtered({query:'Đoàn Minh Triết'}).length,1);
assert.equal(store.filtered()[0].message,'Doanminhtriet đang quét','Display events in timestamp order');
store.observe({active:[],history:[{...task,status:'failed',errors:['timeout']} ]});
assert.equal(store.filtered({level:'error'})[0].message.includes('timeout'),true);
for(let i=0;i<1010;i++) store.add(`line ${i}`);
assert.equal(store.entries.length,1000);
assert.equal(new context.ScanLogStore([null, {}, {time:'x',message:'valid'}]).entries.length,1);

context.startScanJournal();
context.startScanJournal();
assert.equal(streams.length,1,'Initialization must not duplicate the SSE connection');
assert.equal(streams[0].url,'/api/log-stream');
streams[0].onopen();
assert.match(elements['scan-log-connection'].textContent,/Đã kết nối/);
streams[0].onmessage({data:JSON.stringify({type:'heartbeat',message:''})});
streams[0].onmessage({data:JSON.stringify({type:'error',message:'<script>alert(1)</script>'})});
assert(!elements['scan-log-output'].innerHTML.includes('<script>'));
assert.match(elements['scan-log-output'].innerHTML,/&lt;script&gt;/);
const beforePause = elements['scan-log-output'].innerHTML;
listeners['scan-log-pause:click']();
streams[0].onmessage({data:JSON.stringify({type:'info',message:'New while paused'})});
assert.equal(elements['scan-log-output'].innerHTML,beforePause);
assert.match(elements['scan-log-count'].textContent,/vẫn thu nhật ký/);
listeners['scan-log-pause:click']();
assert.match(elements['scan-log-output'].innerHTML,/New while paused/);
assert.match([...storage.values()][0],/New while paused/);
streams[0].onerror();
assert.match(elements['scan-log-connection'].textContent,/kết nối lại/);
listeners['scan-log-clear:click']();
assert.equal([...storage.values()][0],'[]');
console.log('Scan journal deduplication, transitions, bounded storage, filtering, SSE, pause/resume and escaping passed.');

const results = new context.ScanLogStore();
results.add('[matched] Nguyễn An', 'success', task);
results.add('[no_match] Nguyễn Bình', 'warning', task);
results.add('[no_face] Ảnh trống', 'warning');
results.add('Lỗi xuất báo cáo', 'error');
results.add('Hoàn thành', 'success');
assert.equal(results.filtered({result:'matched'}).length, 1);
assert.equal(results.filtered({result:'matched', query:'Bình'}).length, 0);
assert.equal(results.filtered({result:'no_match', task:'scan-1', level:'warning'}).length, 1);
assert.equal(results.filtered({result:'no_face'}).length, 1);
assert.equal(results.filtered({result:'error'}).length, 1);
assert.equal(results.filtered({result:'other'}).length, 1);
elements['scan-log-result'].value = 'matched';
streams[0].onmessage({data:JSON.stringify({type:'success',message:'[matched] Mặt khớp'})});
streams[0].onmessage({data:JSON.stringify({type:'warning',message:'[no_match] Mặt không khớp'})});
listeners['scan-log-result:input']();
assert.match(elements['scan-log-output'].innerHTML, /Mặt khớp/);
assert(!elements['scan-log-output'].innerHTML.includes('Mặt không khớp'));
