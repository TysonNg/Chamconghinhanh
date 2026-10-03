const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function harness(count) {
  const nodes = {
    'btn-pick-random-portrait': {}, 'supp-random-count': {value: count},
    'batch-message': {}, 'batch-create': {},
    'btn-submit-fake-photos': {}, 'supp-project-select': {value: 'p1'},
    'supp-employee-val': {value: 'e1'},
  };
  const calls = [];
  const context = vm.createContext({
    document: {getElementById: id => nodes[id] || null},
    FormData, File, URL, console,
    fetch: async (url, options) => {
      calls.push({url, options});
      return {ok: true, json: async () => ({success: true, count: 1})};
    }
  });
  let source = fs.readFileSync(path.join(__dirname, '../static/js/supplement-batches.js'), 'utf8');
  source = source.slice(0, source.indexOf('  // --- INITIALIZE ---')) + `
    currentSelectedEmployee = {employee_id: 'e1', name: 'Employee', images: ['one.jpg']};
    selectedPhotos = [{file: new File(['x'], 'one.jpg'), target_date:'2026-09-16'}];
    setupRandomPortraitPicker(); setupFormSubmit();
  })();`;
  vm.runInContext(source, context);
  return {nodes, calls};
}

test('random picker rejects invalid and unavailable counts before fetching', async () => {
  for (const value of ['', '0', '-1', '1.5', '2', 'abc']) {
    const {nodes, calls} = harness(value);
    await nodes['btn-pick-random-portrait'].onclick();
    assert.match(nodes['batch-message'].textContent, /Kho hiện có 1 ảnh/);
    assert.equal(calls.length, 0);
  }
});

test('saving selected photos ignores random count and sends requested date only', async () => {
  const {nodes, calls} = harness('999');
  await nodes['batch-create'].onsubmit({preventDefault() {}});
  assert.equal(calls[0].url, '/api/supplement/records');
  assert.deepEqual(JSON.parse(calls[0].options.body.get('configs')), [{target_date:'2026-09-16'}]);
  assert.equal(calls[0].options.body.has('replace_timestamp'), false);
  assert.equal(calls[0].options.body.has('modify_exif'), false);
});

test('saving uses stable IDs and never sends a shift', async () => { const {nodes,calls}=harness('1'); await nodes['batch-create'].onsubmit({preventDefault(){}}); const form=calls[0].options.body; assert.equal(form.get('project_id'),'p1'); assert.equal(form.get('employee_id'),'e1'); assert.equal(form.has('shift'),false); });

test('manual file selection and rendering preserve missing time',()=>{
 const nodes={}; const ctx=vm.createContext({document:{getElementById:id=>nodes[id]||null},URL,Blob,console,window:{}});
 let source=fs.readFileSync(path.join(__dirname,'../static/js/supplement-batches.js'),'utf8');
 source=source.slice(0,source.indexOf('  // --- INITIALIZE ---'))+`\n handleFilesSelected([new Blob(['x'])]); window.photos=selectedPhotos; })();`;
 vm.runInContext(source,ctx);assert.equal(ctx.window.photos.length,1);assert.equal(ctx.window.photos[0].target_time,undefined);
});
