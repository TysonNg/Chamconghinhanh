const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function harness(count) {
  const nodes = {
    'btn-pick-random-portrait': {}, 'supp-random-count': {value: count},
    'batch-message': {}, 'batch-create': {},
    'btn-submit-fake-photos': {}, 'supp-project-select': {value: 'Site'},
    'supp-employee-val': {value: 'Employee'},
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
    currentSelectedEmployee = {name: 'Employee', images: ['one.jpg']};
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
