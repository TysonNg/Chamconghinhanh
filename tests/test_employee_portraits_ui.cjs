const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../static/js/app.js'), 'utf8');

function getFunction(name) {
  const start = source.indexOf('function ' + name + '(');
  const asyncStart = source.indexOf('async function ' + name + '(');
  const actualStart = asyncStart !== -1 && (start === -1 || asyncStart < start) ? asyncStart : start;
  const nextFn = source.indexOf('\nfunction ', actualStart + 1);
  const nextAsyncFn = source.indexOf('\nasync function ', actualStart + 1);
  let end = source.length;
  if (nextFn !== -1 && nextAsyncFn !== -1) end = Math.min(nextFn, nextAsyncFn);
  else if (nextFn !== -1) end = nextFn;
  else if (nextAsyncFn !== -1) end = nextAsyncFn;
  return source.slice(actualStart, end);
}

test('openCurrentEmployeeFolder calls apiPost with active employee and current project', async () => {
  const calls = [];
  const toasts = [];
  const context = vm.createContext({
    currentProjectName: 'CongTyA',
    activeEmpModalName: 'NV001',
    apiPost: async (url, payload) => {
      calls.push({ url, payload });
      return { success: true, message: 'Thư mục đã mở' };
    },
    showToast: (msg, type) => toasts.push({ msg, type })
  });
  vm.runInContext(getFunction('openCurrentEmployeeFolder'), context);
  await context.openCurrentEmployeeFolder();
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, '/api/portraits/open-folder');
  assert.equal(calls[0].payload.project, 'CongTyA');
  assert.equal(calls[0].payload.employee_id, 'NV001');
  assert.equal(toasts[0].type, 'success');
});

test('openProjectPortraitsFolder calls apiPost with current project', async () => {
  const calls = [];
  const toasts = [];
  const context = vm.createContext({
    currentProjectName: 'CongTyA',
    apiPost: async (url, payload) => {
      calls.push({ url, payload });
      return { success: true, message: 'Thư mục đã mở' };
    },
    showToast: (msg, type) => toasts.push({ msg, type })
  });
  vm.runInContext(getFunction('openProjectPortraitsFolder'), context);
  await context.openProjectPortraitsFolder();
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, '/api/portraits/open-folder');
  assert.equal(calls[0].payload.project, 'CongTyA');
  assert.equal(calls[0].payload.employee_id, undefined);
  assert.equal(toasts[0].type, 'success');
});

test('extractImagesFromClipboard extracts image files and items', () => {
  const context = vm.createContext({
    File: class MockFile {
      constructor(chunks, name, opts) {
        this.name = name;
        this.type = opts?.type || '';
      }
    }
  });
  vm.runInContext(getFunction('extractImagesFromClipboard'), context);

  // 1. Test clipboard with files
  const file1 = { name: 'photo1.jpg', type: 'image/jpeg' };
  const file2 = { name: 'document.pdf', type: 'application/pdf' };
  const resFiles = context.extractImagesFromClipboard({ files: [file1, file2], items: [] });
  assert.equal(resFiles.length, 1);
  assert.equal(resFiles[0].name, 'photo1.jpg');

  // 2. Test clipboard with items (e.g. Zalo paste)
  const fakeBlob = { type: 'image/png' };
  const item1 = {
    type: 'image/png',
    getAsFile: () => fakeBlob
  };
  const item2 = {
    type: 'text/plain',
    getAsFile: () => null
  };
  const resItems = context.extractImagesFromClipboard({ files: [], items: [item1, item2] });
  assert.equal(resItems.length, 1);
  assert.match(resItems[0].name, /^zalo_paste_\d+_1\.png$/);
  assert.equal(resItems[0].type, 'image/png');
});

test('handleUploadEmployeePhoto uploads multiple files from drag-drop or paste', async () => {
  const calls = [];
  const context = vm.createContext({
    activeEmpModalName: 'NV002',
    currentProjectName: 'CongTyB',
    allProjectsList: [{ name: 'CongTyB', project_id: 'p2' }],
    FormData,
    File,
    Blob,
    showToast() {},
    loadPortraits: async () => {},
    refreshEmployeePhotosModal() {},
    loadProjects: async () => {},
    fetch: async (url, options) => {
      calls.push({ url, options });
      return { json: async () => ({ success: true, saved_count: 2 }) };
    }
  });
  const fnStart = source.indexOf('async function handleUploadEmployeePhoto(');
  const fnEnd = source.indexOf('async function deleteEmployeePhoto(', fnStart);
  vm.runInContext(source.slice(fnStart, fnEnd), context);

  const file1 = new File(['a'], 'face1.jpg', { type: 'image/jpeg' });
  const file2 = new File(['b'], 'face2.png', { type: 'image/png' });
  await context.handleUploadEmployeePhoto({ files: [file1, file2] });

  assert.equal(calls.length, 1);
  const form = calls[0].options.body;
  assert.equal(form.get('employee_id'), 'NV002');
  assert.equal(form.get('project_id'), 'p2');
  assert.equal(form.getAll('files').length, 2);
});

test('deleteEmployeePhoto handles URL-encoded filenames and cleans path', async () => {
  const calls = [];
  const context = vm.createContext({
    activeEmpModalName: 'NV003',
    currentProjectName: 'CongTyC',
    allProjectsList: [{ name: 'CongTyC', project_id: 'p3' }],
    confirm: (msg) => {
      assert.match(msg, /my photo\.jpg/);
      return true;
    },
    apiPost: async (url, payload) => {
      calls.push({ url, payload });
      return { success: true };
    },
    showToast() {},
    loadPortraits: async () => {},
    refreshEmployeePhotosModal() {},
    loadProjects: async () => {}
  });
  const fnStart = source.indexOf('async function deleteEmployeePhoto(');
  const fnEnd = source.indexOf('async function deleteEmployee(', fnStart);
  vm.runInContext(source.slice(fnStart, fnEnd), context);

  // Pass an encoded filename with spaces or backslashes
  await context.deleteEmployeePhoto(encodeURIComponent('subfolder/my photo.jpg'));
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, '/api/portraits/employee/delete-photo');
  assert.equal(calls[0].payload.employee_id, 'NV003');
  assert.equal(calls[0].payload.filename, 'subfolder/my photo.jpg');
});
