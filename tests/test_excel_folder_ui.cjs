const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');

function functionSection(name, nextName) {
  const start = source.indexOf(`async function ${name}(`);
  const end = source.indexOf(`async function ${nextName}(`, start + 1);
  assert.ok(start >= 0, `${name} must exist`);
  assert.ok(end > start, `${nextName} must follow ${name}`);
  return source.slice(start, end);
}

test('Word batch list renders an open-folder button', async () => {
  const nodes = {
    'excel-extracted-files': { innerHTML: '' },
    'stat-excel-persons': { textContent: '' },
  };
  const context = vm.createContext({
    apiGet: async () => ({
      folders: [{ folder: 'thang-08-2026', count: 1, files: [{ name: 'A.docx', size: 10 }] }],
    }),
    document: { getElementById: (id) => nodes[id] },
    withLoading: async (_button, callback) => callback(),
    formatFileSize: () => '10 B',
    encodeURIComponent,
    TRASH_ICON_SVG: '',
    console,
  });

  vm.runInContext(functionSection('loadExcelExtractedFiles', 'openExcelExtractedFolder'), context);
  await context.loadExcelExtractedFiles(null);

  assert.match(nodes['excel-extracted-files'].innerHTML, /Mở thư mục Word/);
  assert.match(
    nodes['excel-extracted-files'].innerHTML,
    /openExcelExtractedFolder\('thang-08-2026'\)/,
  );
});

test('open-folder action requests the selected Excel output batch', async () => {
  const calls = [];
  const messages = [];
  const context = vm.createContext({
    fetch: async (url, options) => {
      calls.push({ url, options });
      return { json: async () => ({ success: true }) };
    },
    showToast: (message, type) => messages.push({ message, type }),
  });

  vm.runInContext(functionSection('openExcelExtractedFolder', 'startExcelFaceAnalyze'), context);
  await context.openExcelExtractedFolder('thang-08-2026');

  assert.equal(calls[0].url, '/api/open/folder');
  assert.deepEqual(JSON.parse(calls[0].options.body), {
    type: 'excel_output',
    subpath: 'thang-08-2026',
  });
  assert.equal(messages[0].type, 'success');
});

