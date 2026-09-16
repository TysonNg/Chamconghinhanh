const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const launcher = fs.readFileSync(
  path.resolve(__dirname, '..', 'PhanMemQuetMat.bat'),
  'utf8',
);

assert.match(
  launcher,
  /-c "import flask, piexif"/,
  'Launcher must verify piexif before it skips dependency installation.',
);

const main = fs.readFileSync(path.resolve(__dirname, '..', 'main.py'), 'utf8');
assert.ok(
  main.indexOf('from src.app import app, scan_database, FLASK_PORT, FLASK_HOST')
    < main.indexOf('browser_thread = threading.Thread'),
  'Browser must be scheduled only after the Flask app imports successfully.',
);

assert.match(
  main,
  /sys\.version_info\s*<\s*\(3,\s*14\)/,
  'main.py must detect unsupported Python 3.14+ runtimes.',
);
assert.match(
  main,
  /['"]Python312['"]/, 
  'main.py must locate the installed Python 3.12 interpreter.',
);
assert.match(
  main,
  /subprocess\.run\(/,
  'main.py must relaunch with Python 3.12 using Windows-safe argument quoting.',
);

console.log('Launcher and Python runtime startup checks pass.');
