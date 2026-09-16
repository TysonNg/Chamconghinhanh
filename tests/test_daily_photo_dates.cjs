const test = require('node:test');
const assert = require('node:assert/strict');
test('calendar selection contains month and year, and rejects impossible days', () => {
  const {fullDate} = require('../static/js/attendance-date.js');
  assert.equal(fullDate('2026-09', '01'), '2026-09-01');
  assert.equal(fullDate('2024-02', '29'), '2024-02-29');
  assert.throws(() => fullDate('2026-02','29'));
  assert.throws(() => fullDate('2026-09','31'));
  assert.throws(() => fullDate('', '01'));
});
test('Zalo browser downloader refuses day-only output even for empty selection', async () => {
  const Downloader = require('../zalo-service/zalo-browser-downloader');
  const d = new Downloader();
  await assert.rejects(d._downloadPhotos([], 'Site', 'DD'), /YYYY-MM-DD/);
});

test('API downloader uses Vietnam full dates and rejects day-only folders', async () => {
  const Downloader = require('../zalo-service/image-downloader');
  const d = Object.create(Downloader.prototype);
  assert.equal(d._formatDate('2026-08-31T18:00:00Z').ymd, '2026-09-01');
  assert.equal(d._formatDate(null).ymd, '');
  await assert.rejects(d.downloadImages([], 'Site', null, null, 'day'), /YYYY-MM-DD/);
});
