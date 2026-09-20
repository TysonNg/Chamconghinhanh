const test = require('node:test');
const assert = require('node:assert/strict');
const Downloader = require('../zalo-browser-downloader');

test('date-range scan keeps scrolling beyond eight screens until it passes the start date', async () => {
    const batches = [
        '2026-06-08', '2026-06-07', '2026-06-06', '2026-06-05',
        '2026-06-04', '2026-06-03', '2026-06-02', '2026-06-01',
        '2026-05-15', '2026-04-30',
    ].map((date, index) => [{
        id: `m${index}:0`, messageId: `m${index}`, url: `https://cdn.test/${index}`,
        date, dateSource: 'header', timestamp: '',
    }]);
    let batchIndex = 0;
    const downloader = new Downloader();
    downloader.scanDelayMs = 0;
    downloader.page = {
        evaluate: async fn => {
            if (fn.name === 'collectVisiblePhotos') return batches[Math.min(batchIndex, batches.length - 1)];
            batchIndex++;
            return undefined;
        },
    };
    const found = [];
    await downloader._collectPhotos('2026-05-01', '2026-05-31', 200, async photos => found.push(...photos));
    assert.deepEqual(found.map(photo => photo.date), ['2026-05-15']);
    assert.ok(batchIndex >= 9, `expected at least 9 scrolls, got ${batchIndex}`);
});
