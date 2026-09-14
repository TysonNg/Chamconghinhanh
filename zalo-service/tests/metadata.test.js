const test = require('node:test');
const assert = require('node:assert/strict');
const {normalizeSendDate, mergePhotos} = require('../photo-metadata');

test('send times use Vietnam midnight independently of the host timezone', () => {
    for (const tz of ['UTC', 'America/Los_Angeles', 'Asia/Ho_Chi_Minh']) {
        const previous = process.env.TZ;
        process.env.TZ = tz;
        try {
            assert.equal(normalizeSendDate({timestamp: '2026-09-12T16:59:00Z'}), '2026-09-12');
            assert.equal(normalizeSendDate({timestamp: '2026-09-12T17:01:00Z'}), '2026-09-13');
            assert.equal(normalizeSendDate({timestamp: 1789232460}), '2026-09-13');
            assert.equal(normalizeSendDate({timestamp: 1789232460000}), '2026-09-13');
        } finally {
            if (previous === undefined) delete process.env.TZ;
            else process.env.TZ = previous;
        }
    }
});

test('unknown dates, CDN timestamps and impossible dates do not acquire a day', () => {
    assert.equal(normalizeSendDate({url: 'https://cdn.test/photo?t=1789232460'}), '');
    assert.equal(normalizeSendDate({date: '2026-02-30', dateSource: 'header'}), '');
    assert.equal(normalizeSendDate({date: '2026-09-12'}), '');
    assert.equal(normalizeSendDate({timestamp: '2026-09-12T17:01:00'}), '');
});

test('full-year headers survive year rollover, send time takes precedence', () => {
    assert.equal(normalizeSendDate({date: '2025-12-31', dateSource: 'header'}), '2025-12-31');
    assert.equal(normalizeSendDate({date: '2025-12-31', dateSource: 'header', timestamp: '2025-12-31T17:01:00Z'}), '2026-01-01');
});

test('dedup merges chat and media copies but retains different sends of the same image', () => {
    const photos = mergePhotos([
        {id: 'url:a', url: 'a', date: '2026-09-12'},
        {id: 'm1:0', messageId: 'm1', url: 'a', date: '2026-09-12'},
        {id: 'm2:0', messageId: 'm2', url: 'a', date: '2026-09-13'},
    ]);
    assert.equal(photos.length, 2);
    assert.equal(photos[0].messageId, 'm1');
});

test('same URL shown under two different days is not collapsed', () => {
    assert.equal(mergePhotos([
        {id: 'url:a', url: 'a', date: '2026-09-12'},
        {id: 'url:a', url: 'a', date: '2026-09-13'},
    ]).length, 2);
});
