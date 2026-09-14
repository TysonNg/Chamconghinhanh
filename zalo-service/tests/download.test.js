const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const Downloader = require('../zalo-browser-downloader');

// Network/browser operations are the boundary; folder selection and writes are real.
function fixture(t) {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'zalo-download-test-'));
    t.after(() => {
        assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
        assert.ok(path.basename(root).startsWith('zalo-download-test-'));
        fs.rmSync(root, {recursive: true, force: true});
    });
    const downloader = new Downloader(root);
    downloader._log = message => downloader.progress.log.push(message);
    downloader._downloadImageBuffer = async () => Buffer.from('preview');
    downloader._resolvePhoto = async photo => photo;
    downloader._fetchImage = async url => ({buffer: Buffer.from(url), width: 120, height: 80, extension: 'jpg'});
    return {downloader, root};
}

test('an undated photo is not silently saved under the selected start date', async t => {
    const {downloader, root} = fixture(t);
    await downloader._downloadPhotos([{id: 'missing', url: 'preview', date: ''}], 'Project', 'DD', '2026-09-12', '2026-09-12');
    assert.equal(fs.existsSync(path.join(root, 'input_images', 'Project', '12')), false);
    assert.equal(downloader.progress.unknownDate, 1);
    assert.equal(downloader.progress.downloaded, 0);
});

test('a send date outside the selected range never gets written', async t => {
    const {downloader, root} = fixture(t);
    await downloader._downloadPhotos([{id: 'next-day', url: 'preview', date: '2026-09-13', dateSource: 'message'}], 'Project', 'DD', '2026-09-12', '2026-09-12');
    assert.equal(fs.existsSync(path.join(root, 'input_images', 'Project', '13')), false);
    assert.equal(downloader.progress.downloaded, 0);
});

test('full image bytes take precedence over a preview', async t => {
    const {downloader, root} = fixture(t);
    await downloader._downloadPhotos([{id: 'full', url: 'preview', downloadUrl: 'full-image', date: '2026-09-12', dateSource: 'message'}], 'Project', 'DD', '2026-09-12', '2026-09-12');
    const dir = path.join(root, 'input_images', 'Project', '12');
    const file = fs.readdirSync(dir).find(name => name.endsWith('.jpg'));
    assert.equal(fs.readFileSync(path.join(dir, file), 'utf8'), 'full-image');
    assert.equal(downloader.progress.lowQuality, 0);
});

test('unavailable full image falls back to preview with a warning', async t => {
    const {downloader} = fixture(t);
    downloader._fetchImage = async url => {
        if (url === 'full-image') throw new Error('expired');
        return {buffer: Buffer.from('preview'), width: 120, height: 80, extension: 'jpg'};
    };
    await downloader._downloadPhotos([{id: 'fallback', url: 'preview', downloadUrl: 'full-image', date: '2026-09-12', dateSource: 'message'}], 'Project', 'DD');
    assert.equal(downloader.progress.downloaded, 1);
    assert.equal(downloader.progress.lowQuality, 1);
    assert.match(downloader.progress.log.join('\n'), /120.*80/);
});

test('failure of both sources does not leave a saved image', async t => {
    const {downloader, root} = fixture(t);
    downloader._downloadImageBuffer = downloader._fetchImage = async () => { throw new Error('unavailable'); };
    await downloader._downloadPhotos([{id: 'failed', url: 'preview', downloadUrl: 'full-image', date: '2026-09-12', dateSource: 'message'}], 'Project', 'DD');
    const dir = path.join(root, 'input_images', 'Project', '12');
    assert.deepEqual(fs.existsSync(dir) ? fs.readdirSync(dir) : [], []);
    assert.equal(downloader.progress.failed, 1);
});

test('existing files and gaps in their numbering are preserved', async t => {
    const {downloader, root} = fixture(t);
    const dir = path.join(root, 'input_images', 'Project', '12');
    fs.mkdirSync(dir, {recursive: true});
    fs.writeFileSync(path.join(dir, 'zalo_12_003.jpg'), 'old');
    await downloader._downloadPhotos([{id: 'new', url: 'preview', date: '2026-09-12', dateSource: 'header'}], 'Project', 'DD');
    assert.equal(fs.readFileSync(path.join(dir, 'zalo_12_003.jpg'), 'utf8'), 'old');
    assert.equal(fs.readFileSync(path.join(dir, 'zalo_12_004.jpg'), 'utf8'), 'preview');
});

test('date recovered from viewer is filtered before writing and YYYY-MM-DD is supported', async t => {
    const {downloader, root} = fixture(t);
    downloader._resolvePhoto = async photo => ({...photo, timestamp: '2026-09-12T17:01:00Z'});
    await downloader._downloadPhotos([{id: 'new', url: 'preview'}], 'Project', 'YYYY-MM-DD', '2026-09-12', '2026-09-12');
    assert.equal(downloader.progress.downloaded, 0);
    assert.equal(downloader.progress.filteredOut, 1);
    await downloader._downloadPhotos([{id: 'new', url: 'preview'}], 'Project', 'YYYY-MM-DD', '2026-09-13', '2026-09-13');
    assert.equal(fs.existsSync(path.join(root, 'input_images', 'Project', '2026-09-13', 'zalo_2026-09-13_001.jpg')), true);
});
