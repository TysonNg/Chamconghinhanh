const {test, before, after} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const puppeteer = require('puppeteer-core');
const Downloader = require('../zalo-browser-downloader');
let browser, page;
before(async () => {
    const executablePath = ['C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'].find(p => fs.existsSync(p));
    browser = await puppeteer.launch({executablePath, headless: true});
    page = await browser.newPage();
});
after(async () => { if (browser) await browser.close(); });
function downloader() {
    const d = new Downloader();
    d.page = page;
    d.scanDelayMs = 50;
    d._log = () => {};
    return d;
}
test('group enumeration opens the group directory and collects virtualized pages by ID', async () => {
    await page.setContent(`<button title="Danh bạ" onclick="document.querySelector('#directory').hidden=false">Danh bạ</button>
        <div class="conv-item" data-id="person"><span class="conv-item-title__name">Bạn bè</span></div>
        <div id="directory" hidden><button onclick="document.querySelector('#groups').hidden=false">Danh sách nhóm</button></div>
        <div id="groups" hidden style="height:80px;overflow-y:auto"><div style="height:400px"><div id="rows"></div></div></div>`);
    await page.evaluate(() => {
        const list = document.querySelector('#groups');
        let index = 0;
        const render = () => document.querySelector('#rows').innerHTML =
            `<div class="group-item" data-id="g${index}"><span class="group-item-title__name">Nhóm trùng tên</span></div>`;
        render();
        list.addEventListener('scroll', () => { index = Math.min(4, Math.floor(list.scrollTop / 60)); render(); });
    });
    const d = downloader();
    d.getConnectionStatus = async () => ({isLoggedIn: true});
    const groups = await d.getGroups();
    assert.deepEqual(groups.map(g => g.id).sort(), ['g0', 'g1', 'g2', 'g3', 'g4']);
    assert.ok(groups.every(g => g.name === 'Nhóm trùng tên'));
});
test('current Zalo group directory reads stable item IDs and matches the displayed total', async () => {
    await page.setContent(`<button title="Danh bạ">Danh bạ</button><button>Danh sách nhóm và cộng đồng</button>
        <div class="contact-pages"><div class="card-list-title">Nhóm và cộng đồng (2)</div>
        <div class="contact-item-v2-wrapper"><span class="name">Nhóm A</span></div>
        <div class="contact-item-v2-wrapper"><span class="name">Nhóm B</span></div></div>`);
    await page.evaluate(() => document.querySelectorAll('.contact-item-v2-wrapper').forEach((row, i) => {
        row.__reactInternalInstance$fixture = {memoizedProps: {}, return: {memoizedProps: {itemId: `g${i + 1}`}}};
    }));
    const d = downloader();
    d.getConnectionStatus = async () => ({isLoggedIn: true});
    assert.deepEqual((await d.getGroups()).map(g => g.id), ['g1', 'g2']);
});
test('unbounded photo scan passes eight screens and waits for empty lazy-loaded batches', async () => {
    await page.setContent('<div id="innerScrollContainer" style="height:80px;overflow-y:auto"><div id="tiles" style="height:400px"></div></div>');
    await page.evaluate(() => {
        const list = document.querySelector('#innerScrollContainer');
        const tiles = document.querySelector('#tiles');
        let index = 0, loading = false, previousTop = 0;
        const render = () => tiles.innerHTML = `<img src="https://cdn.test/${index}.png" width="20" height="20">`;
        render();
        list.addEventListener('scroll', () => {
            if (loading || index >= 12 || list.scrollTop === previousTop) return;
            previousTop = list.scrollTop;
            loading = true;
            tiles.innerHTML = '';
            setTimeout(() => { index++; tiles.style.height = `${400 + index * 100}px`; render(); loading = false; }, 25);
        });
    });
    const found = await downloader()._collectPhotos();
    assert.equal(found.length, 13);
    assert.ok(found.some(p => p.url === 'https://cdn.test/12.png'));
});
test('photo scan does not stop at an old out-of-order date while newer images remain', async () => {
    await page.setContent('<div id="innerScrollContainer" style="height:80px;overflow-y:auto"><div id="tiles" style="height:400px"></div></div>');
    await page.evaluate(() => {
        const list = document.querySelector('#innerScrollContainer');
        const tiles = document.querySelector('#tiles');
        let index = 0;
        const render = () => tiles.innerHTML = `<div id="date_${index ? '12092026' : '01012026'}"></div><img src="https://cdn.test/${index}.png" width="20" height="20">`;
        render();
        list.addEventListener('scroll', () => { if (!index) { index++; render(); } });
    });
    const found = await downloader()._collectPhotos('2026-09-01', '2026-09-30');
    assert.deepEqual(found.map(p => p.date), ['2026-09-12']);
});
test('default scan collects more than 1000 images, while an explicit limit reports an incomplete scan', async () => {
    await page.setContent('<div id="innerScrollContainer"></div>');
    await page.evaluate(() => {
        document.querySelector('#innerScrollContainer').innerHTML = Array.from({length: 1002}, (_, i) =>
            `<img src="https://cdn.test/${i}.png" width="1" height="1">`).join('');
    });
    const d = downloader();
    assert.equal((await d._collectPhotos()).length, 1002);
    assert.equal(d.progress.scanComplete, true);
    assert.equal((await d._collectPhotos(null, null, 2)).length, 2);
    assert.equal(d.progress.scanComplete, false);
    assert.equal(d.progress.scanStopReason, 'limit');
});
test('an empty media store never borrows images from the background chat', async () => {
    await page.setContent('<div id="innerScrollContainer"></div><div class="chat-message"><img class="zimg-el" src="https://cdn.test/background.png" width="20" height="20"></div>');
    assert.deepEqual(await downloader()._collectPhotos(), []);
});
test('opening a group uses its ID when two conversations have the same name', async () => {
    await page.setContent('<div class="conv-item" data-id="wrong"><span class="conv-item-title__name">Trùng tên</span></div><div class="conv-item" data-id="right"><span class="conv-item-title__name">Trùng tên</span></div>');
    await page.evaluate(() => document.querySelectorAll('.conv-item').forEach(el => el.onclick = () => { document.body.dataset.opened = el.dataset.id; }));
    await downloader()._openGroup('Trùng tên', 'right');
    assert.equal(await page.evaluate(() => document.body.dataset.opened), 'right');
});
test('current conversation rows use the React item ID to distinguish duplicate names', async () => {
    await page.setContent('<div class="conv-item"><span class="conv-item-title__name">Trùng tên</span></div><div class="conv-item"><span class="conv-item-title__name">Trùng tên</span></div>');
    await page.evaluate(() => document.querySelectorAll('.conv-item').forEach((row, i) => {
        row.__reactInternalInstance$fixture = {memoizedProps: {}, return: {memoizedProps: {id: `g${i + 1}`}}};
        row.onclick = () => { document.body.dataset.opened = `g${i + 1}`; };
    }));
    await downloader()._openGroup('Trùng tên', 'g2');
    assert.equal(await page.evaluate(() => document.body.dataset.opened), 'g2');
});
test('a visible same-name group with a different ID is not opened for the requested group', async () => {
    await page.setContent('<div class="conv-item" data-id="wrong"><span class="conv-item-title__name">Trùng tên</span></div>');
    await page.evaluate(() => document.querySelector('.conv-item').onclick = () => { document.body.dataset.opened = 'wrong'; });
    await assert.rejects(downloader()._openGroup('Trùng tên', 'right'), /Không tìm thấy nhóm/);
    assert.equal(await page.evaluate(() => document.body.dataset.opened), undefined);
});
test('a closed media store reports an error instead of completing from background chat images', async () => {
    await page.setContent('<div class="chat-message"><img class="zimg-el" src="https://cdn.test/background.png" width="20" height="20"></div>');
    const d = downloader();
    await assert.rejects(d._collectPhotos(), /Kho Ảnh\/Video/);
    assert.equal(d.progress.scanComplete, false);
});
test('opening media selects the photo section instead of another See all button', async () => {
    await page.setContent(`<div class="chat-right-menu">
        <section><span>Tệp</span><button onclick="document.body.dataset.wrong='yes'">Xem tất cả</button></section>
        <section><span>Ảnh/Video</span><button onclick="document.body.insertAdjacentHTML('beforeend', '<div id=innerScrollContainer style=height:80px></div>')">Xem tất cả</button></section>
        </div>`);
    await downloader()._openMediaStore();
    assert.equal(await page.evaluate(() => !!document.querySelector('#innerScrollContainer')), true);
    assert.equal(await page.evaluate(() => document.body.dataset.wrong), undefined);
});
test('transient image download errors are retried before an image is reported missing', async () => {
    await page.setContent('<div></div>');
    const png = await page.evaluate(() => {
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = 2;
        return canvas.toDataURL('image/png').split(',')[1];
    });
    let attempts = 0;
    const handler = request => {
        if (request.url() !== 'https://cdn.test/retry.png') return request.continue();
        attempts++;
        return request.respond({status: attempts < 3 ? 503 : 200, contentType: 'image/png',
            headers: {'Access-Control-Allow-Origin': '*'}, body: Buffer.from(png, 'base64')});
    };
    await page.setRequestInterception(true);
    page.on('request', handler);
    try {
        const image = await downloader()._fetchImage('https://cdn.test/retry.png');
        assert.equal(image.width, 2);
        assert.deepEqual(image.buffer, Buffer.from(png, 'base64'));
        assert.equal(attempts, 3);
    } finally {
        page.off('request', handler);
        await page.setRequestInterception(false);
    }
});
