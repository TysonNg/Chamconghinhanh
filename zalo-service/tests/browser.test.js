const {test, before, after} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const puppeteer = require('puppeteer-core');
const {collectVisiblePhotos, normalizeSendDate} = require('../photo-metadata');
const Downloader = require('../zalo-browser-downloader');
let browser, page;
let png, fullPng;
before(async () => {
    const executablePath = [
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    ].find(p => fs.existsSync(p));
    assert.ok(executablePath, 'Chromium required for local DOM fixtures');
    browser = await puppeteer.launch({executablePath, headless: true});
    page = await browser.newPage();
    [png, fullPng] = await page.evaluate(() => [1, 16].map(size => {
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = size;
        canvas.getContext('2d').fillRect(0, 0, size, size);
        return canvas.toDataURL('image/png').split(',')[1];
    }));
});
after(async () => { if (browser) await browser.close(); });

test('DOM scan associates each album image with its own full-year header', async () => {
    await page.setContent(`<div id="innerScrollContainer">
        <div id="date_12092026">12/09/2026</div>
        <div id="item_12092026"><img src="https://cdn.test/a?t=1789232460"><img src="https://cdn.test/b"></div>
        <div id="date_13092026">13/09/2026</div>
        <div><img src="https://cdn.test/c"></div></div>`);
    const photos = await page.evaluate(collectVisiblePhotos);
    assert.deepEqual(photos.map(normalizeSendDate), ['2026-09-12', '2026-09-12', '2026-09-13']);
    assert.equal(new Set(photos.map(p => p.id)).size, 3);
});

test('virtualized next scan never inherits the previous scan last date', async () => {
    await page.setContent('<div id="innerScrollContainer"><div id="date_12092026"></div><div><img src="https://cdn.test/a"></div></div>');
    await page.evaluate(collectVisiblePhotos);
    await page.setContent('<div id="innerScrollContainer"><div><img src="https://cdn.test/b?t=1789232460"></div></div>');
    assert.equal(normalizeSendDate((await page.evaluate(collectVisiblePhotos))[0]), '');
});

test('a yearless header is parsed with the current year and does not borrow the date above it', async () => {
    await page.setContent('<div id="innerScrollContainer"><div id="date_31122025"></div><div class="media-item-date-group">Ngày 1 Tháng 1</div><div><img src="https://cdn.test/a"></div></div>');
    const date = normalizeSendDate((await page.evaluate(collectVisiblePhotos))[0]);
    // Yearless header "Ngày 1 Tháng 1" is parsed with the current year, NOT borrowing 2025 from above
    const expectedYear = new Date().getFullYear();
    assert.equal(date, `${expectedYear}-01-01`);
});

test('real image decoding rejects HTML and preserves image bytes', async () => {
    const d = new Downloader();
    d.page = page;
    const image = await d._fetchImage(`data:image/png;base64,${png}`);
    assert.equal(image.width, 1);
    assert.equal(image.height, 1);
    assert.equal(image.extension, 'png');
    assert.deepEqual(image.buffer, Buffer.from(png, 'base64'));
    await assert.rejects(d._fetchImage('data:text/html,<html>expired</html>'));
});

test('viewer native download supplies full bytes and a send date for an undated tile', async () => {
    await page.setContent(`<div id="innerScrollContainer"><img id="tile" src="data:image/png;base64,${png}"></div>`);
    await page.evaluate(pngData => {
        document.querySelector('#tile').onclick = () => {
            const dialog = document.createElement('div');
            dialog.setAttribute('role', 'dialog');
            dialog.innerHTML = `<time datetime="2026-09-12T17:01:00+07:00"></time><img src="data:image/png;base64,${pngData}"><a download="original.png" href="data:image/png;base64,${pngData}" title="Tải xuống">Tải xuống</a>`;
            document.body.append(dialog);
            document.addEventListener('keydown', e => { if (e.key === 'Escape') dialog.remove(); }, {once: true});
        };
    }, fullPng);
    const d = new Downloader();
    d.page = page;
    const photo = (await page.evaluate(collectVisiblePhotos))[0];
    const resolved = await d._resolvePhoto(photo);
    assert.ok(resolved.fullImage, resolved.qualityWarning);
    assert.deepEqual(resolved.fullImage.buffer, Buffer.from(fullPng, 'base64'));
    assert.equal(resolved.fullImage.width, 16);
    assert.equal(normalizeSendDate(resolved), '2026-09-12');
    assert.equal(await page.$('[role="dialog"]'), null);
});
test('expired QR refresh clicks the action inside the overlay and returns a new image', async () => {
    await page.setContent(`<div class="qrcode"><div class="qr-container"><svg width="236" height="236" xmlns="http://www.w3.org/2000/svg"><rect width="236" height="236" fill="white"/><rect x="10" y="10" width="40" height="40"/></svg></div><div class="qrcode-expired"><p>Expired</p><a class="btn" href="#">Refresh</a></div></div>`);
    await page.evaluate(() => {
        document.querySelector('.qrcode-expired a').onclick = event => {
            event.preventDefault();
            document.querySelector('.qrcode-expired').style.display = 'none';
        };
    });
    const d = new Downloader();
    d.page = page;
    d.progress.status = 'waiting_qr';
    d._log = () => {};
    const image = await d.refreshQrCode();
    assert.ok(image?.startsWith('data:image/png;base64,'), 'Refresh must activate the action and capture the QR');
    assert.equal(d.currentQrBase64, image);
});
