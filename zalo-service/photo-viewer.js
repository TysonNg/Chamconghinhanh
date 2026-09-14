const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// Decode inside Chromium, but return the original encoded bytes (never a canvas export).
async function readImage(page, source) {
    const data = await page.evaluate(async input => {
        let blob;
        if (input.base64) {
            blob = new Blob([Uint8Array.from(atob(input.base64), c => c.charCodeAt(0))]);
        } else {
            const response = await fetch(input.url, {signal: AbortSignal.timeout(30000)});
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            blob = await response.blob();
        }
        const bytes = new Uint8Array(await blob.arrayBuffer());
        let extension;
        if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) extension = 'jpg';
        else if ([137, 80, 78, 71, 13, 10, 26, 10].every((b, i) => bytes[i] === b)) extension = 'png';
        else if (String.fromCharCode(...bytes.slice(0, 6)).match(/^GIF8[79]a$/)) extension = 'gif';
        else if (String.fromCharCode(...bytes.slice(0, 4)) === 'RIFF' && String.fromCharCode(...bytes.slice(8, 12)) === 'WEBP') extension = 'webp';
        else throw new Error('Dữ liệu tải về không phải ảnh JPEG/PNG/GIF/WebP hợp lệ');
        const bitmap = await createImageBitmap(blob);
        const width = bitmap.width, height = bitmap.height;
        bitmap.close();
        if (!width || !height) throw new Error('Ảnh không có kích thước hợp lệ');
        const base64 = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result.split(',')[1]);
            reader.onerror = () => reject(new Error('Không đọc được dữ liệu ảnh'));
            reader.readAsDataURL(blob);
        });
        return {base64, width, height, extension};
    }, source);
    return {buffer: Buffer.from(data.base64, 'base64'), width: data.width, height: data.height, extension: data.extension};
}

async function downloadFromViewer(page, dialog) {
    const control = await dialog.evaluateHandle(root => Array.from(root.querySelectorAll('a, button, [role="button"], [title], [aria-label]')).find(el => {
        if (!el.getClientRects().length) return false;
        return el.matches('a[download]') || /^(tải xuống|tải về|lưu về máy|download)(\s|$)/i.test(
            (el.getAttribute('title') || el.getAttribute('aria-label') || el.textContent || '').trim());
    }) || null);
    if (!control.asElement()) {
        await control.dispose();
        throw new Error('Trình xem chưa cung cấp nút tải ảnh có thể xác minh');
    }
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'zalo-image-download-'));
    let client, timer, guid;
    try {
        client = await page.target().createCDPSession();
        const {frameTree} = await client.send('Page.getFrameTree');
        await client.send('Browser.setDownloadBehavior', {behavior: 'allowAndName', downloadPath: directory, eventsEnabled: true});
        const completed = new Promise((resolve, reject) => {
            timer = setTimeout(() => reject(new Error('Hết thời gian chờ tải bản đầy đủ')), 30000);
            client.on('Browser.downloadWillBegin', event => {
                if (event.frameId === frameTree.frame.id && !guid) guid = event.guid;
            });
            client.on('Browser.downloadProgress', event => {
                if (event.guid !== guid) return;
                if (event.state === 'completed') resolve();
                if (event.state === 'canceled') reject(new Error('Zalo không tải được bản đầy đủ'));
            });
        });
        // Attach the rejection handler before click, so click failure never leaves an unhandled timer.
        completed.catch(() => {});
        await control.asElement().click();
        await completed;
        if (!/^[a-zA-Z0-9-]+$/.test(guid || '')) throw new Error('Mã tải ảnh không hợp lệ');
        const buffer = fs.readFileSync(path.join(directory, guid));
        return await readImage(page, {base64: buffer.toString('base64')});
    } finally {
        clearTimeout(timer);
        if (client) {
            if (guid) await client.send('Browser.cancelDownload', {guid}).catch(() => {});
            await client.send('Browser.setDownloadBehavior', {behavior: 'default', eventsEnabled: false}).catch(() => {});
            await client.detach().catch(() => {});
        }
        await control.dispose();
        // Only remove this invocation's generated temporary directory.
        if (path.dirname(path.resolve(directory)) === path.resolve(os.tmpdir()) && path.basename(directory).startsWith('zalo-image-download-')) {
            fs.rmSync(directory, {recursive: true, force: true, maxRetries: 3});
        }
    }
}

async function resolvePhoto(page, photo) {
    // Restrict selection to the scanned media/chat surface; a URL alone must never
    // select a thumbnail in another open dialog. Process before virtualized rows disappear.
    const tile = await page.evaluateHandle(p => {
        const images = Array.from(document.querySelectorAll('#innerScrollContainer img, .chat-message img.zimg-el, .msg-item img.zimg-el, .img-center-box img.zimg-el'));
        return images.find(img => {
            if ((img.currentSrc || img.src) !== p.url || !img.getClientRects().length) return false;
            if (p.elementToken && img.dataset.zaloPhotoToken !== p.elementToken) return false;
            if (!p.messageId) return true;
            const message = img.closest('[data-msg-id], [data-message-id]') || img.closest('.chat-message, .msg-item');
            return (message?.dataset.msgId || message?.dataset.messageId) === p.messageId
                && Array.from(message.querySelectorAll('img')).indexOf(img) === p.imageIndex;
        }) || null;
    }, photo);
    if (!tile.asElement()) {
        await tile.dispose();
        return {...photo, qualityWarning: 'Không tìm thấy đúng ảnh trong giao diện để mở bản đầy đủ'};
    }
    const previousDialogs = await page.evaluateHandle(() => Array.from(document.querySelectorAll('[role="dialog"]')).filter(el => el.getClientRects().length));
    let dialog;
    const result = {...photo};
    try {
        await tile.asElement().click();
        // Deliberately capability-based. Do not guess Zalo class names or CDN variants.
        dialog = await page.waitForFunction(previous => {
            const dialogs = Array.from(document.querySelectorAll('[role="dialog"]')).filter(el => el.getClientRects().length && !previous.includes(el) && el.querySelector('img'));
            return dialogs.length === 1 ? dialogs[0] : false;
        }, {timeout: 3000}, previousDialogs);
        const metadata = await dialog.evaluate(root => {
            const time = root.querySelector('time[datetime]');
            const fullDate = (time?.textContent || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            return {timestamp: time?.getAttribute('datetime') || '',
                date: fullDate ? `${fullDate[3]}-${fullDate[2].padStart(2, '0')}-${fullDate[1].padStart(2, '0')}` : ''};
        });
        if (!result.timestamp && metadata.timestamp) result.timestamp = metadata.timestamp;
        if (!result.date && metadata.date) { result.date = metadata.date; result.dateSource = 'viewer'; }
        try {
            result.fullImage = await downloadFromViewer(page, dialog);
        } catch (error) {
            result.qualityWarning = error.message;
        }
    } catch {
        result.qualityWarning = 'Chưa xác minh được trình xem/nút tải bản đầy đủ của ảnh này';
    } finally {
        // Closing is necessary even if the viewer has no semantic dialog role.
        await page.keyboard.press('Escape').catch(() => {});
        if (dialog) {
            await page.waitForFunction(root => !root.isConnected || !root.getClientRects().length, {timeout: 2000}, dialog).catch(() => {});
            await dialog.dispose();
        }
        await previousDialogs.dispose();
        await tile.dispose();
    }
    return result;
}

module.exports = {readImage, resolvePhoto};
