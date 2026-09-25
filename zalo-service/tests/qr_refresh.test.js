const test = require("node:test");
const assert = require("node:assert/strict");
const ZaloBrowserDownloader = require("../zalo-browser-downloader");

test("ZaloBrowserDownloader exposes refreshQrCode, cancelDownload, and _captureQrImage methods", () => {
    const downloader = new ZaloBrowserDownloader();
    assert.equal(typeof downloader.refreshQrCode, "function");
    assert.equal(typeof downloader.cancelDownload, "function");
    assert.equal(typeof downloader._captureQrImage, "function");
    assert.equal(typeof downloader.syncSessionFromCredentials, "function");
    assert.equal(typeof downloader.setBrowserVisible, "function");
});

test("syncSessionFromCredentials returns false gracefully when no cookies provided", async () => {
    const downloader = new ZaloBrowserDownloader();
    const result1 = await downloader.syncSessionFromCredentials(null);
    assert.equal(result1, false);
    const result2 = await downloader.syncSessionFromCredentials({});
    assert.equal(result2, false);
    const result3 = await downloader.syncSessionFromCredentials({ cookie: [] });
    assert.equal(result3, false);
});

test("refreshQrCode throws error when not in waiting_qr state or no page", async () => {
    const downloader = new ZaloBrowserDownloader();
    downloader.progress.status = "idle";
    await assert.rejects(async () => {
        await downloader.refreshQrCode();
    }, /Không có tiến trình chờ quét QR nào đang chạy/);
});

test("cancelDownload resets downloading state and progress", async () => {
    const downloader = new ZaloBrowserDownloader();
    downloader.isDownloading = true;
    downloader.progress.status = "waiting_qr";
    downloader.progress.qrImage = "data:image/png;base64,fake";
    downloader.currentQrBase64 = "data:image/png;base64,fake";

    await downloader.cancelDownload();

    assert.equal(downloader.abortLogin, true);
    assert.equal(downloader.isDownloading, false);
    assert.equal(downloader.progress.status, "idle");
    assert.equal(downloader.progress.qrImage, null);
    assert.equal(downloader.currentQrBase64, null);
});
