const test = require('node:test');
const assert = require('node:assert/strict');
const Downloader = require('../zalo-browser-downloader');
test('connection status rejects expired browser sessions', async () => {
    const d = new Downloader();
    d.browser = {connected: true};
    d.page = {isClosed: () => false, url: () => 'https://id.zalo.me/account', $: async () => ({})};
    d.isLoggedIn = true;
    const status = await d.getConnectionStatus();
    assert.equal(status.isLoggedIn, false);
});
test('connection status requires the Zalo chat interface', async () => {
    const d = new Downloader();
    d.browser = {connected: true};
    d.page = {isClosed: () => false, url: () => 'https://chat.zalo.me/', $: async () => ({})};
    assert.equal((await d.getConnectionStatus()).isLoggedIn, true);
    d.page.$ = async () => null;
    assert.equal((await d.getConnectionStatus()).isLoggedIn, false);
});
test('login requests share one browser login operation', async () => {
    const d = new Downloader();
    let calls = 0, complete;
    d.ensureLoggedIn = async () => { calls++; await new Promise(resolve => { complete = resolve; }); };
    const first = d.startLogin();
    const second = d.startLogin();
    assert.equal(first, second);
    assert.equal(calls, 1);
    complete();
    await first;
    assert.equal(d.loginPromise, null);
});
test('closed browser clears stale connection and QR', async () => {
    const d = new Downloader();
    d.browser = {connected: false};
    d.isLoggedIn = true;
    d.currentQrBase64 = 'stale';
    await d.getConnectionStatus();
    assert.equal(d.isLoggedIn, false);
    assert.equal(d.currentQrBase64, null);
});
test('download reuses a pending login without starting another login', async () => {
    const d = new Downloader();
    d.ensureProjectStructure = () => {};
    d._log = () => {};
    let loggedIn = false;
    d.loginPromise = Promise.resolve().then(() => { loggedIn = true; });
    d.ensureLoggedIn = async () => { throw new Error('duplicate login'); };
    d._openGroup = async () => { assert.equal(loggedIn, true); };
    d._checkAndClickSyncPrompt = async () => {};
    d._refreshGroupMessages = async () => {};
    d._openMediaStore = async () => {};
    d._collectPhotos = async () => [];
    const result = await d.downloadGroupPhotos({groupName: 'Fixture'});
    assert.equal(result.status, 'done');
});
test('failed login exposes an error instead of leaving download progress waiting for QR', async () => {
    const d = new Downloader();
    d.ensureLoggedIn = async () => { throw new Error('Login timed out'); };
    d.progress.status = 'waiting_qr';
    d.progress.qrExpiresAt = Date.now();
    await assert.rejects(d.startLogin(), /Login timed out/);
    assert.equal(d.progress.status, 'error');
    assert.equal(d.progress.error, 'Login timed out');
    assert.equal(d.progress.qrExpiresAt, null);
});
