const {test} = require('node:test');
const assert = require('node:assert/strict');
const {classifySendShift} = require('../photo-metadata');
const Downloader = require('../zalo-browser-downloader');
test('Vietnam shift boundaries and unknown time', () => {
  for (const [time, shift] of [['04:59:59','outside'],['05:00:00','morning'],['15:00:00','morning'],['15:00:01','outside'],['15:59:59','outside'],['16:00:00','afternoon'],['23:59:59','afternoon'],['00:00:00','outside']]) {
    assert.equal(classifySendShift({timestamp:`2026-10-01T${time}+07:00`}).shift, shift, time);
  }
  assert.equal(classifySendShift({timestamp:'2026-10-01T22:00:00Z'}).shift,'morning');
  assert.equal(classifySendShift({date:'2026-10-01', dateSource:'header'}).shift,'unknown');
});
test('session cookies use the supported browser context API', async () => {
  const d = Object.create(Downloader.prototype);
  d._log = () => {};
  d.progress = {};
  d.getOrLaunchBrowser = async () => {};
  let cookies;
  d.page = {browserContext: () => ({setCookie: async (...values) => {cookies = values;}}), goto: async () => {}, $: async () => ({})};
  assert.equal(await d.syncSessionFromCredentials({cookie:[{key:'session',value:'fixture',domain:'.zalo.me'}]}),true);
  assert.equal(cookies[0].name,'session');
  assert.equal(d.isLoggedIn,true);
});


test('an authenticated chat is accepted before checking QR or unrelated canvases', async () => {
  const d = Object.create(Downloader.prototype);
  d.getOrLaunchBrowser = async () => {};
  d._log = () => {};
  d._checkAndClickSyncPrompt = async () => {};
  d.progress = {};
  d.page = {url: () => 'https://chat.zalo.me/', $: async selector => {
    if (selector.includes('canvas')) throw new Error('canvas is not a QR login indicator');
    return selector.includes('#contact-search-input') ? {} : null;
  }};
  assert.equal(await d.ensureLoggedIn(true), true);
  assert.equal(d.isLoggedIn, true);
});
