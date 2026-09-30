const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { SyncManager, getVietnamNowParts, getVietnamDateOffset } = require("../sync-manager");

test("getVietnamNowParts returns valid dateStr and timeStr", () => {
    const now = getVietnamNowParts();
    assert.match(now.dateStr, /^\d{4}-\d{2}-\d{2}$/);
    assert.match(now.timeStr, /^\d{1,2}:\d{1,2}$/);
    assert.ok(now.year >= 2026);
    assert.ok(now.month >= 1 && now.month <= 12);
    assert.ok(now.day >= 1 && now.day <= 31);
});

test("getVietnamDateOffset calculates past dates correctly", () => {
    const today = getVietnamDateOffset(0);
    const yesterday = getVietnamDateOffset(1);
    const threeDaysAgo = getVietnamDateOffset(3);

    assert.match(today, /^\d{4}-\d{2}-\d{2}$/);
    assert.match(yesterday, /^\d{4}-\d{2}-\d{2}$/);
    assert.match(threeDaysAgo, /^\d{4}-\d{2}-\d{2}$/);
    assert.ok(today > yesterday);
    assert.ok(yesterday > threeDaysAgo);
});

test("detectGaps accurately identifies missing dates and counts existing images", () => {
    // Create a temporary mock project folder structure
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "sync-gap-test-"));
    const inputImagesDir = path.join(tempDir, "input_images");
    const testProject = "TestProject";
    const projectDir = path.join(inputImagesDir, testProject);
    fs.mkdirSync(projectDir, { recursive: true });

    const todayStr = getVietnamDateOffset(0);
    const yesterdayStr = getVietnamDateOffset(1);
    const twoDaysAgoStr = getVietnamDateOffset(2);
    const threeDaysAgoStr = getVietnamDateOffset(3);

    // Yesterday has 3 mock images
    const yesterdayDir = path.join(projectDir, yesterdayStr);
    fs.mkdirSync(yesterdayDir, { recursive: true });
    fs.writeFileSync(path.join(yesterdayDir, "zalo_001.jpg"), "fake-bytes");
    fs.writeFileSync(path.join(yesterdayDir, "zalo_002.png"), "fake-bytes");
    fs.writeFileSync(path.join(yesterdayDir, "zalo_003.webp"), "fake-bytes");
    fs.writeFileSync(path.join(yesterdayDir, "zalo_001.jpg.json"), "{}"); // should be ignored in image count

    // threeDaysAgo has 1 image
    const threeDaysAgoDir = path.join(projectDir, threeDaysAgoStr);
    fs.mkdirSync(threeDaysAgoDir, { recursive: true });
    fs.writeFileSync(path.join(threeDaysAgoDir, "zalo_001.jpeg"), "fake-bytes");

    // twoDaysAgo is NOT created (missing)

    const syncMgr = new SyncManager(tempDir);
    const gaps = syncMgr.detectGaps(testProject, 5);

    assert.equal(gaps.projectName, testProject);
    assert.equal(gaps.lookbackDays, 5);
    assert.equal(gaps.dates.length, 5);

    // Check yesterday
    const yestEntry = gaps.dates.find(d => d.date === yesterdayStr);
    assert.ok(yestEntry);
    assert.equal(yestEntry.photoCount, 3);
    assert.equal(yestEntry.status, "ok");

    // Check 2 days ago
    const twoDaysEntry = gaps.dates.find(d => d.date === twoDaysAgoStr);
    assert.ok(twoDaysEntry);
    assert.equal(twoDaysEntry.photoCount, 0);
    assert.equal(twoDaysEntry.status, "missing");

    // Check missingDates array
    assert.ok(gaps.missingDates.includes(twoDaysAgoStr));
    assert.ok(gaps.hasMissing);
    assert.equal(typeof gaps.earliestMissingDate, "string");

    // Cleanup temp dir
    fs.rmSync(tempDir, { recursive: true, force: true });
});

test("loadConfig and saveConfig persists custom settings", () => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "sync-cfg-test-"));
    const syncMgr = new SyncManager(tempDir);
    syncMgr.configPath = path.join(tempDir, "custom_sync_config.json");

    const ok = syncMgr.saveConfig({
        enabled: true,
        scheduleTime: "21:45",
        lookbackDays: 7,
        mappings: [
            { groupId: "g1", groupName: "Group 1", projectName: "Project A", active: true }
        ]
    });
    assert.equal(ok, true);

    const reloaded = syncMgr.loadConfig();
    assert.equal(reloaded.enabled, true);
    assert.equal(reloaded.scheduleTime, "21:45");
    assert.equal(reloaded.lookbackDays, 7);
    assert.equal(reloaded.mappings.length, 1);
    assert.equal(reloaded.mappings[0].groupName, "Group 1");

    fs.rmSync(tempDir, { recursive: true, force: true });
});
