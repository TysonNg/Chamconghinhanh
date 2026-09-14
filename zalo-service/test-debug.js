const fs = require("fs");

async function run() {
    try {
        const url = "https://zalo-chat-static.zadn.vn/v1/sync-v2-worker.fa66597fb3affab4031b.js";
        const res = await fetch(url);
        const text = await res.text();

        let idx = 0;
        while ((idx = text.indexOf("mini_cloud", idx)) !== -1) {
            console.log(`\nmini_cloud at ${idx}:\n`, text.slice(Math.max(0, idx - 100), Math.min(text.length, idx + 300)));
            idx += 10;
        }
    } catch (e) {
        console.error(e);
    }
}
run();
