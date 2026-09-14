const fs = require("fs");
const puppeteer = require("puppeteer-core");

function getChromePath() {
    const paths = [
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
        "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    ];
    for (const p of paths) {
        if (fs.existsSync(p)) return p;
    }
    throw new Error("Chrome or Edge not found");
}

async function test() {
    const chromePath = getChromePath();
    const creds = JSON.parse(fs.readFileSync("./credentials.json", "utf-8"));
    const browser = await puppeteer.launch({
        executablePath: chromePath,
        headless: true,
        args: [
            "--no-sandbox",
            "--disable-setuid-sandbox",
        ]
    });

    try {
        const page = await browser.newPage();
        
        // Navigate to id.zalo.me first to establish context
        await page.goto("https://id.zalo.me/account", { waitUntil: "domcontentloaded" });

        // Now set cookies on the active page
        for (const c of creds.cookie) {
            try {
                await page.setCookie({
                    name: c.key || c.name,
                    value: c.value,
                    domain: c.domain,
                    path: c.path || "/",
                });
            } catch (e) {
                console.log("Error setting cookie:", c.key, e.message);
            }
        }

        console.log("Cookies after set:", (await page.cookies()).map(c => ({ name: c.name, domain: c.domain })));

        // Now go to chat.zalo.me
        console.log("Going to https://chat.zalo.me...");
        await page.goto("https://chat.zalo.me", { waitUntil: "networkidle2" });
        console.log("Final URL:", page.url());
        console.log("Final title:", await page.title());

        await page.screenshot({ path: "zalo-test-screen2.png" });

    } finally {
        await browser.close();
    }
}

test().catch(err => console.error("Test error:", err));
