const puppeteer = require("puppeteer-core");
const fs = require("fs");
const path = require("path");

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

async function run() {
    const chromePath = getChromePath();
    const profileDir = path.join(__dirname, "zalo-browser-profile");
    if (!fs.existsSync(profileDir)) {
        fs.mkdirSync(profileDir, { recursive: true });
    }

    console.log("Launching browser with profileDir:", profileDir);
    // Launch headless: false so the QR code appears for the user to scan if not logged in
    const browser = await puppeteer.launch({
        executablePath: chromePath,
        userDataDir: profileDir,
        headless: false,
        args: [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--window-size=1200,800"
        ]
    });

    try {
        const pages = await browser.pages();
        const page = pages.length > 0 ? pages[0] : await browser.newPage();
        await page.setViewport({ width: 1200, height: 800 });

        console.log("Navigating to https://chat.zalo.me...");
        await page.goto("https://chat.zalo.me", { waitUntil: "domcontentloaded" });

        console.log("Waiting for user login (or already logged in)...");
        // Wait up to 60 seconds to detect chat.zalo.me elements
        let loggedIn = false;
        for (let i = 0; i < 30; i++) {
            const url = page.url();
            const title = await page.title();
            console.log(`[${i}s] URL: ${url} | Title: ${title}`);

            if (url.includes("chat.zalo.me") && !url.includes("id.zalo.me") && !url.includes("/login")) {
                // Check if main UI elements exist
                const hasConv = await page.$("#contact-search-input, .conv-item, #main-tab");
                if (hasConv) {
                    console.log("Detected logged in Zalo Web UI!");
                    loggedIn = true;
                    break;
                }
            }
            await new Promise(r => setTimeout(r, 2000));
        }

        if (loggedIn) {
            console.log("SUCCESS: Logged in and session saved to profileDir!");
            await page.screenshot({ path: "zalo-chat-main.png" });
        } else {
            console.log("Timeout waiting for login.");
        }

    } finally {
        await browser.close();
    }
}

run().catch(console.error);
