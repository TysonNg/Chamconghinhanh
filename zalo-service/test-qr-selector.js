const puppeteer = require("puppeteer-core");

async function run() {
    const browser = await puppeteer.launch({
        executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        headless: true,
        args: ["--no-sandbox"]
    });
    try {
        const page = await browser.newPage();
        await page.goto("https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F", { waitUntil: "networkidle2" });
        
        for (const frame of page.frames()) {
            console.log("Frame URL:", frame.url());
            const img = await frame.$("img");
            if (img) {
                const src = await frame.evaluate(el => el.src, img);
                console.log("Found img in frame! Src:", src?.slice(0, 100));
            }
        }
        
        // Also take a screenshot of .qr-container directly:
        const qrContainer = await page.$(".qr-container");
        if (qrContainer) {
            await qrContainer.screenshot({ path: "qr-box.png" });
            console.log("Saved qr-box.png!");
        }
    } finally {
        await browser.close();
    }
}
run().catch(console.error);
