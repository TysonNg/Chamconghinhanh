/**
 * Zalo Service - Express API Server
 * Provides REST API for Flask app to interact with Zalo
 * Runs on port 3001
 */

const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");

// Global exception safety
process.on("uncaughtException", (err) => {
    console.error("[Zalo Service] Uncaught Exception:", err);
});
process.on("unhandledRejection", (reason) => {
    console.error("[Zalo Service] Unhandled Rejection:", reason);
});

const ZaloClient = require("./zalo-client");
const ImageDownloader = require("./image-downloader");

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

// Singleton instances
const zaloClient = new ZaloClient();
const imageDownloader = new ImageDownloader();
const ZaloBrowserDownloader = require("./zalo-browser-downloader");
const browserDownloader = new ZaloBrowserDownloader(path.resolve(__dirname, ".."));
const { collectVisiblePhotos, normalizeSendDate } = require("./photo-metadata");

// SSE clients for progress updates
let sseClients = [];

// Setup progress listener for imageDownloader
imageDownloader.onProgress((data) => {
    for (const client of sseClients) {
        try {
            client.write(`data: ${JSON.stringify(data)}\n\n`);
        } catch {}
    }
});

// Periodic progress push for browserDownloader
setInterval(() => {
    if (browserDownloader.isDownloading && sseClients.length > 0) {
        const p = browserDownloader.getStatus().progress;
        for (const client of sseClients) {
            try {
                client.write(`data: ${JSON.stringify(p)}\n\n`);
            } catch {}
        }
    }
}, 800);

// ==================== API Routes ====================

/**
 * GET /api/status
 * Check Zalo connection status
 */
app.get("/api/status", (req, res) => {
    res.json({
        success: true,
        data: zaloClient.getStatus(),
    });
});

/**
 * POST /api/login/qr
 * Start QR code login
 */
app.post("/api/login/qr", async (req, res) => {
    try {
        const { force } = req.body || {};
        if (force) {
            console.log("[Server] Force refreshing QR login...");
            await zaloClient.logout();
        }

        if (zaloClient.isLoggedIn && !force) {
            return res.json({
                success: true,
                message: "Already logged in",
                data: zaloClient.getStatus(),
            });
        }

        if (zaloClient.loginInProgress && !force) {
            return res.json({
                success: true,
                message: "Login already in progress",
                data: zaloClient.getStatus(),
            });
        }

        // Start QR login in background (non-blocking)
        zaloClient.loginWithQR().catch((err) => {
            console.error("[Server] QR login error:", err.message);
        });

        // Wait a moment for QR to be generated
        await new Promise((resolve) => setTimeout(resolve, 3000));

        res.json({
            success: true,
            message: "QR login started",
            data: zaloClient.getStatus(),
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message,
        });
    }
});

/**
 * GET /api/login/qr/image
 * Get the current QR code image
 */
app.get("/api/login/qr/image", (req, res) => {
    let imgData = null;
    const qrPath = path.join(__dirname, "qr.png");
    if (fs.existsSync(qrPath)) {
        try {
            const fileBuf = fs.readFileSync(qrPath);
            imgData = fileBuf.toString("base64");
        } catch {}
    }

    if (!imgData && zaloClient.qrImageBase64) {
        imgData = zaloClient.qrImageBase64;
    }

    if (imgData) {
        if (!imgData.startsWith("data:image/")) {
            imgData = "data:image/png;base64," + imgData;
        }
        return res.json({
            success: true,
            data: {
                image: imgData,
                status: zaloClient.qrStatus,
                userInfo: zaloClient.qrUserInfo,
            },
        });
    }

    res.json({
        success: false,
        error: "No QR code available",
        data: { status: zaloClient.qrStatus },
    });
});

/**
 * GET /api/login/qr/raw
 * Serve the QR image directly as binary PNG
 */
app.get("/api/login/qr/raw", (req, res) => {
    const qrPath = path.join(__dirname, "qr.png");
    if (fs.existsSync(qrPath)) {
        return res.sendFile(qrPath);
    }
    if (zaloClient.qrImageBase64) {
        let b64 = zaloClient.qrImageBase64;
        if (b64.startsWith("data:image/")) {
            b64 = b64.split(",")[1];
        }
        const buf = Buffer.from(b64, "base64");
        res.setHeader("Content-Type", "image/png");
        return res.send(buf);
    }
    res.status(404).send("No QR code available");
});

/**
 * GET /api/login/qr/status
 * Poll for QR login status
 */
app.get("/api/login/qr/status", (req, res) => {
    res.json({
        success: true,
        data: {
            status: zaloClient.qrStatus,
            isLoggedIn: zaloClient.isLoggedIn,
            loginInProgress: zaloClient.loginInProgress,
            userInfo: zaloClient.qrUserInfo,
        },
    });
});

/**
 * POST /api/logout
 * Logout from Zalo
 */
app.post("/api/logout", async (req, res) => {
    try {
        await zaloClient.logout();
        res.json({
            success: true,
            message: "Logged out successfully",
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message,
        });
    }
});

/**
 * GET /api/groups
 * Get all Zalo groups
 */
app.get("/api/groups", async (req, res) => {
    try {
        const groups = await zaloClient.getGroups();
        res.json({
            success: true,
            data: groups,
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message,
        });
    }
});

/**
 * POST /api/groups/:groupId/download
 * Download images from a group
 * Body: { groupName, dateFrom, dateTo, count }
 */
app.post("/api/groups/:groupId/download", async (req, res) => {
    try {
        const { groupId } = req.params;
        const { groupName, projectName, dateFrom, dateTo, count = 1000, folderFormat = "YYYY-MM-DD", headless = true } = req.body;

        if (!groupName) {
            return res.status(400).json({
                success: false,
                error: "groupName is required",
            });
        }

        if (!["date", "YYYY-MM-DD"].includes(folderFormat)) {
            return res.status(400).json({success: false, error: "Chỉ hỗ trợ YYYY-MM-DD"});
        }
        const targetProject = projectName || groupName;
        console.log(`[Server] Browser download request: group=${groupName} (${groupId}) -> project=${targetProject}, from=${dateFrom}, to=${dateTo}, format=${folderFormat}, headless=${headless}`);

        if (browserDownloader.isDownloading) {
            return res.status(400).json({
                success: false,
                error: "Đang có tiến trình tải ảnh đang chạy, vui lòng chờ.",
            });
        }

        // Start browser downloader in background
        browserDownloader
            .downloadGroupPhotos({
                groupId,
                groupName,
                projectName: targetProject,
                fromDate: dateFrom,
                toDate: dateTo,
                folderFormat,
                headless: headless !== false,
                maxImages: count
            })
            .then((result) => {
                console.log("[Server] Browser download completed successfully:", result);
                for (const client of sseClients) {
                    try {
                        client.write(`data: ${JSON.stringify(result)}\n\n`);
                    } catch {}
                }
            })
            .catch((err) => {
                console.error("[Server] Browser download error:", err.message);
                for (const client of sseClients) {
                    try {
                        client.write(`data: ${JSON.stringify(browserDownloader.getStatus().progress)}\n\n`);
                    } catch {}
                }
            });

        // Return immediately to frontend
        res.json({
            success: true,
            message: `Bắt đầu quét & tải ảnh từ nhóm "${groupName}"...`,
            data: { status: "downloading" },
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message,
        });
    }
});

/**
 * GET /api/download/progress
 * SSE endpoint for download progress
 */
app.get("/api/download/progress", (req, res) => {
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");

    // Send current progress immediately
    const p = browserDownloader.isDownloading || browserDownloader.getStatus().progress.status !== "idle"
        ? browserDownloader.getStatus().progress
        : imageDownloader.getProgress();
    res.write(`data: ${JSON.stringify(p)}\n\n`);

    // Add to SSE clients
    sseClients.push(res);

    // Remove on disconnect
    req.on("close", () => {
        sseClients = sseClients.filter((c) => c !== res);
    });
});

/**
 * GET /api/download/progress/poll
 * Polling endpoint for download progress (alternative to SSE)
 */
app.get("/api/download/progress/poll", (req, res) => {
    const p = browserDownloader.isDownloading || browserDownloader.getStatus().progress.status !== "idle"
        ? browserDownloader.getStatus().progress
        : imageDownloader.getProgress();
    res.json({
        success: true,
        data: p,
    });
});

/**
 * POST /api/download/qr/refresh
 * Request browser downloader to refresh the QR code
 */
app.post("/api/download/qr/refresh", async (req, res) => {
    try {
        const newQr = await browserDownloader.refreshQrCode();
        res.json({
            success: true,
            data: { image: newQr, status: browserDownloader.getStatus().progress.status }
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message
        });
    }
});

/**
 * POST /api/download/cancel
 * Cancel download / QR waiting
 */
app.post("/api/download/cancel", async (req, res) => {
    try {
        await browserDownloader.cancelDownload();
        res.json({
            success: true,
            message: "Đã hủy tiến trình tải ảnh."
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.message
        });
    }
});

/**
 * Debug endpoint to inspect current browser DOM and screenshot
 */
app.get("/api/debug/dom", async (req, res) => {
    try {
        if (!browserDownloader.page) {
            return res.json({ error: "No page available in browserDownloader" });
        }
        const page = browserDownloader.page;
        const screenshotPath = path.join(__dirname, "debug_page.png");
        await page.screenshot({ path: screenshotPath });

        const domInfo = await page.evaluate(() => {
            const allImages = Array.from(document.querySelectorAll("img")).map(img => ({
                src: img.src ? img.src.slice(0, 100) : "",
                className: img.className,
                alt: img.alt,
                parentTag: img.parentElement ? img.parentElement.tagName : "",
                parentClass: img.parentElement ? img.parentElement.className : "",
                width: img.width,
                height: img.height
            }));

            const rightSidebar = document.querySelector(".chat-right-menu, .media-store-view, [data-id*='RightMenu'], .conversation-info")?.innerHTML?.slice(0, 500) || "No right sidebar found";

            const chatMessages = Array.from(document.querySelectorAll(".chat-message, .msg-item, .chat-item, [data-id*='msg']")).length;

            // Deep inspect innerScrollContainer
            const isc = document.querySelector('#innerScrollContainer');
            let innerScrollInfo = null;
            if (isc) {
                // Find all elements with text containing date patterns
                const dateElements = [];
                const walker = document.createTreeWalker(isc, NodeFilter.SHOW_ELEMENT, null);
                let node;
                while (node = walker.nextNode()) {
                    const text = (node.textContent || '').trim();
                    if (text.length < 80 && text.length > 3 && /Ngày|Tháng|Năm|\d{1,2}\/\d{1,2}|\d{2}\d{2}\d{4}/.test(text)) {
                        dateElements.push({
                            tag: node.tagName,
                            className: (node.className || '').slice(0, 100),
                            id: node.id || '',
                            text: text.slice(0, 80),
                            childCount: node.children.length,
                            display: window.getComputedStyle(node).display,
                            dataAttrs: Object.fromEntries(
                                Array.from(node.attributes || [])
                                    .filter(a => a.name.startsWith('data-'))
                                    .map(a => [a.name, a.value.slice(0, 50)])
                            )
                        });
                    }
                }

                // First 3 media-store images full parent chain
                const mediaImgs = isc.querySelectorAll('img');
                const imgChains = [];
                for (let i = 0; i < Math.min(3, mediaImgs.length); i++) {
                    const chain = [];
                    let el = mediaImgs[i];
                    for (let d = 0; el && el !== isc && d < 8; d++) {
                        chain.push({
                            tag: el.tagName,
                            className: (el.className || '').slice(0, 100),
                            id: el.id || '',
                            dataAttrs: Object.fromEntries(
                                Array.from(el.attributes || [])
                                    .filter(a => a.name.startsWith('data-'))
                                    .map(a => [a.name, a.value.slice(0, 50)])
                            )
                        });
                        el = el.parentElement;
                    }
                    imgChains.push(chain);
                }

                // Direct children of innerScrollContainer
                const directChildren = Array.from(isc.children).slice(0, 20).map(child => ({
                    tag: child.tagName,
                    className: (child.className || '').slice(0, 100),
                    id: child.id || '',
                    text: (child.textContent || '').trim().slice(0, 60),
                    childCount: child.children.length,
                    dataAttrs: Object.fromEntries(
                        Array.from(child.attributes || [])
                            .filter(a => a.name.startsWith('data-'))
                            .map(a => [a.name, a.value.slice(0, 50)])
                    )
                }));

                // Scroll container analysis
                const scrollAnalysis = [];
                let currEl = isc;
                while (currEl && currEl !== document.body) {
                    const style = window.getComputedStyle(currEl);
                    scrollAnalysis.push({
                        tag: currEl.tagName,
                        className: (currEl.className || '').slice(0, 50),
                        id: currEl.id || '',
                        overflowY: style.overflowY,
                        scrollHeight: currEl.scrollHeight,
                        clientHeight: currEl.clientHeight,
                        scrollTop: currEl.scrollTop
                    });
                    currEl = currEl.parentElement;
                }

                innerScrollInfo = {
                    totalImgs: mediaImgs.length,
                    dateElements,
                    imgChains,
                    directChildren,
                    scrollAnalysis,
                    innerHTML_first_1000: isc.innerHTML.slice(0, 1000)
                };
            }

            return {
                url: window.location.href,
                title: document.title,
                totalImages: allImages.length,
                images: allImages,
                chatMessagesCount: chatMessages,
                rightSidebarPreview: rightSidebar,
                innerScrollInfo
            };
        });

        // Test collectVisiblePhotos
        let testCollect = [];
        try {
            const raw = await page.evaluate(collectVisiblePhotos);
            testCollect = raw.map(p => ({
                id: p.id,
                source: p.source,
                date: p.date,
                normalizedDate: normalizeSendDate(p),
                dateSource: p.dateSource,
                url: (p.url || '').slice(0, 60)
            }));
        } catch (e) {
            testCollect = [{ error: e.message }];
        }

        // Test scroll step
        const scrollTest = await page.evaluate(() => {
            const isc = document.querySelector('#innerScrollContainer');
            if (!isc) return { error: "No isc" };
            let scrollEl = isc;
            while (scrollEl && scrollEl !== document.body) {
                const s = window.getComputedStyle(scrollEl);
                if (s.overflowY === 'scroll' || s.overflowY === 'auto') break;
                scrollEl = scrollEl.parentElement;
            }
            if (!scrollEl || scrollEl === document.body) return { error: "No scroll container found" };
            const before = {
                tag: scrollEl.tagName,
                className: scrollEl.className,
                scrollTop: scrollEl.scrollTop,
                scrollHeight: scrollEl.scrollHeight,
                clientHeight: scrollEl.clientHeight
            };
            scrollEl.scrollTop += 800;
            scrollEl.dispatchEvent(new Event('scroll', { bubbles: true }));
            const after = {
                scrollTop: scrollEl.scrollTop
            };
            return { before, after };
        });

        res.json({ success: true, domInfo, testCollect, scrollTest, screenshotPath });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

/**
 * Health check
 */
app.get("/api/health", (req, res) => {
    res.json({ status: "ok", service: "zalo-service", port: PORT });
});

// ==================== Start Server ====================

async function start() {
    app.listen(PORT, "0.0.0.0", () => {
        console.log(`\n[Zalo Service] Running on http://127.0.0.1:${PORT}`);
        console.log(`[Zalo Service] Status: ${zaloClient.isLoggedIn ? "Logged in" : "Not logged in"}`);
        console.log(`[Zalo Service] API endpoints:`);
        console.log(`  GET  /api/health`);
        console.log(`  GET  /api/status`);
        console.log(`  POST /api/login/qr`);
        console.log(`  GET  /api/login/qr/image`);
        console.log(`  GET  /api/login/qr/status`);
        console.log(`  POST /api/logout`);
        console.log(`  GET  /api/groups`);
        console.log(`  POST /api/groups/:groupId/download`);
        console.log(`  GET  /api/download/progress`);
        console.log(`  GET  /api/download/progress/poll\n`);
    });

    // Try auto-login with saved credentials in background without blocking server listen
    zaloClient.tryAutoLogin()
        .then((success) => {
            if (success) {
                console.log("[Server] Auto-login completed successfully");
            }
        })
        .catch((err) => {
            console.log("[Server] Auto-login skipped/failed:", err.message);
        });
}

start();
