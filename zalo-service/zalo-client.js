/**
 * Zalo Client Wrapper
 * Manages connection, login, and session persistence using zca-js
 */

const { Zalo, ThreadType, LoginQRCallbackEventType } = require("zca-js");
const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");

const CREDENTIALS_PATH = path.join(__dirname, "credentials.json");
const QR_IMAGE_PATH = path.join(__dirname, "qr.png");

class ZaloClient {
    constructor() {
        this.zalo = new Zalo();
        this.api = null;
        this.isLoggedIn = false;
        this.loginInProgress = false;
        this.qrImageBase64 = null;
        this.qrStatus = "idle"; // idle | waiting_scan | scanned | logged_in | error
        this.qrUserInfo = null; // { avatar, display_name } after scan
        this.loginAbort = null;
        this.credentials = null;
    }

    /**
     * Try to login with saved credentials
     */
    async tryAutoLogin() {
        if (!fs.existsSync(CREDENTIALS_PATH)) {
            console.log("[ZaloClient] No saved credentials found");
            return false;
        }

        try {
            const saved = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, "utf-8"));
            console.log("[ZaloClient] Found saved credentials, attempting login...");

            this.zalo = new Zalo();
            this.api = await this.zalo.login(saved);
            this.isLoggedIn = true;
            this.qrStatus = "logged_in";
            this.credentials = saved;
            console.log("[ZaloClient] Auto-login successful!");
            return true;
        } catch (err) {
            console.error("[ZaloClient] Auto-login failed:", err.message);
            // Remove invalid credentials
            try { fs.unlinkSync(CREDENTIALS_PATH); } catch {}
            return false;
        }
    }

    /**
     * Start QR code login flow
     * Returns a promise that resolves when login is complete
     */
    async loginWithQR() {
        if (this.loginInProgress) {
            throw new Error("Login already in progress");
        }
        if (this.isLoggedIn) {
            throw new Error("Already logged in");
        }

        this.loginInProgress = true;
        this.qrStatus = "waiting_scan";
        this.qrImageBase64 = null;
        this.qrUserInfo = null;

        try {
            this.zalo = new Zalo();
            this.api = await this.zalo.loginQR(
                { qrPath: QR_IMAGE_PATH },
                (event) => {
                    this._handleQREvent(event);
                }
            );

            this.isLoggedIn = true;
            this.qrStatus = "logged_in";
            this.loginInProgress = false;

            console.log("[ZaloClient] QR login successful!");
            return true;
        } catch (err) {
            this.loginInProgress = false;
            this.qrStatus = "error";
            console.error("[ZaloClient] QR login failed:", err.message);
            throw err;
        }
    }

    /**
     * Handle QR code events from zca-js
     */
    _handleQREvent(event) {
        switch (event.type) {
            case LoginQRCallbackEventType.QRCodeGenerated:
                console.log("[ZaloClient] QR Code generated");
                let imgData = event.data.image;
                if (imgData && !imgData.startsWith("data:image/")) {
                    imgData = "data:image/png;base64," + imgData;
                }
                this.qrImageBase64 = imgData;
                this.qrStatus = "waiting_scan";
                // Also read the saved QR file if available
                if (fs.existsSync(QR_IMAGE_PATH)) {
                    try {
                        const qrData = fs.readFileSync(QR_IMAGE_PATH);
                        this.qrImageBase64 = "data:image/png;base64," + qrData.toString("base64");
                    } catch {}
                }
                break;

            case LoginQRCallbackEventType.QRCodeExpired:
                console.log("[ZaloClient] QR Code expired, retrying...");
                this.qrStatus = "expired";
                event.actions.retry();
                break;

            case LoginQRCallbackEventType.QRCodeScanned:
                console.log("[ZaloClient] QR Code scanned by:", event.data.display_name);
                this.qrStatus = "scanned";
                this.qrUserInfo = event.data;
                break;

            case LoginQRCallbackEventType.QRCodeDeclined:
                console.log("[ZaloClient] QR Code declined");
                this.qrStatus = "declined";
                break;

            case LoginQRCallbackEventType.GotLoginInfo:
                console.log("[ZaloClient] Got login info, saving credentials...");
                // Save credentials for auto-login
                this.credentials = {
                    imei: event.data.imei,
                    cookie: event.data.cookie,
                    userAgent: event.data.userAgent,
                };
                fs.writeFileSync(
                    CREDENTIALS_PATH,
                    JSON.stringify(this.credentials, null, 2),
                    "utf-8"
                );
                console.log("[ZaloClient] Credentials saved!");
                break;
        }
    }

    /**
     * Logout and clean up
     */
    async logout() {
        if (this.api && this.api.listener) {
            try { this.api.listener.stop(); } catch {}
        }
        this.api = null;
        this.isLoggedIn = false;
        this.loginInProgress = false;
        this.qrStatus = "idle";
        this.qrImageBase64 = null;
        this.qrUserInfo = null;
        this.credentials = null;

        // Remove saved credentials
        try { fs.unlinkSync(CREDENTIALS_PATH); } catch {}
        try { fs.unlinkSync(QR_IMAGE_PATH); } catch {}

        console.log("[ZaloClient] Logged out");
    }

    /**
     * Get all groups the user is in
     */
    async getGroups() {
        if (!this.isLoggedIn || !this.api) {
            throw new Error("Not logged in");
        }

        try {
            // Get all group IDs first
            const allGroups = await this.api.getAllGroups();
            const groupIds = Object.keys(allGroups.gridVerMap || {});

            if (groupIds.length === 0) {
                return [];
            }

            // Get group info for all groups
            const groupInfo = await this.api.getGroupInfo(groupIds);
            const groups = [];

            if (groupInfo && groupInfo.gridInfoMap) {
                for (const [id, info] of Object.entries(groupInfo.gridInfoMap)) {
                    groups.push({
                        id: id,
                        name: info.name || "Unknown",
                        totalMember: info.totalMember || 0,
                        avatar: info.avt || "",
                        desc: info.desc || "",
                    });
                }
            }

            // Sort by name
            groups.sort((a, b) => a.name.localeCompare(b.name, "vi"));
            return groups;
        } catch (err) {
            console.error("[ZaloClient] Error getting groups:", err.message);
            throw err;
        }
    }

    /**
     * Get chat history from a group and extract image URLs
     * @param {string} groupId
     * @param {number} count - number of messages to fetch
     * @returns {Array} - array of image info objects
     */
    async getGroupImages(groupId, count = 200) {
        if (!this.isLoggedIn || !this.api) {
            throw new Error("Not logged in");
        }

        try {
            const history = await this.api.getGroupChatHistory(groupId, count);
            const images = [];

            if (history && history.groupMsgs) {
                for (const msg of history.groupMsgs) {
                    const msgData = msg.data || msg;
                    const content = msgData.content;

                    // Check if message contains image
                    if (content && typeof content === "object") {
                        // Image messages have href/thumb properties
                        if (content.href && this._isImageUrl(content.href)) {
                            images.push({
                                msgId: msgData.msgId || msgData.cliMsgId,
                                url: content.href,
                                thumb: content.thumb || "",
                                timestamp: parseInt(msgData.ts) || Date.now(),
                                sender: msgData.uidFrom || "",
                                senderName: msgData.dName || "",
                            });
                        }
                        // Some image messages have params with URL
                        if (content.params) {
                            try {
                                const params = typeof content.params === "string"
                                    ? JSON.parse(content.params)
                                    : content.params;
                                if (params.hdUrl || params.url || params.normalUrl) {
                                    const imgUrl = params.hdUrl || params.normalUrl || params.url;
                                    images.push({
                                        msgId: msgData.msgId || msgData.cliMsgId,
                                        url: imgUrl,
                                        thumb: params.thumbUrl || content.thumb || "",
                                        timestamp: parseInt(msgData.ts) || Date.now(),
                                        sender: msgData.uidFrom || "",
                                        senderName: msgData.dName || "",
                                    });
                                }
                            } catch {}
                        }
                    }

                    // Check msgType for image type (msgType that contain photos)
                    if (msgData.msgType === "chat.photo" || msgData.msgType === "webchat.photo") {
                        if (typeof content === "string") {
                            try {
                                const parsed = JSON.parse(content);
                                if (parsed.hdUrl || parsed.url || parsed.normalUrl || parsed.oriUrl) {
                                    const imgUrl = parsed.oriUrl || parsed.hdUrl || parsed.normalUrl || parsed.url;
                                    images.push({
                                        msgId: msgData.msgId || msgData.cliMsgId,
                                        url: imgUrl,
                                        thumb: parsed.thumbUrl || "",
                                        timestamp: parseInt(msgData.ts) || Date.now(),
                                        sender: msgData.uidFrom || "",
                                        senderName: msgData.dName || "",
                                    });
                                }
                            } catch {}
                        }
                    }
                }
            }

            return images;
        } catch (err) {
            console.error("[ZaloClient] Error getting group images:", err.message);
            throw err;
        }
    }

    /**
     * Check if URL looks like an image
     */
    _isImageUrl(url) {
        if (!url) return false;
        const lower = url.toLowerCase();
        return (
            lower.includes(".jpg") ||
            lower.includes(".jpeg") ||
            lower.includes(".png") ||
            lower.includes(".bmp") ||
            lower.includes(".webp") ||
            lower.includes("photo") ||
            lower.includes("image")
        );
    }

    /**
     * Get connection status
     */
    getStatus() {
        return {
            isLoggedIn: this.isLoggedIn,
            loginInProgress: this.loginInProgress,
            qrStatus: this.qrStatus,
            qrUserInfo: this.qrUserInfo,
            hasQrImage: !!this.qrImageBase64,
        };
    }
}

module.exports = ZaloClient;
