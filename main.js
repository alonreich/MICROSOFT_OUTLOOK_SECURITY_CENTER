const path = require('node:path');
const fs = require('node:fs');
const fsPromises = fs.promises;
const net = require('node:net');
const crypto = require('node:crypto');
const { spawn, execFile } = require('node:child_process');
const electron = require('electron');
const { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, safeStorage, shell, dialog } = electron;

const isServiceMode = process.argv.includes('--service');
const isHiddenMode = process.argv.includes('--hidden') || process.argv.includes('--minimized');
const APP_ROOT = __dirname;
const LOG_DIR = path.join(APP_ROOT, 'logs');
const FORENSICS_DIR = path.join(LOG_DIR, 'forensics');
const LOG_FILE = path.join(LOG_DIR, 'deskguard_outlook.log');

const USER_DATA = app.getPath('userData');
const vaultDB = require('./vault_db');
try {
    vaultDB.init(USER_DATA);
} catch (err) {
    console.error('Failed to initialize DeskGuard Vault DB:', err.message);
}
const PIPE_NAME = `\\\\.\\pipe\\mos_service_${crypto.createHash('sha256').update(USER_DATA).digest('hex').slice(0, 12)}`;
const PIPE_AUTH_TOKEN = `MOS_AUTH_${crypto.createHash('sha256').update(USER_DATA).digest('hex').slice(0, 16)}`;

function applyWindowsStartupSetting(enabled) {
    const regKeyPath = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run';
    const keyName = 'DeskGuardForMicrosoftOutlook';

    try {
        execFile('reg', ['delete', regKeyPath, '/v', 'MicrosoftOutlookSecurityCenter', '/f'], () => {});
        execFile('reg', ['delete', regKeyPath, '/v', 'electron.app.Electron', '/f'], () => {});
    } catch {}

    if (enabled) {
        let launchCmd;
        if (app.isPackaged) {
            launchCmd = `"${process.execPath}" --hidden`;
        } else {
            launchCmd = `"${process.execPath}" "${APP_ROOT}" --hidden`;
        }

        try {
            execFile('reg', ['add', regKeyPath, '/v', keyName, '/t', 'REG_SZ', '/d', launchCmd, '/f'], (err) => {
                if (err) {
                    logToFile(`[Startup Registry] Failed to add Run key: ${err.message}`, 'ERROR');
                } else {
                    logToFile(`[Startup Registry] Successfully configured Windows Run key: ${launchCmd}`, 'INFO');
                }
            });
        } catch (e) {
            logToFile(`[Startup Registry] Exception configuring Run key: ${e.message}`, 'ERROR');
        }

        try {
            app.setLoginItemSettings({
                openAtLogin: true,
                path: process.execPath,
                args: app.isPackaged ? ['--hidden'] : [APP_ROOT, '--hidden']
            });
        } catch {}
    } else {
        try {
            execFile('reg', ['delete', regKeyPath, '/v', keyName, '/f'], (err) => {
                if (!err) {
                    logToFile(`[Startup Registry] Successfully removed Windows Run key: ${keyName}`, 'INFO');
                }
            });
        } catch (e) {
            logToFile(`[Startup Registry] Exception removing Run key: ${e.message}`, 'ERROR');
        }

        try {
            app.setLoginItemSettings({
                openAtLogin: false
            });
        } catch {}
    }
}

function killProcessTree(child) {
    if (!child) return;
    const pid = child.pid;
    try { child.removeAllListeners('exit'); } catch {}
    if (pid) {
        try {
            const { execSync } = require('node:child_process');
            execSync(`taskkill /pid ${pid} /T /F 2>nul`);
        } catch {}
    }
    try { if (!child.killed) child.kill('SIGKILL'); } catch {}
}

function ensureOutlookProgrammaticAccessPolicies() {
    const versions = ['16.0', '15.0', '14.0'];
    const settings = [
        ['PromptOOMSend', '2'],
        ['PromptOOMAddressBookAccess', '2'],
        ['PromptOOMAddressInformationAccess', '2'],
        ['PromptOOMSaveAs', '2'],
        ['AdminSecurityMode', '3']
    ];

    versions.forEach(v => {
        const paths = [
            `HKCU\\Software\\Microsoft\\Office\\${v}\\Outlook\\Security`,
            `HKCU\\Software\\Policies\\Microsoft\\Office\\${v}\\Outlook\\Security`
        ];
        paths.forEach(p => {
            settings.forEach(([name, val]) => {
                try {
                    execFile('reg', ['add', p, '/v', name, '/t', 'REG_DWORD', '/d', val, '/f'], () => {});
                } catch {}
            });
        });
    });
}

[LOG_DIR, FORENSICS_DIR].forEach(d => { if (!fs.existsSync(d)) try { fs.mkdirSync(d, { recursive: true }); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } } });
if (!fs.existsSync(LOG_FILE)) try { fs.writeFileSync(LOG_FILE, `[${new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '')}] DeskGuard: Initialization Success. Monitoring is active.\n`); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }

async function logToFile(msg, level = "INFO") {
    const ts = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
    const logLine = `[${ts}] [${level}] ${msg}\n`;
    try { 
        const exists = fs.existsSync(LOG_FILE);
        if (exists) {
            const stats = await fsPromises.stat(LOG_FILE).catch(() => null);
            if (stats && stats.size > 10 * 1024 * 1024) {
                await fsPromises.rename(LOG_FILE, LOG_FILE + '.1').catch(() => {});
            }
        }
        await fsPromises.appendFile(LOG_FILE, logLine).catch(() => {});
        broadcastToUi({ type: 'live-log', message: `[${level}] ${msg}` });
    } catch(err) { console.error(err); }
}

const Store = require('electron-store');

const DEFAULT_CONFIG = { 
    enabled: true, 
    firstRun: true,
    vtApiKey: '', 
    spamKeywords: ['viagra', 'lottery', 'urgent', 'bitcoin', 'winner', 'unpaid', 'invoice', 'payment', 'account', 'verify', 'security', 'update', 'action', 'urgent-action', 'account-compromise', 'limited-access', 'security-alert', 'suspicious-activity'], 
    rubrics: { 
        weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }, 
        toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }, 
        spamThresholdPercent: 50 
    }, 
    whitelist: { emails: [], ips: [], domains: [], combos: [] }, 
    blacklist: { emails: [], ips: [], domains: [], combos: [] },
    launchAtStartup: true,
    scanningSpeed: 50,
    historyScanEnabled: true,
    onAccessEnabled: true,
    onDemandLimit: 1000,
    deepHistoryScanEnabled: false,
    threatIntelligenceLevel: 1
};

const DEFAULT_DATA = {
    processedIds: [],
    releasedFingerprints: [],
    stats: { spam: [], safe: [], malicious: [], suspicious: [] }
};

function recoverCorruptStoreFile(storeName, defaultObj) {
    const filePath = path.join(USER_DATA, `${storeName}.json`);
    const bakPath = path.join(USER_DATA, `${storeName}.json.bak`);

    if (!fs.existsSync(filePath)) {
        if (fs.existsSync(bakPath)) {
            try {
                const bakContent = fs.readFileSync(bakPath, 'utf8');
                JSON.parse(bakContent);
                fs.copyFileSync(bakPath, filePath);
                logToFile(`[Storage Recovery] Missing ${storeName}.json restored from ${storeName}.json.bak.`, 'INFO');
                return;
            } catch (err) {
                logToFile(`[Storage Recovery] Backup for missing ${storeName}.json is corrupt: ${err.message}`, 'WARN');
            }
        }
        return;
    }

    let needsRecovery = false;
    let parseErrorMsg = '';
    try {
        const content = fs.readFileSync(filePath, 'utf8');
        if (!content || content.trim().length === 0) {
            needsRecovery = true;
            parseErrorMsg = 'File is empty';
        } else {
            JSON.parse(content);
        }
    } catch (err) {
        needsRecovery = true;
        parseErrorMsg = err.message;
    }

    if (!needsRecovery) {
        try {
            const content = fs.readFileSync(filePath, 'utf8');
            fs.writeFileSync(bakPath, content, 'utf8');
        } catch (err) {
            logToFile(`[Storage Backup] Initial backup creation warning for ${storeName}: ${err.message}`, 'WARN');
        }
        return;
    }

    logToFile(`[Storage Recovery] Corrupt ${storeName}.json detected (${parseErrorMsg}). Attempting backup restore...`, 'WARN');
    let restored = false;
    if (fs.existsSync(bakPath)) {
        try {
            const bakContent = fs.readFileSync(bakPath, 'utf8');
            if (bakContent && bakContent.trim().length > 0) {
                JSON.parse(bakContent);
                fs.writeFileSync(filePath, bakContent, 'utf8');
                logToFile(`[Storage Recovery] Successfully restored ${storeName}.json from ${storeName}.json.bak.`, 'INFO');
                restored = true;
            }
        } catch (err) {
            logToFile(`[Storage Recovery] Corrupt backup ${storeName}.json.bak encountered: ${err.message}`, 'ERROR');
        }
    }

    if (!restored) {
        logToFile(`[Storage Recovery] Unrecoverable ${storeName}.json. Initializing with safe defaults.`, 'WARN');
        try {
            fs.writeFileSync(filePath, JSON.stringify(defaultObj, null, 4), 'utf8');
            fs.writeFileSync(bakPath, JSON.stringify(defaultObj, null, 4), 'utf8');
        } catch (err) {
            logToFile(`[Storage Recovery] Error writing defaults for ${storeName}: ${err.message}`, 'ERROR');
        }
    }
}

function createStoreInstance(storeName, defaultObj) {
    recoverCorruptStoreFile(storeName, defaultObj);
    let storeInstance;
    try {
        storeInstance = new Store({
            cwd: USER_DATA,
            name: storeName,
            clearInvalidConfig: false,
            defaults: defaultObj
        });
    } catch (err) {
        logToFile(`[Storage Error] Failed to initialize ${storeName} store: ${err.message}. Retrying with backup recovery...`, 'ERROR');
        const bakPath = path.join(USER_DATA, `${storeName}.json.bak`);
        const filePath = path.join(USER_DATA, `${storeName}.json`);
        if (fs.existsSync(bakPath)) {
            try {
                const bak = fs.readFileSync(bakPath, 'utf8');
                JSON.parse(bak);
                fs.writeFileSync(filePath, bak, 'utf8');
            } catch {}
        }
        storeInstance = new Store({
            cwd: USER_DATA,
            name: storeName,
            clearInvalidConfig: false,
            defaults: defaultObj
        });
    }
    backupStoreBeforeWrite(storeName);
    return storeInstance;
}

let configStore = null;

if (isServiceMode) {
    configStore = createStoreInstance('config', DEFAULT_CONFIG);
}

class AsyncWriteQueue {
    constructor() {
        this.queue = [];
        this.isProcessing = false;
    }

    enqueue(operation) {
        return new Promise((resolve, reject) => {
            this.queue.push({ operation, resolve, reject });
            this.processNext();
        });
    }

    async processNext() {
        if (this.isProcessing || this.queue.length === 0) return;
        this.isProcessing = true;
        const { operation, resolve, reject } = this.queue.shift();
        try {
            const result = await operation();
            resolve(result);
        } catch (err) {
            reject(err);
        } finally {
            this.isProcessing = false;
            this.processNext();
        }
    }
}

const serviceWriteQueue = new AsyncWriteQueue();

function backupStoreBeforeWrite(storeName) {
    try {
        const filePath = path.join(USER_DATA, `${storeName}.json`);
        const bakPath = path.join(USER_DATA, `${storeName}.json.bak`);
        if (fs.existsSync(filePath)) {
            const content = fs.readFileSync(filePath, 'utf8');
            if (content && content.trim().length > 0) {
                JSON.parse(content);
                fs.writeFileSync(bakPath, content, 'utf8');
            }
        }
    } catch (err) {
        logToFile(`[Storage Backup] Warning: Could not create backup before write for ${storeName}: ${err.message}`, 'WARN');
    }
}

async function serviceSetStore(key, value) {
    if (!isServiceMode) return;
    return serviceWriteQueue.enqueue(async () => {
        if (key === 'processedIds') {
            if (Array.isArray(value)) {
                processedIdsCache.clear();
                value.slice(-MAX_PROCESSED_IDS).forEach(id => processedIdsCache.add(id));
                uncommittedProcessedIdsCount = processedIdsCache.size;
                isPersistenceDirty = true;
                scheduleStatePersistenceFlush(true);
            }
            return;
        }
        if (key === 'stats') {
            statsCache = value;
            isPersistenceDirty = true;
            scheduleStatePersistenceFlush(true);
            return;
        }
        if (key === 'releasedFingerprints') {
            if (Array.isArray(value)) {
                value.forEach(fp => { if (fp) releasedFingerprintsCache.add(fp); });
            } else if (typeof value === 'string' && value) {
                releasedFingerprintsCache.add(value);
            }
            while (releasedFingerprintsCache.size > 5000) {
                const oldest = releasedFingerprintsCache.values().next().value;
                if (oldest === undefined) break;
                releasedFingerprintsCache.delete(oldest);
            }
            isPersistenceDirty = true;
            scheduleStatePersistenceFlush(false);
            return;
        }

        if (!configStore) {
            logToFile(`[Storage Error] Attempted write to uninitialized configStore: ${key}`, 'ERROR');
            return;
        }

        try {
            backupStoreBeforeWrite('config');
            configStore.set(key, value);
        } catch (err) {
            const errMsg = `Disk write exception for config [${key}]: ${err.message}`;
            logToFile(errMsg, 'ERROR');
            broadcastToUi({
                type: 'storage-error',
                store: 'config',
                key,
                error: err.message
            });
            return;
        }

        broadcastToUi({
            type: 'status-sync',
            enabled: !!configStore.get('enabled'),
            stats: statsCache,
            config: configStore.store
        });

        if (key === 'enabled' || key === 'historyScanEnabled') {
            if (configStore.get('enabled')) {
                requestScannerRestart(key);
            } else {
                if (standbyWatcherTimer) {
                    clearInterval(standbyWatcherTimer);
                    standbyWatcherTimer = null;
                }
                if (currentScanChild) {
                    killProcessTree(currentScanChild);
                    currentScanChild = null;
                    isScanning = false;
                }
                broadcastToUi({ type: 'outlook-status', running: false, standby: false });
            }
        } else if (key === 'launchAtStartup') {
            applyWindowsStartupSetting(!!value);
        } else if (key === 'scanningSpeed') {
            if (currentScanChild) {
                try {
                    currentScanChild.stdin.write(JSON.stringify({ type: 'config-update', scanningSpeed: value }) + '\n');
                } catch (e) {}
            }
        } else if (['rubrics', 'spamKeywords', 'whitelist', 'blacklist', 'vtApiKey'].includes(key)) {
            requestScannerRestart(key);
        }
    });
}

function getDecryptedVtKey(key) {
    if (!key) return '';
    try {
        return safeStorage.decryptString(Buffer.from(key, 'base64'));
    } catch {
        return key;
    }
}

class SafeIPCParser {
    constructor(onMessage) {
        this.buf = Buffer.alloc(0);
        this.onMessage = onMessage;
        this.nextLen = -1;
    }
    push(data) {
        this.buf = Buffer.concat([this.buf, data]);
        while (true) {
            if (this.nextLen === -1) {
                if (this.buf.length < 4) break;
                this.nextLen = this.buf.readInt32LE(0);
                this.buf = this.buf.subarray(4);

                // Framing validation: guard against corrupt/malformed packet headers
                if (this.nextLen <= 0 || this.nextLen > 10485760) {
                    const errDetail = `IPC Desync Error: Invalid packet length header (${this.nextLen}). Purging buffer.`;
                    console.error(errDetail);
                    logToFile(errDetail, 'ERROR');
                    this.buf = Buffer.alloc(0);
                    this.nextLen = -1;
                    break;
                }
            }
            if (this.buf.length < this.nextLen) break;
            const msg = this.buf.subarray(0, this.nextLen).toString('utf8');
            this.buf = this.buf.subarray(this.nextLen);
            this.nextLen = -1;
            try {
                const parsed = JSON.parse(msg);
                if (parsed) this.onMessage(parsed);
            } catch (err) {
                console.error("IPC Parse Error:", err, "Raw:", msg);
                logToFile("IPC Parse Error: " + (err && err.message ? err.message : String(err)), "ERROR");
            }
        }
    }
}

const MAX_PROCESSED_IDS = 100000;
const MAX_STATS_PER_CAT = 5000;

function loadInitialConfig() {
    try {
        const configPath = path.join(USER_DATA, 'config.json');
        if (fs.existsSync(configPath)) {
            const raw = fs.readFileSync(configPath, 'utf8');
            const parsed = JSON.parse(raw);
            return { ...DEFAULT_CONFIG, ...parsed };
        }
    } catch (e) {}
    return { ...DEFAULT_CONFIG };
}

function loadInitialData() {
    const dataPath = path.join(USER_DATA, 'data.json');
    const bakPath = path.join(USER_DATA, 'data.json.bak');
    let rawData = null;
    let loadedFromBak = false;

    if (fs.existsSync(dataPath)) {
        try {
            const content = fs.readFileSync(dataPath, 'utf8');
            if (content && content.trim().length > 0) {
                rawData = JSON.parse(content);
            }
        } catch (err) {
            logToFile(`[State Persistence] Warning: data.json read error (${err.message}), falling back to backup.`, 'WARN');
        }
    }

    if (!rawData && fs.existsSync(bakPath)) {
        try {
            const bakContent = fs.readFileSync(bakPath, 'utf8');
            if (bakContent && bakContent.trim().length > 0) {
                rawData = JSON.parse(bakContent);
                loadedFromBak = true;
                logToFile(`[State Persistence] Successfully restored state from data.json.bak.`, 'INFO');
            }
        } catch (err) {
            logToFile(`[State Persistence Error] data.json.bak read error: ${err.message}`, 'ERROR');
        }
    }

    let initialData = rawData || { ...DEFAULT_DATA };
    if (!initialData.stats || typeof initialData.stats !== 'object') {
        initialData.stats = { malicious: [], suspicious: [], spam: [], safe: [] };
    }

    for (const cat of ['malicious', 'suspicious', 'spam', 'safe']) {
        if (!Array.isArray(initialData.stats[cat])) {
            initialData.stats[cat] = [];
        } else {
            initialData.stats[cat].forEach(item => {
                if (!item.subject && item.details) item.subject = item.details;
                if (!item.details && item.subject) item.details = item.subject;
                if (item.userMoved && !item.userMovedStory) {
                    item.userMovedStory = `Originally quarantined by DeskGuard, but you manually moved this email to folder '${item.currentFolder || 'another folder'}' inside Outlook. DeskGuard respects your choice and will keep it here.`;
                }
            });
        }
    }

    // VaultDB Fallback: If in-memory count is 0, attempt restoring from SQLite Vault DB
    const totalInMemory = Object.values(initialData.stats).reduce((sum, arr) => sum + (arr ? arr.length : 0), 0);
    if (totalInMemory === 0 && vaultDB && typeof vaultDB.getAllCategorizedEmails === 'function') {
        try {
            const vaultItems = vaultDB.getAllCategorizedEmails(2000);
            const vaultCount = Object.values(vaultItems).reduce((sum, arr) => sum + (arr ? arr.length : 0), 0);
            if (vaultCount > 0) {
                initialData.stats = vaultItems;
                logToFile(`[State Persistence] Restored ${vaultCount} scanned emails from SQLite Vault DB.`, 'INFO');
            }
        } catch (vErr) {
            logToFile(`[State Persistence] Vault DB fallback error: ${vErr.message}`, 'WARN');
        }
    }

    logToFile(`[State Persistence] Initialized state: ${(initialData.processedIds || []).length} processed IDs, ${(initialData.releasedFingerprints || []).length} released fingerprints, ${Object.values(initialData.stats).reduce((s, a) => s + (a ? a.length : 0), 0)} categorized incidents.`);
    return initialData;
}

const initialSavedData = loadInitialData();

let mainWindow = null, tray = null, isQuitting = false;
let configCache = loadInitialConfig();
let isEnabled = configCache.enabled !== undefined ? !!configCache.enabled : true;
let uiPipeClient = null, serviceSession = null, serviceSpawnInFlight = false;
let pipeServer = null, activeConnections = new Set(), isScanning = false, currentScanChild = null;
let psWorker = null;
let statsBuffer = { malicious: [], suspicious: [], spam: [], safe: [] };
let bufferTimer = null;
let watchdogTimer = null;
let lastHeartbeat = Date.now();
let statsCache = initialSavedData.stats;

function broadcastToUi(msg) {
    if (isServiceMode) {
        const raw = JSON.stringify(msg) + "\n";
        activeConnections.forEach(s => { try { s.write(raw); } catch(e){} });
    } else if (mainWindow && mainWindow.webContents) {
        if (msg.type === 'scan-update') {
            mainWindow.webContents.send("outlook-scan-update", msg.data);
            if (msg.data && msg.data.status) {
                const s = msg.data.status;
                const d = msg.data.details || 'Unknown Item';
                const snd = msg.data.sender ? ` [From: ${msg.data.sender}]` : "";
                const eng = msg.data.tier ? ` [Policy: ${msg.data.tier}]` : "";
                
                if (s === 'THREAT BLOCKED') {
                    logToFile(`CRITICAL: Malicious Activity Prevented on "${d}"${snd}${eng}`, 'WARN');
                } else if (s === 'SPAM FILTERED') {
                    logToFile(`Filter: Junk Email Isolated: "${d}"${snd}${eng}`, 'INFO');
                } else if (s === 'Finished' && d.startsWith('Audit complete')) {
                    logToFile(`Success: ${d}`, 'INFO');
                } else if (s === 'Finished') {
                    if (msg.data.tier) {
                        logToFile(`Audit: Verified Safe Item: "${d}"${snd}${eng}`, 'INFO');
                    }
                } else if (s === 'SCANNING') {
                    if (msg.data.count % 50 === 0) {
                        logToFile(`Status: ${d}`, 'INFO');
                    }
                } else if (s === 'INFO') {
                    if (d.startsWith('Forensics:')) {
                        logToFile(`[Forensic Engine] ${d.replace('Forensics:', '').trim()}`, 'INFO');
                    } else if (d.startsWith('TRACE')) {
                        logToFile(`[Analysis Trace] ${d.replace('TRACE', '').trim()}`, 'INFO');
                    } else if (d.startsWith('Analysis:')) {
                        logToFile(`[Security Logic] ${d.replace('Analysis:', '').trim()}`, 'INFO');
                    } else {
                        logToFile(`Engine: ${d}`, 'INFO');
                    }
                }
            }
        }
        else if (msg.type === 'status-sync') mainWindow.webContents.send("status-sync", msg.enabled);
        else if (msg.type === 'stats-update') mainWindow.webContents.send("stats-update", msg.data);
        else if (msg.type === 'live-log') mainWindow.webContents.send("live-log", msg.message);
        else if (msg.type === 'outlook-status') mainWindow.webContents.send("outlook-status", { running: !!msg.running, standby: !!msg.standby });
        else if (msg.type === 'duplicate-update') mainWindow.webContents.send("duplicate-update", msg);
        else mainWindow.webContents.send("from-main", msg);
    }
}

// IN-MEMORY STATE PERSISTENCE & TELEMETRY INGESTION ENGINE
const processedIdsCache = new Set(initialSavedData.processedIds || []);
const releasedFingerprintsCache = new Set(initialSavedData.releasedFingerprints || []);
let uncommittedProcessedIdsCount = 0;
let isPersistenceDirty = false;
let flushTimer = null;
let isFlushInProgress = false;
let flushPending = false;
let flushRetryCount = 0;
const FLUSH_INTERVAL_MS = 5000;
const MAX_UNCOMMITTED_ITEMS = 500;

function addProcessedId(fid) {
    if (!fid) return false;
    if (processedIdsCache.has(fid)) return false;

    processedIdsCache.add(fid);
    uncommittedProcessedIdsCount++;
    isPersistenceDirty = true;

    while (processedIdsCache.size > MAX_PROCESSED_IDS) {
        const oldest = processedIdsCache.values().next().value;
        if (oldest === undefined) break;
        processedIdsCache.delete(oldest);
    }

    scheduleStatePersistenceFlush();
    return true;
}

function scheduleStatePersistenceFlush(immediate = false) {
    if (!isServiceMode) return;

    if (immediate || uncommittedProcessedIdsCount >= MAX_UNCOMMITTED_ITEMS) {
        if (flushTimer) {
            clearTimeout(flushTimer);
            flushTimer = null;
        }
        flushStatePersistence();
        return;
    }

    if (!flushTimer) {
        flushTimer = setTimeout(() => {
            flushTimer = null;
            flushStatePersistence();
        }, FLUSH_INTERVAL_MS);
    }
}

function scheduleProcessedIdsFlush(immediate = false) {
    scheduleStatePersistenceFlush(immediate);
}

function flushStats() {
    return flushStatePersistence();
}

async function flushStatePersistence() {
    if (!isServiceMode) return;

    const hasStatsData = Object.values(statsBuffer).some(a => a.length > 0);
    const hasIdsData = uncommittedProcessedIdsCount > 0;

    if (!hasStatsData && !hasIdsData && !isPersistenceDirty) return;

    if (isFlushInProgress) {
        flushPending = true;
        return;
    }

    isFlushInProgress = true;
    flushPending = false;

    try {
        let statsUpdated = false;
        let currentStats = statsCache || { malicious: [], suspicious: [], spam: [], safe: [] };

        if (hasStatsData) {
            const newFingerprints = new Set();
            for (const cat in statsBuffer) {
                statsBuffer[cat].forEach(item => {
                    const fid = item.fingerprint || item.entryId || item.originalEntryId;
                    if (fid) newFingerprints.add(fid);
                });
            }
            for (const cat in currentStats) {
                if (!currentStats[cat]) currentStats[cat] = [];
                currentStats[cat] = currentStats[cat].filter(item => {
                    const fid = item.fingerprint || item.entryId || item.originalEntryId;
                    return !newFingerprints.has(fid);
                });
            }
            for (const cat in statsBuffer) {
                if (!currentStats[cat]) currentStats[cat] = [];
                const combined = [...currentStats[cat], ...statsBuffer[cat]];
                const seen = new Set();
                currentStats[cat] = combined.filter(item => {
                    const fid = item.fingerprint || item.entryId || item.originalEntryId;
                    if (!fid || seen.has(fid)) return false;
                    seen.add(fid); return true;
                }).slice(-MAX_STATS_PER_CAT);
                statsBuffer[cat] = [];
            }
            statsCache = currentStats;
            statsUpdated = true;
        }

        const dataPath = path.join(USER_DATA, 'data.json');
        const tmpPath = path.join(USER_DATA, 'data.json.tmp');
        const bakPath = path.join(USER_DATA, 'data.json.bak');

        const fullPayload = {
            processedIds: Array.from(processedIdsCache),
            releasedFingerprints: Array.from(releasedFingerprintsCache),
            stats: statsCache
        };

        const jsonString = JSON.stringify(fullPayload, null, 2);

        // Atomic asynchronous disk write: write to tmp file first
        await fsPromises.writeFile(tmpPath, jsonString, 'utf8');

        // Backup existing data.json
        if (fs.existsSync(dataPath)) {
            await fsPromises.copyFile(dataPath, bakPath).catch(() => {});
        }

        // Atomic swap via rename
        await fsPromises.rename(tmpPath, dataPath);

        uncommittedProcessedIdsCount = 0;
        isPersistenceDirty = false;
        flushRetryCount = 0;

        if (statsUpdated) {
            broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
        }
    } catch (err) {
        flushRetryCount++;
        const backoffMs = Math.min(30000, 1000 * Math.pow(2, flushRetryCount));
        logToFile(`[State Persistence Error] Flush failed (attempt ${flushRetryCount}): ${err.message}. Retrying in ${backoffMs}ms...`, 'WARN');
        const tmpPath = path.join(USER_DATA, 'data.json.tmp');
        await fsPromises.unlink(tmpPath).catch(() => {});
        if (!flushTimer) {
            flushTimer = setTimeout(() => {
                flushTimer = null;
                flushStatePersistence();
            }, backoffMs);
        }
    } finally {
        isFlushInProgress = false;
        if (flushPending) {
            setImmediate(() => flushStatePersistence());
        }
    }
}

function flushStatePersistenceSync() {
    if (!isServiceMode) return;
    try {
        if (flushTimer) {
            clearTimeout(flushTimer);
            flushTimer = null;
        }
        const hasStatsData = Object.values(statsBuffer).some(a => a.length > 0);
        let currentStats = statsCache || { malicious: [], suspicious: [], spam: [], safe: [] };
        if (hasStatsData) {
            for (const cat in statsBuffer) {
                if (!currentStats[cat]) currentStats[cat] = [];
                const combined = [...currentStats[cat], ...statsBuffer[cat]];
                const seen = new Set();
                currentStats[cat] = combined.filter(item => {
                    const fid = item.fingerprint || item.entryId || item.originalEntryId;
                    if (!fid || seen.has(fid)) return false;
                    seen.add(fid); return true;
                }).slice(-MAX_STATS_PER_CAT);
                statsBuffer[cat] = [];
            }
            statsCache = currentStats;
        }

        const dataPath = path.join(USER_DATA, 'data.json');
        const tmpPath = path.join(USER_DATA, 'data.json.tmp');
        const bakPath = path.join(USER_DATA, 'data.json.bak');

        const fullPayload = {
            processedIds: Array.from(processedIdsCache),
            releasedFingerprints: Array.from(releasedFingerprintsCache),
            stats: currentStats
        };

        const jsonString = JSON.stringify(fullPayload, null, 2);
        fs.writeFileSync(tmpPath, jsonString, 'utf8');
        if (fs.existsSync(dataPath)) {
            try { fs.copyFileSync(dataPath, bakPath); } catch {}
        }
        fs.renameSync(tmpPath, dataPath);
        uncommittedProcessedIdsCount = 0;
        isPersistenceDirty = false;
        logToFile('[State Persistence] Clean shutdown flush completed.', 'INFO');
    } catch (err) {
        logToFile(`[State Persistence Error] Shutdown flush failed: ${err.message}`, 'ERROR');
    }
}

async function cleanupForensics() {
    try {
        const stats = statsCache || { malicious: [], suspicious: [], spam: [], safe: [] };
        const activeFingerprints = new Set();
        const cats = ['malicious', 'suspicious', 'spam', 'safe'];
        cats.forEach(cat => {
            (stats[cat] || []).concat(statsBuffer[cat] || []).forEach(item => {
                if (item.entryId) activeFingerprints.add(crypto.createHash('sha256').update(String(item.entryId)).digest('hex'));
                if (item.originalEntryId) activeFingerprints.add(crypto.createHash('sha256').update(String(item.originalEntryId)).digest('hex'));
                if (item.fingerprint) activeFingerprints.add(crypto.createHash('sha256').update(String(item.fingerprint)).digest('hex'));
            });
        });
        const files = await fsPromises.readdir(FORENSICS_DIR);
        for (const f of files) {
            const hash = f.replace('.json', '');
            if (!activeFingerprints.has(hash)) {
                await fsPromises.unlink(path.join(FORENSICS_DIR, f)).catch(() => {});
            }
        }
    } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }
}

function getPsWorker() {
    if (psWorker && !psWorker.killed && psWorker.exitCode === null) return psWorker;
    logToFile('Spawning Security Engine Worker process...');
    psWorker = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-Mode', 'Worker', '-ParentPid', process.pid.toString()], { windowsHide: true });
    const parser = new SafeIPCParser(p => {
                if (p.type === 'cmd-response') {
                    if (p.rid && reqHandlers.has(p.rid)) {
                        const resolve = reqHandlers.get(p.rid);
                        reqHandlers.delete(p.rid);
                        resolve(p.data !== undefined ? p.data : p);
                    }
                    broadcastToUi(p);
                    return;
                }
                if (p.type === 'store-update') {
                    if (p.key === 'releasedFingerprints' && p.value) {
                        if (!releasedFingerprintsCache.has(p.value)) {
                            serviceSetStore('releasedFingerprints', p.value);
                        }
                    }
                    return;
                }
                if (p.type === 'store-data' || p.type === 'duplicate-update' || p.type === 'delete-summary') {
                    if (p.type === 'duplicate-update') {
                        if (p.status === 'Progress') logToFile(`Duplicate Scan: ${p.details}`);
                        else if (p.status === 'Found') logToFile(`Duplicate Found: ${p.current}`);
                        else if (p.status === 'Finished') logToFile(`Duplicate Scan Finished. Found ${p.items ? p.items.length : 0} items.`);
                        else if (p.status === 'Paused') logToFile(`Duplicate Scan Paused by user.`);
                    }
                    if (p.type === 'delete-summary') {
                        logToFile(`Cleanup: Permanently deleted ${p.count} redundant email copies.`);
                    }
                    broadcastToUi(p); 
                    return;
                }
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'MONITORING', 'INFO', 'ERROR'].includes(p.status)) {
                    if (p.status === 'INFO' || p.status === 'MONITORING') {
                        logToFile(`Worker: ${p.details || ''}`);
                    } else if (p.status === 'ERROR') {
                        logToFile(`Worker Error: ${p.details || ''}`, 'ERROR');
                    } else {
                        logToFile(`Worker Result [${p.status}]: ${p.details || ''}`);
                    }
                    broadcastToUi({ type: 'scan-update', data: p }); } });
    psWorker.stdout.on('data', d => parser.push(d));
    psWorker.on('exit', (code) => {
        logToFile(`Security Engine Worker process exited with code ${code}`);
    });
    return psWorker;
}

async function runDiagnostics() {
    logToFile("--- SYSTEM SELF-DIAGNOSTIC START ---");
    
    // 1. Check PowerShell
    try {
        const { execSync } = require('child_process');
        const policy = execSync('powershell.exe -NoProfile -Command "Get-ExecutionPolicy"').toString().trim();
        logToFile(`PowerShell Execution Policy: ${policy}`);
        if (['Restricted', 'AllSigned'].includes(policy)) {
            logToFile("CRITICAL: PowerShell policy may block security engine. Recommend 'RemoteSigned' or 'Bypass'.", "WARN");
        }
    } catch (e) {
        logToFile("CRITICAL: PowerShell not found or inaccessible.", "ERROR");
    }

    // 2. Check Outlook COM
    const outlookRunning = await new Promise(resolve => {
        execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe'], (err, stdout) => {
            resolve(stdout.toLowerCase().includes('outlook.exe'));
        });
    });
    logToFile(`Outlook Process Status: ${outlookRunning ? 'Running' : 'Not Detected'}`);

    // 3. Check Workspace
    logToFile(`Workspace Root: ${APP_ROOT}`);
    logToFile(`User Data Path: ${USER_DATA}`);
    
    logToFile("--- SYSTEM SELF-DIAGNOSTIC COMPLETE ---");
}

function isOutlookProcessRunning() {
    return new Promise(resolve => {
        execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe', '/NH'], (err, stdout) => {
            if (err || !stdout) {
                resolve(false);
                return;
            }
            resolve(stdout.toLowerCase().includes('outlook.exe'));
        });
    });
}

async function ensureOutlookRunning() {
    const isRunning = await isOutlookProcessRunning();
    return !!isRunning;
}

let standbyWatcherTimer = null;
let isStandbyMode = false;
let consecutiveEngineFailures = 0;
const MAX_ENGINE_FAILURES = 3;
let restartTimer = null;

function startOutlookStandbyWatcher() {
    if (standbyWatcherTimer) clearInterval(standbyWatcherTimer);
    isStandbyMode = true;
    logToFile('[Standby] Outlook is not currently active. Security Engine in passive standby mode.', 'INFO');
    broadcastToUi({ type: 'outlook-status', running: false, standby: true });

    standbyWatcherTimer = setInterval(async () => {
        if (!configStore || !configStore.get('enabled')) {
            clearInterval(standbyWatcherTimer);
            standbyWatcherTimer = null;
            return;
        }
        const running = await ensureOutlookRunning();
        if (running) {
            logToFile('[Standby] Outlook process detected. Activating Security Engine...', 'INFO');
            clearInterval(standbyWatcherTimer);
            standbyWatcherTimer = null;
            isStandbyMode = false;
            consecutiveEngineFailures = 0;
            broadcastToUi({ type: 'outlook-status', running: true, standby: false });
            runOutlookScanner();
        }
    }, 5000);
}

async function runOutlookScanner() {
    if (!isServiceMode) {
        logToFile('Attempted to run scanner in UI mode. Redirecting to service...');
        return;
    }
    if (!configStore || !configStore.get('enabled')) return;
    if (isScanning && currentScanChild && !currentScanChild.killed && currentScanChild.exitCode === null) {
        logToFile('Scanner already running. Skipping duplicate start.');
        return;
    }
    
    const isReady = await ensureOutlookRunning();
    if (!isReady) {
        logToFile('Engine: Microsoft Outlook is not available. Entering Standby mode.', 'WARN');
        broadcastToUi({ type: 'outlook-status', running: false, standby: true });
        isScanning = false;
        startOutlookStandbyWatcher();
        return;
    }

    if (standbyWatcherTimer) {
        clearInterval(standbyWatcherTimer);
        standbyWatcherTimer = null;
    }
    isStandbyMode = false;

    await cleanupForensics();

    isScanning = true;
    lastHeartbeat = Date.now();
    if (watchdogTimer) {
        clearInterval(watchdogTimer);
        watchdogTimer = null;
    }
    watchdogTimer = setInterval(async () => { 
        const idleTime = Date.now() - lastHeartbeat;
        if (idleTime > 90000) { 
            logToFile(`Watchdog: Security Engine unresponsive for ${Math.round(idleTime/1000)}s. Attempting graceful recovery...`, 'WARN'); 
            broadcastToUi({ type: 'outlook-status', running: false, standby: false });
            if (watchdogTimer) {
                clearInterval(watchdogTimer);
                watchdogTimer = null;
            }
            if (currentScanChild) {
                killProcessTree(currentScanChild);
                currentScanChild = null;
            }
            isScanning = false; 
            const stillRunning = await ensureOutlookRunning();
            if (stillRunning) {
                logToFile('Watchdog: Performing hard engine restart.', 'ERROR');
                requestScannerRestart('watchdog-timeout');
            } else {
                startOutlookStandbyWatcher();
            }
        } else {
            broadcastToUi({ type: 'outlook-status', running: true, standby: false });
        }
    }, 10000);

    logToFile('Spawning Security Engine process...');
    currentScanChild = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-ParentPid', process.pid.toString()], { windowsHide: true });
    
    const vtKeyEnc = configStore.get('vtApiKey');
    let vtKeyDec = '';
    if (vtKeyEnc) { try { vtKeyDec = safeStorage.decryptString(Buffer.from(vtKeyEnc, 'base64')); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } } }

    const scanMode = configStore.get('historyScanEnabled') ? 'History' : 'OnAccess';
    logToFile(`Security Engine initialized [Mode: ${scanMode}, Speed: ${configStore.get('scanningSpeed')}%]`);

    currentScanChild.stdin.write(JSON.stringify({ 
        mode: scanMode,
        scanningSpeed: configStore.get('scanningSpeed'),
        processedIds: Array.from(processedIdsCache), 
        releasedFingerprints: Array.from(releasedFingerprintsCache),
        spamKeywords: configStore.get('spamKeywords'), 
        rubrics: configStore.get('rubrics'), 
        whitelist: configStore.get('whitelist'), 
        blacklist: configStore.get('blacklist'), 
        vtKey: vtKeyDec,
        threatIntelligenceLevel: configStore.get('threatIntelligenceLevel'),
        onAccessEnabled: configStore.get('onAccessEnabled') !== false,
        onDemandLimit: configStore.get('onDemandLimit') || 1000,
        deepHistoryScanEnabled: !!configStore.get('deepHistoryScanEnabled')
    }) + '\n');

    const parser = new SafeIPCParser(p => { 
                if (p.type === 'heartbeat') { 
                    lastHeartbeat = Date.now(); 
                    consecutiveEngineFailures = 0;
                    broadcastToUi({ type: 'outlook-status', running: true, standby: false });
                    return; 
                }
                if (p.type === 'store-update') {
                    if (p.key === 'releasedFingerprints' && p.value) {
                        if (!releasedFingerprintsCache.has(p.value)) {
                            serviceSetStore('releasedFingerprints', p.value);
                        }
                    }
                    return;
                }
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'MONITORING', 'INFO', 'ERROR'].includes(p.status)) {
                    // Transparent UTF-8 Base64 decoding
                    if (p.fullHeaders && typeof p.fullHeaders === 'string') {
                        try {
                            const decodedH = Buffer.from(p.fullHeaders, 'base64').toString('utf8');
                            if (decodedH && !decodedH.includes('\ufffd')) {
                                p.fullHeaders = decodedH;
                            }
                        } catch {}
                    }
                    if (p.body && typeof p.body === 'string') {
                        try {
                            const decodedB = Buffer.from(p.body, 'base64').toString('utf8');
                            if (decodedB && !decodedB.includes('\ufffd')) {
                                p.body = decodedB;
                            }
                        } catch {}
                    }

                    if (!p.subject && p.details) p.subject = p.details;
                    if (!p.details && p.subject) p.details = p.subject;
                    if (!p.from && p.sender) p.from = p.sender;
                    if (!p.sender && p.from) p.sender = p.from;
                    if (!p.to && p.recips) p.to = p.recips;
                    if (!p.recipient && p.to) p.recipient = p.to;

                    if (p.status === 'INFO' || p.status === 'MONITORING') {
                        logToFile(`Engine: ${p.details || ''}`);
                    } else if (p.status === 'ERROR') {
                        logToFile(`Engine Error: ${p.details || ''}`, 'ERROR');
                    } else {
                        logToFile(`Scan Result [${p.status}]: ${p.subject || p.details || ''} (${p.sender || 'N/A'})`);
                    }
                    
                    if (p.status !== 'MONITORING' && p.status !== 'INFO' && p.status !== 'ERROR') {
                        if (!p.verdict) return; // Skip status messages without a verdict
                        const v = String(p.verdict).toLowerCase();
                        const cat = v.includes('malicious') ? 'malicious' : (v.includes('suspicious') ? 'suspicious' : (v.includes('spam') ? 'spam' : 'safe'));
                        if (p.userMoved && !p.userMovedStory) {
                            p.userMovedStory = `Originally quarantined by DeskGuard, but you manually moved this email to folder '${p.currentFolder || 'another folder'}' inside Outlook. DeskGuard respects your choice and will keep it here.`;
                        }
                        const fid = p.fingerprint || p.entryId || p.originalEntryId;
                        if (fid) {
                            addProcessedId(fid);
                        }
                        statsBuffer[cat].push(p);
                        scheduleStatePersistenceFlush();
                        try {
                            vaultDB.insertEmail(p);
                        } catch (err) {
                            logToFile(`[VaultDB Error] ${err.message}`, 'WARN');
                        }
                        if (p.fullHeaders || p.body) {
                            const snapshotJson = JSON.stringify({ 
                                fullHeaders: p.fullHeaders || '', 
                                body: p.body || '' 
                            });
                            const idsToPersist = [p.entryId, p.originalEntryId, p.fingerprint].filter(Boolean);
                            for (const idKey of idsToPersist) {
                                const fHash = crypto.createHash('sha256').update(String(idKey)).digest('hex');
                                const fPath = path.join(FORENSICS_DIR, `${fHash}.json`);
                                fsPromises.writeFile(fPath, snapshotJson).catch((err) => {
                                    logToFile(`[Forensics Error] Failed to persist snapshot for ${fHash} (${idKey}): ${err.message}`, 'ERROR');
                                });
                            }
                        }
                    }
                    broadcastToUi({ type: 'scan-update', data: p }); } });
    currentScanChild.stdout.on('data', d => parser.push(d));

    currentScanChild.on('exit', async (code) => { 
        logToFile(`Security Engine process exited with code ${code}`);
        isScanning = false; 
        if (watchdogTimer) {
            clearInterval(watchdogTimer);
            watchdogTimer = null;
        }

        if (code !== 0 && configStore && configStore.get('enabled')) {
            const outlookStillActive = await ensureOutlookRunning();
            if (!outlookStillActive) {
                logToFile('Outlook was closed by user. Security Engine entering standby mode.', 'INFO');
                consecutiveEngineFailures = 0;
                startOutlookStandbyWatcher();
            } else {
                logToFile("Engine exited unexpectedly while Outlook is active. Requesting automatic recovery...", "WARN");
                requestScannerRestart('unexpected-exit');
            }
        }
    });
}

let isRestarting = false;
function requestScannerRestart(reason) {
    if (restartTimer) {
        clearTimeout(restartTimer);
        restartTimer = null;
    }

    if (currentScanChild && ['whitelist', 'blacklist', 'spamKeywords', 'rubrics'].includes(reason)) {
        logToFile(`Live Policy Injection: Updating [${reason}] without restart.`);
        try {
            const vtKeyEnc = configStore.get('vtApiKey');
            let vtKeyDec = '';
            if (vtKeyEnc) { try { vtKeyDec = safeStorage.decryptString(Buffer.from(vtKeyEnc, 'base64')); } catch { } }
            currentScanChild.stdin.write(JSON.stringify({ 
                type: 'config-update',
                whitelist: configStore.get('whitelist'),
                blacklist: configStore.get('blacklist'),
                spamKeywords: configStore.get('spamKeywords'),
                rubrics: configStore.get('rubrics'),
                vtKey: vtKeyDec,
                threatIntelligenceLevel: configStore.get('threatIntelligenceLevel')
            }) + '\n');
            return;
        } catch (e) {
            logToFile(`Live Injection Failed: ${e.message}. Falling back to hard restart.`, 'WARN');
        }
    }

    const isVerifiedFailure = (reason === 'unexpected-exit' || reason === 'watchdog-timeout');
    if (isVerifiedFailure) {
        consecutiveEngineFailures++;
        if (consecutiveEngineFailures >= MAX_ENGINE_FAILURES) {
            logToFile('Engine: Maximum restart attempts reached. Protection paused. Manual restart required.', 'WARN');
            broadcastToUi({ type: 'engine-fault', reason: 'Outlook unresponsive' });
            broadcastToUi({ type: 'outlook-status', running: false, standby: true });
            isScanning = false;
            startOutlookStandbyWatcher();
            return;
        }
    } else {
        consecutiveEngineFailures = 0;
    }

    const delay = isVerifiedFailure ? Math.min(30000, 2000 * Math.pow(2, consecutiveEngineFailures)) : 500;
    logToFile(`Engine restart scheduled in ${delay}ms (attempt ${consecutiveEngineFailures}/${MAX_ENGINE_FAILURES}, reason: ${reason})...`, 'WARN');

    restartTimer = setTimeout(async () => {
        restartTimer = null;
        if (isRestarting) return;
        isRestarting = true;
        try {
            if (!configStore || !configStore.get('enabled')) return;
            const active = await ensureOutlookRunning();
            if (!active) {
                startOutlookStandbyWatcher();
                return;
            }
            logToFile(`Hard Engine Restart triggered [Reason: ${reason}]`);
            if (currentScanChild) {
                killProcessTree(currentScanChild);
                currentScanChild = null;
            }
            isScanning = false;
            await new Promise(r => setTimeout(r, 600));
            await runOutlookScanner();
        } finally {
            isRestarting = false;
        }
    }, delay);
}

function startService() {
    pipeServer = net.createServer(s => {
        let auth = false;
        let buf = '';
        s.on('data', d => {
            buf += d.toString(); let idx = buf.indexOf('\n');
            while (idx > -1) {
                const raw = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
                try {
                    const m = JSON.parse(raw); if (!m) continue;
                    if (!auth) {
                        if (m.type === 'auth' && m.token === PIPE_AUTH_TOKEN) {
                            auth = true;
                            activeConnections.add(s);
                            s.write(JSON.stringify({
                                type: 'status-sync',
                                enabled: !!configStore.get('enabled'),
                                stats: statsCache,
                                config: configStore.store
                            }) + '\n');
                        } else {
                            s.destroy();
                        }
                        continue;
                    }
                    if (m.type === 'store-get') {
                        let val;
                        if (m.key === '') val = configStore ? configStore.store : DEFAULT_CONFIG;
                        else if (m.key === 'processedIds') val = Array.from(processedIdsCache);
                        else if (m.key === 'stats') val = statsCache;
                        else if (m.key === 'releasedFingerprints') val = Array.from(releasedFingerprintsCache);
                        else val = configStore ? configStore.get(m.key) : DEFAULT_CONFIG[m.key];
                        s.write(JSON.stringify({ type: 'store-data', rid: m.rid, key: m.key, value: val }) + '\n');
                    }
                    if (m.type === 'store-set') { 
                        serviceSetStore(m.key, m.value);
                    }
                    if (m.type === 'cmd') { 
                        if (m.payload === 'Exit') {
                            logToFile('Security Service received Exit command. Shutting down gracefully...');
                            try {
                                if (restartTimer) {
                                    clearTimeout(restartTimer);
                                    restartTimer = null;
                                }
                                if (standbyWatcherTimer) {
                                    clearInterval(standbyWatcherTimer);
                                    standbyWatcherTimer = null;
                                }
                                if (watchdogTimer) {
                                    clearInterval(watchdogTimer);
                                    watchdogTimer = null;
                                }
                                if (currentScanChild) {
                                    killProcessTree(currentScanChild);
                                    currentScanChild = null;
                                }
                                if (psWorker) {
                                    killProcessTree(psWorker);
                                    psWorker = null;
                                }
                                flushStatePersistenceSync();
                            } catch {}
                            process.exit(0);
                        }
                        if (m.payload === 'Reset') { 
                            try {
                                if (restartTimer) {
                                    clearTimeout(restartTimer);
                                    restartTimer = null;
                                }
                                if (standbyWatcherTimer) {
                                    clearInterval(standbyWatcherTimer);
                                    standbyWatcherTimer = null;
                                }
                                if (watchdogTimer) {
                                    clearInterval(watchdogTimer);
                                    watchdogTimer = null;
                                }
                                if (currentScanChild) {
                                    killProcessTree(currentScanChild);
                                    currentScanChild = null;
                                }
                                if (psWorker) {
                                    killProcessTree(psWorker);
                                    psWorker = null;
                                }
                            } catch (err) { }
                            try {
                                if (flushTimer) {
                                    clearTimeout(flushTimer);
                                    flushTimer = null;
                                }
                                processedIdsCache.clear();
                                releasedFingerprintsCache.clear();
                                uncommittedProcessedIdsCount = 0;
                                isPersistenceDirty = false;
                                statsBuffer = { malicious: [], suspicious: [], spam: [], safe: [] };
                                statsCache = { ...DEFAULT_DATA.stats };
                                const isSafeUpgrade = m.data && m.data.mode === 'safe';
                                if (!isSafeUpgrade) {
                                    backupStoreBeforeWrite('config');
                                    if (configStore) configStore.clear(); 
                                } 
                                const dataPath = path.join(USER_DATA, 'data.json');
                                const bakPath = path.join(USER_DATA, 'data.json.bak');
                                try {
                                    if (fs.existsSync(dataPath)) fs.unlinkSync(dataPath);
                                    if (fs.existsSync(bakPath)) fs.unlinkSync(bakPath);
                                } catch {}
                                try {
                                    if (vaultDB) vaultDB.close();
                                    const vaultFiles = [
                                        path.join(USER_DATA, 'deskguard_vault.db'),
                                        path.join(USER_DATA, 'deskguard_vault.db-wal'),
                                        path.join(USER_DATA, 'deskguard_vault.db-shm')
                                    ];
                                    vaultFiles.forEach(f => {
                                        if (fs.existsSync(f)) { try { fs.unlinkSync(f); } catch (e) {} }
                                    });
                                } catch {}
                            } catch (err) {
                                logToFile(`[Storage Reset Error]: ${err.message}`, 'ERROR');
                            }
                            try {
                                if (fs.existsSync(LOG_DIR)) fs.rmSync(LOG_DIR, { recursive: true, force: true });
                                if (fs.existsSync(FORENSICS_DIR)) fs.rmSync(FORENSICS_DIR, { recursive: true, force: true });
                            } catch (err) { }
                            process.exit(0); 
                        } 
                        if (m.payload === 'Release' || m.payload === 'Quarantine' || m.payload === 'Delete' || m.payload === 'CleanDuplicates' || m.payload === 'Check-Existence' || m.payload === 'DuplicateScan' || m.payload === 'CloudVirusScan' || m.payload === 'OpenInOutlook') {
                            const worker = getPsWorker();
                            let cmdData = { ...m.data };
                            if (m.payload === 'CloudVirusScan') {
                                let vtKeyDec = '';
                                const vtKeyEnc = configStore && configStore.get('vtApiKey');
                                if (vtKeyEnc) { try { vtKeyDec = safeStorage.decryptString(Buffer.from(vtKeyEnc, 'base64')); } catch {} }
                                cmdData.vtKey = vtKeyDec;
                            }
                            if (worker && worker.stdin) worker.stdin.write(JSON.stringify({ action: m.payload, rid: m.rid, ...cmdData }) + '\n'); 
                        }
                        if (m.payload === 'ResetDuplicateStack') {
                            const worker = getPsWorker();
                            if (worker && worker.stdin) worker.stdin.write(JSON.stringify({ action: 'ResetDuplicateStack' }) + '\n');
                        }
                        if (m.payload === 'StopAllScans') {
                            logToFile('Security Service: StopAllScans received. Halting all scanner and worker processes.');
                            if (standbyWatcherTimer) { clearInterval(standbyWatcherTimer); standbyWatcherTimer = null; }
                            if (watchdogTimer) { clearInterval(watchdogTimer); watchdogTimer = null; }
                            if (restartTimer) { clearTimeout(restartTimer); restartTimer = null; }
                            if (currentScanChild) { killProcessTree(currentScanChild); currentScanChild = null; }
                            if (psWorker) { killProcessTree(psWorker); psWorker = null; }
                            isScanning = false;
                            broadcastToUi({ type: 'outlook-status', running: false, standby: false });
                        }
                    }
                } catch (err) { }
            }
        });
        s.on('close', () => activeConnections.delete(s));
    });

    pipeServer.on('error', (err) => {
        if (err.code === 'EADDRINUSE') {
            logToFile(`[Service IPC] Pipe ${PIPE_NAME} is already in use by active service instance. Exiting duplicate.`, 'INFO');
            process.exit(0);
        } else {
            logToFile(`[Service IPC Error]: ${err.message}`, 'ERROR');
        }
    });

    pipeServer.listen(PIPE_NAME, () => {
        logToFile(`Security Service IPC Layer: ACTIVE on ${PIPE_NAME}`);
        if (configStore && configStore.get('enabled')) {
            isOutlookProcessRunning().then(running => {
                if (running) {
                    runOutlookScanner();
                } else {
                    startOutlookStandbyWatcher();
                }
            });
        }
    });
}

const reqHandlers = new Map();
let pipeBuffer = '';
function setupPipeClient() {
    uiPipeClient.on('data', d => {
        pipeBuffer += d.toString(); let idx = pipeBuffer.indexOf('\n');
        while (idx > -1) {
            const raw = pipeBuffer.slice(0, idx).trim(); pipeBuffer = pipeBuffer.slice(idx + 1); idx = pipeBuffer.indexOf('\n');
            try {
                const r = JSON.parse(raw); if (!r) continue;
                if (r.type === 'store-data') {
                    if (r.key === '') {
                        configCache = { ...DEFAULT_CONFIG, ...r.value };
                    }
                    if (r.key === 'stats') statsCache = r.value;
                    if (r.rid && reqHandlers.has(r.rid)) { const resolve = reqHandlers.get(r.rid); reqHandlers.delete(r.rid); resolve(r.value); }
                } else if (r.type === 'cmd-response') {
                    if (r.rid && reqHandlers.has(r.rid)) {
                        const resolve = reqHandlers.get(r.rid);
                        reqHandlers.delete(r.rid);
                        resolve(r.data !== undefined ? r.data : r);
                    }
                    broadcastToUi(r);
                } else if (r.type === 'stats-update') {
                    if (r.data && r.data.full) statsCache = r.data.stats;
                    broadcastToUi(r);
                } else if (r.type === 'status-sync') {
                    if (r.stats) statsCache = r.stats;
                    if (r.config) {
                        configCache = { ...DEFAULT_CONFIG, ...r.config };
                        if (r.config.launchAtStartup !== undefined) {
                            applyWindowsStartupSetting(!!r.config.launchAtStartup);
                        }
                    }
                    isEnabled = r.enabled;
                    updateTrayState();
                    broadcastToUi({ type: 'stats-update', data: { full: true, stats: statsCache } });
                    broadcastToUi(r);
                } else if (r.type === 'storage-error') {
                    logToFile(`[Service Storage Error] ${r.store}: ${r.error}`, 'ERROR');
                    broadcastToUi({ type: 'live-log', message: `[Storage Error] ${r.store}: ${r.error}` });
                } else broadcastToUi(r);
            } catch (err) { }
        }
    });

    uiPipeClient.on('close', () => {
        uiPipeClient = null;
        if (!isQuitting) {
            logToFile('Security Service pipe connection closed. Attempting reconnect in 2s...', 'WARN');
            setTimeout(spawnService, 2000);
        }
    });
}

function spawnService() {
    if (isServiceMode || serviceSpawnInFlight) return;
    serviceSpawnInFlight = true;

    const probe = net.connect(PIPE_NAME, () => {
        logToFile('Security Service is already running. Connected directly to IPC pipe.');
        uiPipeClient = probe;
        serviceSpawnInFlight = false;
        uiPipeClient.write(JSON.stringify({ type: 'auth', token: PIPE_AUTH_TOKEN }) + '\n');
        setupPipeClient();
    });

    probe.on('error', () => {
        probe.destroy();
        logToFile('No active Security Service found. Spawning background service...');
        const env = { ...process.env };
        delete env.ELECTRON_RUN_AS_NODE;
        const child = spawn(process.execPath, [APP_ROOT, '--service'], {
            detached: true,
            windowsHide: true,
            env,
            stdio: 'ignore'
        });
        child.unref();

        let attempts = 0;
        const maxAttempts = 60;
        const tryConnect = () => {
            const client = net.connect(PIPE_NAME, () => {
                uiPipeClient = client;
                serviceSpawnInFlight = false;
                logToFile('Successfully connected to Security Service IPC pipe.');
                uiPipeClient.write(JSON.stringify({ type: 'auth', token: PIPE_AUTH_TOKEN }) + '\n');
                setupPipeClient();
            });
            client.on('error', () => {
                client.destroy();
                if (++attempts < maxAttempts) {
                    setTimeout(tryConnect, 200);
                } else {
                    serviceSpawnInFlight = false;
                    logToFile('Failed to connect to Security Service after maximum attempts.', 'ERROR');
                }
            });
        };
        setTimeout(tryConnect, 300);
    });
}

app.on('ready', () => {
    Menu.setApplicationMenu(null);
    ensureOutlookProgrammaticAccessPolicies();

    if (isServiceMode) { 
        if (configStore) {
            const shouldStart = configStore.get('launchAtStartup');
            if (shouldStart !== undefined) {
                applyWindowsStartupSetting(!!shouldStart);
            }
            const currentKeywords = configStore.get('spamKeywords') || [];
            const defaultKeywords = ['viagra', 'lottery', 'urgent', 'bitcoin', 'winner', 'unpaid', 'invoice', 'payment', 'account', 'verify', 'security', 'update', 'action', 'urgent-action', 'account-compromise', 'limited-access', 'security-alert', 'suspicious-activity'];
            let keywordsChanged = false;
            defaultKeywords.forEach(kw => { if (!currentKeywords.includes(kw)) { currentKeywords.push(kw); keywordsChanged = true; } });
            if (keywordsChanged) serviceSetStore('spamKeywords', currentKeywords);
        }

        startService();
        runDiagnostics();
    } else {
        const gotLock = app.requestSingleInstanceLock();
        if (!gotLock) {
            app.quit();
            return;
        }

        app.on('second-instance', () => {
            if (mainWindow) {
                if (mainWindow.isMinimized()) mainWindow.restore();
                if (!mainWindow.isVisible()) mainWindow.show();
                mainWindow.focus();
            }
        });

        const icon = nativeImage.createFromPath(path.join(APP_ROOT, 'tray_off.png')).resize({ width: 16, height: 16 });
        tray = new Tray(icon);
        tray.setToolTip('DeskGuard for Microsoft Outlook');
        tray.on('click', () => {
            if (mainWindow) {
                if (mainWindow.isVisible()) mainWindow.hide();
                else {
                    mainWindow.show();
                    mainWindow.focus();
                }
            }
        });
        updateTrayState();

        mainWindow = new BrowserWindow({
            width: 1500,
            height: 900,
            backgroundColor: '#0a0e1c',
            show: false,
            frame: false,
            webPreferences: {
                preload: path.join(APP_ROOT, 'preload.js'),
                contextIsolation: true,
                sandbox: true
            }
        });
        mainWindow.loadFile('index.html');
        mainWindow.on('close', e => {
            if (!isQuitting) {
                e.preventDefault();
                mainWindow.hide();
            }
        });
        mainWindow.once('ready-to-show', () => {
            if (!isHiddenMode) {
                mainWindow.show();
            } else {
                logToFile('DeskGuard launched on boot in background mode (--hidden). Running minimized to system tray.');
            }
        });
        spawnService();
    }
});

app.on('before-quit', () => {
    isQuitting = true;
    if (isServiceMode) {
        if (currentScanChild) {
            killProcessTree(currentScanChild);
            currentScanChild = null;
        }
        if (psWorker) {
            killProcessTree(psWorker);
            psWorker = null;
        }
        flushStatePersistenceSync();
    }
    if (uiPipeClient) {
        try {
            uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Exit' }) + '\n');
        } catch {}
    }
});

process.on('SIGINT', () => {
    if (isServiceMode) {
        if (currentScanChild) killProcessTree(currentScanChild);
        if (psWorker) killProcessTree(psWorker);
        flushStatePersistenceSync();
    }
    process.exit(0);
});
process.on('SIGTERM', () => {
    if (isServiceMode) {
        if (currentScanChild) killProcessTree(currentScanChild);
        if (psWorker) killProcessTree(psWorker);
        flushStatePersistenceSync();
    }
    process.exit(0);
});

async function setProtectionState(targetEnabled) {
    const state = !!targetEnabled;
    isEnabled = state;
    configCache.enabled = state;
    configCache.onAccessEnabled = state;
    configCache.deepHistoryScanEnabled = state;
    updateTrayState();

    if (configStore) {
        try {
            configStore.set('enabled', state);
            configStore.set('onAccessEnabled', state);
            configStore.set('deepHistoryScanEnabled', state);
        } catch {}
    } else {
        try {
            const configPath = path.join(USER_DATA, 'config.json');
            if (fs.existsSync(configPath)) {
                const cur = JSON.parse(fs.readFileSync(configPath, 'utf8'));
                cur.enabled = state;
                cur.onAccessEnabled = state;
                cur.deepHistoryScanEnabled = state;
                fs.writeFileSync(configPath, JSON.stringify(cur, null, 2), 'utf8');
            }
        } catch {}
    }

    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('status-sync', {
            enabled: state,
            stats: statsCache,
            config: configCache
        });
        mainWindow.webContents.send('outlook-status', {
            running: state,
            standby: !state
        });
    }

    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: state }) + '\n');
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'onAccessEnabled', value: state }) + '\n');
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'deepHistoryScanEnabled', value: state }) + '\n');
        if (!state) {
            uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'StopAllScans' }) + '\n');
        }
    }

    if (isServiceMode) {
        if (!state) {
            logToFile('Security Service: Protection stopped. Halting all scanner and worker processes...');
            if (standbyWatcherTimer) { clearInterval(standbyWatcherTimer); standbyWatcherTimer = null; }
            if (watchdogTimer) { clearInterval(watchdogTimer); watchdogTimer = null; }
            if (restartTimer) { clearTimeout(restartTimer); restartTimer = null; }
            if (currentScanChild) { killProcessTree(currentScanChild); currentScanChild = null; }
            if (psWorker) { killProcessTree(psWorker); psWorker = null; }
            isScanning = false;
            broadcastToUi({ type: 'outlook-status', running: false, standby: false });
        } else {
            logToFile('Security Service: Protection activated. Launching engine...');
            consecutiveEngineFailures = 0;
            isStandbyMode = false;
            ensureOutlookRunning().then(running => {
                if (running) {
                    runOutlookScanner();
                } else {
                    startOutlookStandbyWatcher();
                }
            });
        }
    }
}

function updateTrayState() {
    if (isServiceMode || !tray) return;
    const iconName = isEnabled ? 'tray_on.png' : 'tray_off.png';
    const windowIconName = isEnabled ? 'icon_on.png' : 'icon_off.png';
    const icon = nativeImage.createFromPath(path.join(APP_ROOT, iconName)).resize({ width: 16, height: 16 });
    tray.setImage(icon);
    tray.setToolTip(`DeskGuard for Microsoft Outlook - ${isEnabled ? 'Active' : 'Disabled'}`);
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.setIcon(nativeImage.createFromPath(path.join(APP_ROOT, windowIconName)));
    tray.setContextMenu(Menu.buildFromTemplate([
        { label: 'Show Dashboard', click: () => { if (mainWindow) { mainWindow.show(); mainWindow.focus(); } } },
        { label: isEnabled ? 'Security: ACTIVE' : 'Security: DISABLED', enabled: false },
        { label: isEnabled ? 'Stop Protection' : 'Start Protection', click: () => {
            setProtectionState(!isEnabled);
        } },
        { type: 'separator' },
        { label: 'Exit Application', click: () => {
            isQuitting = true;
            if (uiPipeClient) {
                try {
                    uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Exit' }) + '\n');
                } catch {}
            }
            setTimeout(() => app.quit(), 150);
        } }
    ]));
}

const pipeReq = (m) => new Promise(resolve => { 
    if (!uiPipeClient) return resolve(null); 
    const rid = crypto.randomBytes(8).toString('hex'); 
    const timeout = setTimeout(() => { if (reqHandlers.has(rid)) { reqHandlers.delete(rid); resolve(null); } }, 5000);
    reqHandlers.set(rid, (val) => { clearTimeout(timeout); resolve(val); }); 
    uiPipeClient.write(JSON.stringify({ ...m, rid }) + '\n'); 
});

ipcMain.on('window-minimize', () => { if (mainWindow) mainWindow.minimize(); });
ipcMain.on('window-maximize', () => { 
    if (mainWindow) {
        if (mainWindow.isMaximized()) {
            mainWindow.unmaximize();
        } else {
            mainWindow.maximize();
        }
    } 
});
ipcMain.on('window-hide', () => { if (mainWindow) mainWindow.hide(); });
ipcMain.on('window-close', () => { if (mainWindow) mainWindow.hide(); });

ipcMain.handle('get-config', () => {
    let cfg = { ...DEFAULT_CONFIG, ...configCache };
    if (cfg.vtApiKey) {
        cfg.vtApiKey = getDecryptedVtKey(cfg.vtApiKey);
    }
    return cfg;
});
ipcMain.handle('pause-duplicate-scan', () => { try { fs.writeFileSync(path.join(APP_ROOT, '.dup_pause'), '1'); return { ok: true }; } catch (e) { return { ok: false, error: e.message }; } });
ipcMain.handle('resume-duplicate-scan', () => { const p = path.join(APP_ROOT, '.dup_pause'); try { if (fs.existsSync(p)) fs.unlinkSync(p); return { ok: true }; } catch (e) { return { ok: false, error: e.message }; } });
ipcMain.handle('get-stats', () => {
    return statsCache || { ...DEFAULT_DATA.stats };
});
function decodeIfLegacyBase64(str) {
    if (!str || typeof str !== 'string' || str === 'N/A' || str === 'Unavailable') {
        return str || 'N/A';
    }

    if (/[: <>@;]/.test(str)) {
        return str;
    }

    const clean = str.trim().replace(/[\r\n\s]+/g, '');
    if (clean.length === 0 || clean.length % 4 !== 0) {
        return str;
    }

    if (!/^[A-Za-z0-9+/]+={0,2}$/.test(clean)) {
        return str;
    }

    try {
        const decoded = Buffer.from(clean, 'base64').toString('utf8');
        const reEncoded = Buffer.from(decoded, 'utf8').toString('base64');
        if (reEncoded === clean && !decoded.includes('\uFFFD')) {
            return decoded;
        }
    } catch {
        return str;
    }

    return str;
}

ipcMain.handle('get-forensics', async (e, id) => { 
    if (!id) {
        return { fullHeaders: 'Unavailable', body: 'Unavailable' };
    }
    const candidates = typeof id === 'object'
        ? [id.entryId, id.originalEntryId, id.fingerprint].filter(Boolean)
        : [id];
    
    for (const cand of candidates) {
        const fHash = crypto.createHash('sha256').update(String(cand)).digest('hex'); 
        const fPath = path.join(FORENSICS_DIR, `${fHash}.json`); 
        try {
            const rawContent = await fsPromises.readFile(fPath, 'utf8');
            const data = JSON.parse(rawContent);
            return { 
                fullHeaders: decodeIfLegacyBase64(data.fullHeaders) || 'N/A', 
                body: decodeIfLegacyBase64(data.body) || 'N/A' 
            };
        } catch (err) {
            // Check next candidate
        }
    }
    if (vaultDB) {
        for (const cand of candidates) {
            try {
                const item = vaultDB.getEmailById(cand);
                if (item && (item.fullHeaders || item.headers || item.body)) {
                    const h = item.fullHeaders || item.headers || '';
                    const b = item.body || '';
                    return {
                        fullHeaders: decodeIfLegacyBase64(h) || (h ? h : 'N/A'),
                        body: decodeIfLegacyBase64(b) || (b ? b : 'N/A')
                    };
                }
            } catch {}
        }
    }
    logToFile(`[Forensics] Snapshot not found for query: ${typeof id === 'object' ? JSON.stringify(id) : id}`, 'INFO');
    return { fullHeaders: 'Unavailable', body: 'Unavailable' };
});
ipcMain.handle('set-processed-ids', (e, v) => { 
    if (uiPipeClient) { 
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'processedIds', value: v }) + '\n'); 
        return { ok: true }; 
    } 
    return { ok: false, error: 'Service initializing' }; 
});

ipcMain.handle('set-enabled', (e, v) => { 
    setProtectionState(!!v);
    return { ok: true }; 
});
ipcMain.handle('set-history-enabled', (e, v) => {
    configCache.historyScanEnabled = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'historyScanEnabled', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-vt-key', (e, v) => {
    let encKey = '';
    if (v) {
        try {
            encKey = safeStorage.encryptString(v).toString('base64');
        } catch {
            encKey = v;
        }
    }
    configCache.vtApiKey = encKey;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'vtApiKey', value: encKey }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-spam-keywords', (e, v) => {
    configCache.spamKeywords = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'spamKeywords', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-rubrics', (e, v) => {
    configCache.rubrics = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'rubrics', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-whitelist', (e, v) => {
    configCache.whitelist = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'whitelist', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-blacklist', (e, v) => {
    configCache.blacklist = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'blacklist', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('save-column-widths', (e, v) => {
    configCache.columnWidths = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'columnWidths', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-scanning-speed', (e, v) => {
    configCache.scanningSpeed = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'scanningSpeed', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-threat-intel-level', (e, v) => {
    configCache.threatIntelligenceLevel = v;
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'threatIntelligenceLevel', value: v }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('set-first-run', (e, v) => {
    const isFirst = !!v;
    configCache.firstRun = isFirst;
    if (configStore) {
        try { configStore.set('firstRun', isFirst); } catch {}
    } else {
        try {
            const configPath = path.join(USER_DATA, 'config.json');
            const cur = fs.existsSync(configPath) ? JSON.parse(fs.readFileSync(configPath, 'utf8')) : { ...DEFAULT_CONFIG };
            cur.firstRun = isFirst;
            fs.writeFileSync(configPath, JSON.stringify(cur, null, 2), 'utf8');
        } catch {}
    }
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'firstRun', value: isFirst }) + '\n');
    }
    return { ok: true };
});
ipcMain.handle('set-startup', (e, v) => {
    const enabled = !!v;
    configCache.launchAtStartup = enabled;
    applyWindowsStartupSetting(enabled);
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'launchAtStartup', value: enabled }) + '\n');
        return { ok: true };
    }
    return { ok: true };
});
ipcMain.handle('check-startup', async () => {
    const configEnabled = !!(configCache && configCache.launchAtStartup);
    return new Promise(resolve => {
        execFile('reg', ['query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run', '/v', 'DeskGuardForMicrosoftOutlook'], (err, stdout) => {
            const regExists = !err && !!(stdout && stdout.includes('DeskGuardForMicrosoftOutlook'));
            let loginItem = false;
            try {
                loginItem = app.getLoginItemSettings().openAtLogin;
            } catch {}
            resolve({
                configEnabled,
                systemActive: regExists || loginItem,
                regExists,
                loginItem
            });
        });
    });
});
ipcMain.handle('release-email', async (e, d) => {
    if (!uiPipeClient || !d || (!d.entryId && !d.fingerprint)) return { success: false, error: 'Invalid parameters' };
    const rid = crypto.randomBytes(8).toString('hex');
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ success: false, error: 'Release operation timed out' });
        }, 5000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            if (val && val.success) {
                const currentStats = statsCache || { ...DEFAULT_DATA.stats };
                const entryId = d.entryId;
                const fp = (val.data && val.data.fingerprint) || val.fingerprint || d.fingerprint;
                let foundItem = null;
                for (const cat of ['malicious', 'spam', 'suspicious']) {
                    if (currentStats[cat]) {
                        const idx = currentStats[cat].findIndex(i => (entryId && i.entryId === entryId) || (fp && i.fingerprint === fp));
                        if (idx !== -1) {
                            foundItem = currentStats[cat].splice(idx, 1)[0];
                            break;
                        }
                    }
                }
                if (foundItem) {
                    foundItem.verdict = 'Safe';
                    foundItem.tier = 'User Released';
                    foundItem.score = 100;
                    const newEntryId = (val.data && val.data.newEntryId) || val.newEntryId;
                    if (newEntryId) foundItem.entryId = newEntryId;
                    if (!currentStats.safe) currentStats.safe = [];
                    currentStats.safe.unshift(foundItem);
                }
                statsCache = currentStats;
                if (uiPipeClient) {
                    uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'stats', value: currentStats }) + '\n');
                    if (fp) {
                        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'releasedFingerprints', value: fp }) + '\n');
                    }
                }
                broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
                logToFile(`Release: Successfully released email "${foundItem ? foundItem.subject : (d.subject || entryId)}" to Inbox.`);
            }
            resolve(val !== undefined ? val : { success: false });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Release', rid, data: d }) + '\n');
    });
});

ipcMain.handle('quarantine-email', async (e, d) => {
    if (!uiPipeClient || !d || (!d.entryId && !d.fingerprint)) return { success: false, error: 'Invalid parameters' };
    const rid = crypto.randomBytes(8).toString('hex');
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ success: false, error: 'Quarantine operation timed out' });
        }, 5000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            if (val && val.success) {
                const currentStats = statsCache || { ...DEFAULT_DATA.stats };
                const entryId = d.entryId;
                const fp = d.fingerprint;
                let foundItem = null;
                for (const cat of ['safe', 'suspicious']) {
                    if (currentStats[cat]) {
                        const idx = currentStats[cat].findIndex(i => (entryId && i.entryId === entryId) || (fp && i.fingerprint === fp));
                        if (idx !== -1) {
                            foundItem = currentStats[cat].splice(idx, 1)[0];
                            break;
                        }
                    }
                }
                if (foundItem) {
                    foundItem.verdict = 'Quarantined';
                    foundItem.tier = 'User Quarantined';
                    foundItem.score = 0;
                    const newEntryId = (val.data && val.data.newEntryId) || val.newEntryId;
                    if (newEntryId) foundItem.entryId = newEntryId;
                    const targetCat = (d.targetFolder === 3) ? 'malicious' : 'spam';
                    if (!currentStats[targetCat]) currentStats[targetCat] = [];
                    currentStats[targetCat].unshift(foundItem);
                }
                statsCache = currentStats;
                if (uiPipeClient) {
                    uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'stats', value: currentStats }) + '\n');
                }
                broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
                logToFile(`Quarantine: Successfully quarantined email "${foundItem ? foundItem.subject : (d.subject || entryId)}".`);
            }
            resolve(val !== undefined ? val : { success: false });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Quarantine', rid, data: d }) + '\n');
    });
});

ipcMain.handle('delete-email', async (e, d) => {
    if (!uiPipeClient || !d || (!d.entryId && (!d.entryIds || d.entryIds.length === 0))) return { ok: false, error: 'Invalid parameters' };
    const rid = crypto.randomBytes(8).toString('hex');
    const idsToDelete = new Set(d.entryIds ? d.entryIds : [d.entryId]);
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ ok: false, error: 'Delete operation timed out' });
        }, 5000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            const currentStats = statsCache || { ...DEFAULT_DATA.stats };
            for (const cat of ['malicious', 'suspicious', 'spam', 'safe']) {
                if (currentStats[cat]) {
                    currentStats[cat] = currentStats[cat].filter(i => !idsToDelete.has(i.entryId));
                }
            }
            statsCache = currentStats;
            if (uiPipeClient) {
                uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'stats', value: currentStats }) + '\n');
            }
            broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
            const deletedCount = (val && val.data && val.data.count !== undefined) ? val.data.count : (val && val.count !== undefined ? val.count : idsToDelete.size);
            logToFile(`Delete: Removed ${deletedCount} email(s) from Outlook.`);
            resolve({ ok: true, count: deletedCount });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Delete', rid, data: d }) + '\n');
    });
});

ipcMain.handle('verify-existence', async (e, d) => {
    if (!uiPipeClient || !d.items || d.items.length === 0) return { removedCount: 0, userMovedCount: 0 };
    const rid = crypto.randomBytes(8).toString('hex');
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ removedCount: 0, userMovedCount: 0 });
        }, 5000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            const removed = (val && (val.removed || (val.data && val.data.removed))) || [];
            const userMoved = (val && (val.userMoved || (val.data && val.data.userMoved))) || [];
            let statsChanged = false;
            const currentStats = statsCache || { ...DEFAULT_DATA.stats };

            if (removed.length > 0) {
                const removedIds = new Set(removed.map(r => r.entryId));
                for (const cat of ['malicious', 'suspicious', 'spam', 'safe']) {
                    if (currentStats[cat]) {
                        currentStats[cat] = currentStats[cat].filter(i => !removedIds.has(i.entryId));
                    }
                }
                statsChanged = true;
                logToFile(`Verification: Pruned ${removed.length} missing/deleted email(s) from incident records.`);
            }

            if (userMoved.length > 0) {
                const movedMap = new Map();
                for (const um of userMoved) {
                    if (um && um.entryId) {
                        movedMap.set(um.entryId, um);
                        try {
                            vaultDB.updateUserMoved(um.entryId, um.currentFolder || 'Unknown');
                        } catch (dbErr) {
                            logToFile(`Verification: vaultDB.updateUserMoved error for ${um.entryId}: ${dbErr.message}`);
                        }
                    }
                }
                for (const cat of ['malicious', 'suspicious', 'spam', 'safe']) {
                    if (currentStats[cat]) {
                        for (const item of currentStats[cat]) {
                            if (movedMap.has(item.entryId)) {
                                const info = movedMap.get(item.entryId);
                                item.userMoved = true;
                                item.currentFolder = info.currentFolder || 'Unknown';
                                item.userMovedStory = `Relocated by user to "${item.currentFolder}". DeskGuard will not move or re-quarantine this email again.`;
                                statsChanged = true;
                            }
                        }
                    }
                }
                logToFile(`Verification: Detected ${userMoved.length} email(s) relocated by user in Outlook.`);
            }

            if (statsChanged) {
                statsCache = currentStats;
                if (uiPipeClient) {
                    uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'stats', value: currentStats }) + '\n');
                }
                broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
            }
            resolve({ removedCount: removed.length, userMovedCount: userMoved.length });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Check-Existence', rid, data: d }) + '\n');
    });
});

ipcMain.handle('open-logs-folder', () => shell.openPath(LOG_DIR));
ipcMain.handle('app-reset', (e, mode) => { 
    const isSafeUpgrade = (mode === 'safe');
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Reset', data: { mode: isSafeUpgrade ? 'safe' : 'scorched' } }) + '\n'); 
    }
    if (!isSafeUpgrade) {
        configCache = { ...DEFAULT_CONFIG, enabled: false };
        try { applyWindowsStartupSetting(false); } catch (err) {}
    }
    statsCache = { ...DEFAULT_DATA.stats };
    if (vaultDB) {
        try { vaultDB.close(); } catch (err) {}
    }
    setTimeout(() => {
        try {
            const vaultFiles = [
                path.join(USER_DATA, 'deskguard_vault.db'),
                path.join(USER_DATA, 'deskguard_vault.db-wal'),
                path.join(USER_DATA, 'deskguard_vault.db-shm')
            ];
            vaultFiles.forEach(f => { if (fs.existsSync(f)) { try { fs.unlinkSync(f); } catch (e) {} } });
            const dataPath = path.join(USER_DATA, 'data.json');
            const bakPath = path.join(USER_DATA, 'data.json.bak');
            if (fs.existsSync(dataPath)) try { fs.unlinkSync(dataPath); } catch (e) {}
            if (fs.existsSync(bakPath)) try { fs.unlinkSync(bakPath); } catch (e) {}
            if (!isSafeUpgrade && configStore) {
                try { configStore.clear(); } catch (e) {}
            }
            if (fs.existsSync(LOG_DIR)) fs.rmSync(LOG_DIR, { recursive: true, force: true });
            if (fs.existsSync(FORENSICS_DIR)) fs.rmSync(FORENSICS_DIR, { recursive: true, force: true });
        } catch (err) {}
        app.relaunch(); 
        app.exit(); 
    }, 1500); 
});
ipcMain.handle('scan-duplicates', async () => { if (uiPipeClient) { if (!psWorker || psWorker.killed) getPsWorker(); uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'DuplicateScan' }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('delete-duplicates', async (e, d) => { 
    if (!uiPipeClient) return { ok: false, error: 'Service initializing' }; 
    const rid = crypto.randomBytes(8).toString('hex');
    return new Promise(resolve => {
        const timeout = setTimeout(() => { reqHandlers.delete(rid); resolve({ ok: false, error: 'Duplicate cleanup timed out' }); }, 60000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            const reportData = (val && val.data) ? val.data : val;
            const successCount = (reportData && reportData.successCount !== undefined) ? reportData.successCount : ((reportData && reportData.count !== undefined) ? reportData.count : 0);
            const skippedCount = (reportData && reportData.skippedCount !== undefined) ? reportData.skippedCount : 0;
            resolve({ ok: true, count: successCount, successCount, skippedCount, errors: reportData && reportData.errors, skipped: reportData && reportData.skipped, data: reportData });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'CleanDuplicates', rid, data: d }) + '\n');
    });
});
ipcMain.handle('reset-duplicate-engine', async () => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'ResetDuplicateStack' }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('scan-virus', async (e, id) => {
    let vtKeyDec = '';
    const vtKeyEnc = configCache && configCache.vtApiKey;
    if (vtKeyEnc) {
        try { vtKeyDec = getDecryptedVtKey(vtKeyEnc); } catch {}
    }
    if (!vtKeyDec || vtKeyDec === 'MASKED_FOR_SECURITY' || vtKeyDec.trim().length < 16) {
        return { 
            success: false, 
            code: 'NO_API_KEY', 
            error: 'VirusTotal API key is not configured. Please enter your API key to enable live threat intelligence scanning.' 
        };
    }
    if (!uiPipeClient) return { success: false, error: 'Service disconnected' };
    const rid = crypto.randomBytes(8).toString('hex');
    const threatLevel = (configCache && configCache.threatIntelligenceLevel !== undefined)
        ? configCache.threatIntelligenceLevel
        : DEFAULT_CONFIG.threatIntelligenceLevel;
    return new Promise(resolve => {
        const timeout = setTimeout(() => { 
            reqHandlers.delete(rid); 
            resolve({ success: false, code: 'TIMEOUT', error: 'Cloud Scan Timeout. VirusTotal API or Outlook did not respond within 25 seconds.' }); 
        }, 25000);
        reqHandlers.set(rid, (val) => { 
            clearTimeout(timeout); 
            resolve({ success: val.success !== false, data: val.data || val }); 
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'CloudVirusScan', rid, data: { entryId: id, threatIntelligenceLevel: threatLevel, vtKey: vtKeyDec } }) + '\n');
    });
});
ipcMain.handle('export-config', async () => { 
    const { filePath } = await dialog.showSaveDialog({ 
        title: 'Export DeskGuard Configuration', 
        defaultPath: path.join(app.getPath('downloads'), 'deskguard-outlook-config.json'), 
        filters: [{ name: 'JSON Files', extensions: ['json'] }] 
    }); 
    if (!filePath) return { canceled: true }; 
    const cfg = { ...DEFAULT_CONFIG, ...configCache }; 
    // MASK SENSITIVE DATA
    const exportData = { 
        vtApiKey: "MASKED_FOR_SECURITY", 
        spamKeywords: cfg.spamKeywords, 
        rubrics: cfg.rubrics, 
        whitelist: cfg.whitelist, 
        blacklist: cfg.blacklist, 
        launchAtStartup: cfg.launchAtStartup 
    }; 
    fs.writeFileSync(filePath, JSON.stringify(exportData, null, 4)); 
    return { success: true, filePath }; 
});
ipcMain.handle('import-config', async () => {
    const { filePaths } = await dialog.showOpenDialog({
        title: 'Import DeskGuard Configuration',
        filters: [{ name: 'JSON Files', extensions: ['json'] }],
        properties: ['openFile']
    });
    if (!filePaths || filePaths.length === 0) return { canceled: true };
    try {
        const content = fs.readFileSync(filePaths[0], 'utf8');
        const data = JSON.parse(content);
        const keys = ['vtApiKey', 'spamKeywords', 'rubrics', 'whitelist', 'blacklist', 'launchAtStartup', 'scanningSpeed', 'threatIntelligenceLevel', 'historyScanEnabled', 'firstRun'];
        for (const k of keys) {
            if (data[k] !== undefined) {
                let val = data[k];
                if (k === 'vtApiKey' && val && val !== 'MASKED_FOR_SECURITY') {
                    try {
                        val = safeStorage.encryptString(val).toString('base64');
                    } catch {}
                } else if (k === 'vtApiKey' && val === 'MASKED_FOR_SECURITY') {
                    continue;
                }
                configCache[k] = val;
                if (k === 'launchAtStartup') {
                    applyWindowsStartupSetting(!!val);
                }
                if (uiPipeClient) {
                    uiPipeClient.write(JSON.stringify({ type: 'store-set', key: k, value: val }) + '\n');
                }
            }
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
});

ipcMain.handle('search-vault', (e, params = {}) => {
    try {
        const results = vaultDB.searchEmails(params);
        return { ok: true, results };
    } catch (err) {
        logToFile(`[Vault Search Error]: ${err.message}`, 'WARN');
        return { ok: false, error: err.message, results: { total: 0, rows: [] } };
    }
});

ipcMain.handle('get-sender-suggestions', (e, prefix) => {
    try {
        const suggestions = vaultDB.getSenderSuggestions(prefix, 10);
        return { ok: true, suggestions };
    } catch (err) {
        return { ok: false, error: err.message, suggestions: [] };
    }
});

ipcMain.handle('open-email', async (e, d) => {
    if (!d || !d.entryId) return { success: false, error: 'Invalid parameters: entryId is required' };
    const rid = crypto.randomBytes(8).toString('hex');
    if (!uiPipeClient) {
        try {
            const worker = getPsWorker();
            if (!worker || !worker.stdin) return { success: false, error: 'Outlook Worker process unavailable' };
            return new Promise(resolve => {
                const timeout = setTimeout(() => {
                    reqHandlers.delete(rid);
                    resolve({ success: false, error: 'Open in Outlook operation timed out' });
                }, 10000);
                reqHandlers.set(rid, (val) => {
                    clearTimeout(timeout);
                    resolve(val !== undefined ? val : { success: false });
                });
                worker.stdin.write(JSON.stringify({ action: 'OpenInOutlook', rid, entryId: d.entryId, storeId: d.storeId }) + '\n');
            });
        } catch (err) {
            return { success: false, error: err.message };
        }
    }
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ success: false, error: 'Open in Outlook operation timed out' });
        }, 10000);
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            resolve(val !== undefined ? val : { success: false });
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'OpenInOutlook', rid, data: d }) + '\n');
    });
});

ipcMain.handle('set-scan-modes', (e, v) => {
    if (!v) return { ok: false };
    if (v.onAccessEnabled !== undefined) {
        configCache.onAccessEnabled = !!v.onAccessEnabled;
        if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'onAccessEnabled', value: !!v.onAccessEnabled }) + '\n');
    }
    if (v.onDemandLimit !== undefined) {
        configCache.onDemandLimit = Number(v.onDemandLimit) || 1000;
        if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'onDemandLimit', value: Number(v.onDemandLimit) || 1000 }) + '\n');
    }
    if (v.deepHistoryScanEnabled !== undefined) {
        configCache.deepHistoryScanEnabled = !!v.deepHistoryScanEnabled;
        if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'deepHistoryScanEnabled', value: !!v.deepHistoryScanEnabled }) + '\n');
    }
    return { ok: true };
});
