const path = require('node:path');
const fs = require('node:fs');
const fsPromises = fs.promises;
const net = require('node:net');
const crypto = require('node:crypto');
const { spawn, execFile, execSync } = require('node:child_process');
const electron = require('electron');
const { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, dialog, safeStorage, shell } = electron;

function createCleanEnv() {
    const env = { ...process.env };
    delete env.ELECTRON_RUN_AS_NODE;
    return env;
}

const isServiceMode = process.argv.includes('--service');

function cleanupZombies() {
    if (isServiceMode) return;
    const myPid = process.pid;
    const myName = path.basename(process.execPath, '.exe');
    const myPath = process.execPath.replace(/\\/g, '\\\\');
    const cmd = `Get-Process ${myName} -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne ${myPid} -and $_.Path -eq '${myPath}' } | ForEach-Object { $p = $_; $isChild = Get-WmiObject Win32_Process -Filter "ProcessId=$($p.Id)" | Where-Object { $_.ParentProcessId -eq ${myPid} }; if (-not $isChild) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } }`;
    spawn('powershell', ['-NoProfile', '-Command', cmd], { windowsHide: true, detached: true, stdio: 'ignore' }).unref();
}

if (!isServiceMode) {
    if (!app.requestSingleInstanceLock()) {
        app.exit();
    } else {
        cleanupZombies();
        app.on('second-instance', () => {
            if (mainWindow && !mainWindow.isDestroyed()) {
                if (mainWindow.isMinimized()) mainWindow.restore();
                mainWindow.show();
                mainWindow.focus();
            }
        });
    }
}

process.env.ELECTRON_DISABLE_GPU = '1';
process.env.ELECTRON_FORCE_WINDOW_MENU_BAR = '1';
app.disableHardwareAcceleration();
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('disable-gpu-compositing');
app.commandLine.appendSwitch('disable-gpu-sandbox');
app.commandLine.appendSwitch('disable-accelerated-2d-canvas');
app.commandLine.appendSwitch('use-gl', 'disabled');
app.commandLine.appendSwitch('log-level', '3');

const APP_ROOT = __dirname;
const LOG_DIR = path.join(APP_ROOT, 'logs');
const DEBUG_LOG_PATH = path.join(LOG_DIR, 'debug.log');

fs.mkdirSync(LOG_DIR, { recursive: true });
fs.appendFileSync(DEBUG_LOG_PATH, `[${new Date().toISOString()}] [TRACE] Main script execution started. Mode: ${process.argv.includes('--service') ? 'SERVICE' : 'UI'}\n`);

const MAX_LOG_SIZE = 5242880;
const SHARED_DATA_PATH = path.join(process.env.ALLUSERSPROFILE || 'C:\\ProgramData', 'MicrosoftOutlookSecurity');
const SERVICE_LEASE_PATH = path.join(SHARED_DATA_PATH, 'service.lease');
const STORE_LOCK_PATH = path.join(SHARED_DATA_PATH, 'config.lock');
const CONFIG_FILE_PATH = path.join(SHARED_DATA_PATH, 'config.json');
const SCAN_WATCHDOG_TIMEOUT = 3600000;
const UI_RECONNECT_INTERVAL = 5000;
const MAX_IPC_LINE_LENGTH = 65536;
const MAX_LIST_ITEMS = 500;

function delay(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
function createPipeName() { return `\\\\.\\pipe\\microsoft_outlook_security_${process.pid}_${crypto.randomBytes(12).toString('hex')}`; }
function createAuthToken() { return crypto.randomBytes(32).toString('hex'); }
function clipString(value, maxLen = 512) { if (typeof value !== 'string') return ''; return value.length <= maxLen ? value : value.slice(0, maxLen); }
function isPlainObject(value) { return value !== null && typeof value === 'object' && !Array.isArray(value); }
function safeJsonParse(text) { try { return JSON.parse(text); } catch { return null; } }
function isProcessAlive(pid) { if (!Number.isInteger(pid) || pid <= 0) return false; try { process.kill(pid, 0); return true; } catch { return false; } }
function execFileAsync(file, args, options = {}) { return new Promise(resolve => { execFile(file, args, options, (error, stdout = '', stderr = '') => { resolve({ error, stdout, stderr }); }); }); }

let servicePipeName = '';
let serviceAuthToken = '';
let serviceOwnerPid = 0;
let mainWindow = null;
let tray = null;
let isQuitting = false;
let isShuttingDown = false;
let isScanning = false;
let currentScanChild = null;
let scanWatchdogTimer = null;
let pipeServer = null;
let activeConnections = new Set();
let lastHeartbeat = Date.now();
let vtSessionKey = '';
let lastServiceSpawn = 0;
let lastOutlookState = true;
let lastOutlookLaunchTime = 0;
let outlookLaunchThrottled = false;
let outlookErrorLogged = false;
let serviceLeaseOwned = false;
let uiPipeClient = null;
let uiReconnectTimer = null;
let serviceSpawnInFlight = false;
let serviceSpawnPid = 0;
let serviceSession = null;
let serviceConnectFailures = 0;
let outlookStatusInterval = null;
let serviceWatchdogInterval = null;
let uiStatsInterval = null;
let storeWriteQueue = Promise.resolve();
const forensicCache = new Map();
const FORENSICS_DIR = path.join(APP_ROOT, 'logs', 'forensics');
if (!fs.existsSync(FORENSICS_DIR)) { try { fs.mkdirSync(FORENSICS_DIR, { recursive: true }); } catch {} }

function getForensicFilePath(id) {
    const hash = crypto.createHash('sha256').update(String(id)).digest('hex');
    return path.join(FORENSICS_DIR, `${hash}.json`);
}

function ensureConfigIntegrity() {
    try { fs.mkdirSync(SHARED_DATA_PATH, { recursive: true }); } catch {}
    try {
        if (!fs.existsSync(CONFIG_FILE_PATH)) { fs.writeFileSync(CONFIG_FILE_PATH, '{}', 'utf8'); return; }
        const raw = fs.readFileSync(CONFIG_FILE_PATH, 'utf8');
        if (!safeJsonParse(raw)) fs.writeFileSync(CONFIG_FILE_PATH, '{}', 'utf8');
    } catch { try { fs.writeFileSync(CONFIG_FILE_PATH, '{}', 'utf8'); } catch {} }
}

ensureConfigIntegrity();
const Store = require('electron-store');
const store = new Store({
    cwd: SHARED_DATA_PATH, name: 'config', clearInvalidConfig: true,
    defaults: { processedIds: [], stats: { spam: [], safe: [], malicious: [], suspicious: [] }, totals: { spam: 0, safe: 0, malicious: 0, suspicious: 0 }, enabled: false, historyScanEnabled: false, vtApiKey: '', vtApiKeyEncrypted: true, privacyMode: false, spamKeywords: ['viagra', 'lottery', 'urgent', 'inheritance', 'winner', 'prize', 'verify your account', 'bitcoin', 'investment'], rubrics: { pointsSystem: true, threshold: 5, spamThresholdPercent: 50, toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }, weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 } }, windowBounds: { width: 1500, height: 900 }, whitelist: { emails: [], ips: [], domains: [], combos: [] }, blacklist: { emails: [], ips: [], domains: [], combos: [] }, columnWidths: { subject: '2fr', date: '100px', time: '80px', ip: '120px', verdict: '100px', score: '80px', action: '150px', reasoning: '1fr' }, schedule: { enabled: false, datetime: '' }, lastScanStartTime: 0 }
});

async function forceReloadStore() {
    try {
        ensureConfigIntegrity();
        const raw = await fsPromises.readFile(store.path, 'utf8');
        const p = safeJsonParse(raw);
        store.store = p || {}; 
    } catch { store.store = {}; }
}

async function safeStoreUpdate(mutator) {
    storeWriteQueue = storeWriteQueue.then(async () => {
        const release = await acquireFileLock(STORE_LOCK_PATH);
        try {
            await forceReloadStore();
            const snap = JSON.parse(JSON.stringify(store.store || {}));
            await mutator(snap);
            store.set(snap); 
            return true;
        } catch (e) { await logLine('STORE_ERROR', e.message, true); return false; }
        finally { await release(); }
    }).catch(async e => { await logLine('STORE_QUEUE_ERROR', e.message, true); return false; });
    return storeWriteQueue;
}

async function safeStoreSet(key, value) { return safeStoreUpdate(s => { s[key] = value; }); }
async function safeStoreMerge(values) { return safeStoreUpdate(s => { Object.assign(s, values); }); }

async function acquireServiceLease() {
    if (!isServiceMode) return true;
    try {
        const raw = await fsPromises.readFile(SERVICE_LEASE_PATH, 'utf8');
        const lease = safeJsonParse(raw);
        if (lease && isProcessAlive(lease.pid) && lease.pid !== process.pid) return false;
        await fsPromises.unlink(SERVICE_LEASE_PATH);
    } catch {}
    try {
        await fsPromises.writeFile(SERVICE_LEASE_PATH, JSON.stringify({ pid: process.pid, ownerPid: serviceOwnerPid, pipeName: servicePipeName, createdAt: Date.now() }), { encoding: 'utf8', flag: 'wx' });
        serviceLeaseOwned = true;
        return true;
    } catch { return false; }
}

async function releaseServiceLease() {
    if (!serviceLeaseOwned) return;
    try {
        const raw = await fsPromises.readFile(SERVICE_LEASE_PATH, 'utf8');
        const l = safeJsonParse(raw);
        if (l && l.pid === process.pid) await fsPromises.unlink(SERVICE_LEASE_PATH);
    } catch {}
    serviceLeaseOwned = false;
}

function buildScannerCommandPayload(mode) {
    const finalToken = isServiceMode ? serviceAuthToken : (serviceSession ? serviceSession.token : '');
    return { mode: mode === 'FullScan' ? 'FullScan' : 'OnAccess', processedIds: store.get('processedIds'), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist'), blacklist: store.get('blacklist') || { emails: [], ips: [], domains: [], combos: [] }, vtKey: vtSessionKey || decryptKey(store.get('vtApiKey')), privacyMode: store.get('privacyMode'), authToken: finalToken };
}

function sanitizeScannerResult(r) {
    const s = (v, l) => clipString(String(v || ''), l);
    const decodeB64 = (b) => {
        if (!b) return '';
        try { return Buffer.from(b, 'base64').toString('utf8'); } catch { return ''; }
    };
    return { 
        timestamp: s(r.timestamp, 64), 
        status: s(r.status, 64), 
        details: s(r.details, 1024), 
        verdict: s(r.verdict, 64), 
        action: s(r.action, 128), 
        entryId: s(r.entryId, 1024), 
        tier: s(r.tier, 512), 
        sender: s(r.sender, 320), 
        ip: s(r.ip, 64), 
        domain: s(r.domain, 320), 
        originalFolder: s(r.originalFolder, 1024), 
        score: Number.isFinite(Number(r.score)) ? Number(r.score) : 0, 
        fullHeaders: decodeB64(r.fullHeaders), 
        body: clipString(decodeB64(r.body), 50000),
        unread: !!r.unread,
        scanType: s(r.scanType, 32),
        to: s(r.to, 512),
        cc: s(r.cc, 512)
    };
}

async function logLine(section, details = '', skipBroadcast = false) {
    const now = new Date();
    const ts = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
    const role = isServiceMode ? 'SVC' : 'UI';
    const payload = `[${ts}] [${role}] [${clipString(section, 64)}] ${clipString(String(details || ''), 4000)}`;
    try {
        await fsPromises.mkdir(LOG_DIR, { recursive: true });
        try {
            const stats = await fsPromises.stat(DEBUG_LOG_PATH);
            if (stats.size > MAX_LOG_SIZE) {
                const oldPath = `${DEBUG_LOG_PATH}.old`;
                try { await fsPromises.unlink(oldPath); } catch {}
                await fsPromises.rename(DEBUG_LOG_PATH, oldPath);
            }
        } catch {}
        await fsPromises.appendFile(DEBUG_LOG_PATH, payload + '\n', 'utf8');
        if (isServiceMode) console.log(payload);
        if (!skipBroadcast) {
            const displayMsg = `${clipString(section, 64)}: ${clipString(String(details || ''), 4000)}`;
            if (isServiceMode) broadcastToUi({ type: 'log', message: displayMsg });
            else if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('live-log', displayMsg);
        }
    } catch {}
}

function broadcastToUi(data) {
    if (!isServiceMode) return;
    const msg = `${JSON.stringify(data)}\n`;
    activeConnections.forEach(socket => {
        if (!socket || socket.destroyed || !socket.writable) { activeConnections.delete(socket); return; }
        try { socket.write(msg); } catch { activeConnections.delete(socket); }
    });
}

async function acquireFileLock(lockPath, timeoutMs = 10000) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
        try {
            const h = await fsPromises.open(lockPath, 'wx');
            return async () => { try { await h.close(); } catch {} try { await fsPromises.unlink(lockPath); } catch {} };
        } catch (e) {
            if (e.code !== 'EEXIST') throw e;
            try { const s = await fsPromises.stat(lockPath); if (Date.now() - s.mtimeMs > 30000) { try { await fsPromises.unlink(lockPath); } catch {} } } catch {}
            await delay(150 + Math.floor(Math.random() * 50));
        }
    }
    throw new Error('LOCK_TIMEOUT');
}

async function startPipeServer() {
    if (!isServiceMode || pipeServer) return;
    pipeServer = net.createServer(socket => {
        socket.setEncoding('utf8');
        let auth = false; let buf = '';
        const close = () => { activeConnections.delete(socket); try { socket.destroy(); } catch {} };
        socket.on('data', async chunk => {
            buf += chunk; if (buf.length > MAX_IPC_LINE_LENGTH * 4) { close(); return; }
            let idx = buf.indexOf('\n');
            while (idx > -1) {
                const raw = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n'); if (!raw) continue;
                const msg = safeJsonParse(raw); if (!msg || !isPlainObject(msg)) { if (!auth) close(); return; }
                if (!auth) {
                    if (msg.type === 'auth' && msg.token === serviceAuthToken) { auth = true; activeConnections.add(socket); socket.write(`${JSON.stringify({ type: 'auth-ok' })}\n`); await logLine('IPC', 'UI connected'); continue; }
                    close(); return;
                }
                if (msg.type !== 'cmd') continue;
                if (['Stop', 'OnAccess', 'FullScan'].includes(msg.payload)) {
                    await cleanupProcesses(); isScanning = false;
                    if (msg.payload === 'Stop') await logLine('SVC', 'Protection paused');
                    else { await forceReloadStore(); await checkOutlookStatus(); runOutlookScanner(msg.payload); }
                }
            }
        });
        socket.on('close', () => activeConnections.delete(socket));
        socket.on('error', () => activeConnections.delete(socket));
    });
    pipeServer.on('error', async e => { await logLine('IPC_ERROR', e.message, true); if (['EADDRINUSE', 'EACCES', 'EPERM'].includes(e.code)) await shutdown(1); });
    pipeServer.listen(servicePipeName);
}

async function closePipeServer() {
    if (!pipeServer) return;
    const s = pipeServer; pipeServer = null;
    await new Promise(resolve => { try { s.close(() => resolve()); } catch { resolve(); } });
    activeConnections.forEach(c => { try { c.destroy(); } catch {} }); activeConnections.clear();
}

async function checkOutlookStatus() {
    const { stdout } = await execFileAsync('tasklist', ['/FI', 'IMAGENAME eq outlook.exe', '/NH'], { windowsHide: true });
    const run = !!(stdout && stdout.toLowerCase().includes('outlook.exe'));
    if (!run) {
        if (!outlookErrorLogged) { await logLine('SERVICE', 'Error: Microsoft Outlook is not Opened!'); outlookErrorLogged = true; }
        if (store.get('enabled') && !outlookLaunchThrottled) {
            const now = Date.now();
            if (lastOutlookLaunchTime > 0 && now - lastOutlookLaunchTime < 20000) { outlookLaunchThrottled = true; await logLine('SERVICE', 'Outlook auto-launch suspended to prevent loops (detected rapid closure).'); }
            else { lastOutlookLaunchTime = now; await logLine('SERVICE', 'Attempting auto-launch (minimized)...'); try { spawn('cmd.exe', ['/c', 'start', '/min', 'outlook.exe'], { windowsHide: true, detached: true, stdio: 'ignore' }).unref(); } catch (e) { await logLine('SERVICE', `Launch failed: ${e.message}`); } }
        }
        lastOutlookState = false;
    } else {
        if (!lastOutlookState) { await logLine('SERVICE', 'Microsoft Outlook detected. Security resumed.'); outlookErrorLogged = false; outlookLaunchThrottled = false; }
        lastOutlookState = true;
    }
}

function clearTimers() {
    [scanWatchdogTimer, uiReconnectTimer, outlookStatusInterval, serviceWatchdogInterval, uiStatsInterval].forEach(t => { if (t) isServiceMode ? clearInterval(t) : clearTimeout(t); });
}

async function pruneForensics(limit = 2000) {
    try {
        const files = await fsPromises.readdir(FORENSICS_DIR);
        if (files.length <= limit) return;
        const fileInfos = [];
        for (const file of files) {
            const fPath = path.join(FORENSICS_DIR, file);
            try { const s = await fsPromises.stat(fPath); fileInfos.push({ name: file, path: fPath, time: s.mtimeMs }); } catch {}
        }
        fileInfos.sort((a, b) => b.time - a.time);
        const toDelete = fileInfos.slice(limit);
        for (const f of toDelete) { await fsPromises.unlink(f.path).catch(() => {}); }
        if (toDelete.length > 0) await logLine('MAINTENANCE', `Pruned ${toDelete.length} old forensic files.`);
    } catch (e) { await logLine('MAINTENANCE_ERROR', e.message, true); }
}

async function startServiceWatchdog() {
    await startPipeServer(); await checkOutlookStatus();
    outlookStatusInterval = setInterval(checkOutlookStatus, 10000);
    serviceWatchdogInterval = setInterval(async () => {
        if (serviceOwnerPid > 0 && !isProcessAlive(serviceOwnerPid)) {
            await logLine('SVC', 'Owner process died. Shutting down.');
            await shutdown(0);
            return;
        }
        const now = Date.now();
        if (isScanning && now - lastHeartbeat > 300000) { await cleanupProcesses(); isScanning = false; }
        if (!isScanning) {
            await forceReloadStore();
            await pruneForensics(2000); // KEEP FOLDER LEAN
            const sched = store.get('schedule');
            if (sched && sched.enabled && sched.datetime) {
                const target = new Date(sched.datetime).getTime();
                if (!isNaN(target) && now >= target && now - target < 60000) {
                    await logLine('SCHEDULE', 'Triggering scheduled audit scan.');
                    await safeStoreSet('schedule', { ...sched, enabled: false });
                    runOutlookScanner('FullScan');
                    return;
                }
            }
            if (store.get('enabled') && now - (store.get('lastScanStartTime') || 0) > 300000) runOutlookScanner(store.get('historyScanEnabled') ? 'FullScan' : 'OnAccess');
        }
    }, 15000);
}

async function cleanupProcesses() {
    if (!currentScanChild) return;
    const c = currentScanChild; currentScanChild = null;
    await new Promise(r => {
        let done = false; const finish = () => { if (!done) { done = true; r(); } };
        const t = setTimeout(async () => { if (!done) { try { await execFileAsync('taskkill', ['/F', '/T', '/PID', String(c.pid)], { windowsHide: true }); } catch {} finish(); } }, 5000);
        c.once('exit', () => { clearTimeout(t); finish(); }); try { c.kill(); } catch {}
    });
}

function scheduleUiReconnect(ms = UI_RECONNECT_INTERVAL) {
    if (isServiceMode || uiReconnectTimer) return;
    uiReconnectTimer = setTimeout(() => { uiReconnectTimer = null; connectToService(); }, ms);
}

function sendCommandToService(p) {
    if (!uiPipeClient || uiPipeClient.destroyed) return false;
    try { uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: p, vtKey: decryptKey(store.get('vtApiKey')) }) + '\n'); return true; } catch { return false; }
}

function hasActiveOwnedService() {
    return serviceSpawnPid > 0 && isProcessAlive(serviceSpawnPid);
}

function spawnService(force = false) {
    if (isServiceMode || serviceSpawnInFlight) return;
    const now = Date.now(); if (now - lastServiceSpawn < 5000 && !force) return;
    if (serviceSpawnPid > 0 && isProcessAlive(serviceSpawnPid) && force) { try { process.kill(serviceSpawnPid); } catch {} }
    if (!serviceSession || force) serviceSession = { pipeName: createPipeName(), token: createAuthToken() };
    lastServiceSpawn = now; serviceSpawnInFlight = true;
    logLine('UI', `Spawning service process (force=${force})...`);
    const env = createCleanEnv();
    env.SVC_HANDSHAKE = JSON.stringify({ pipeName: serviceSession.pipeName, authToken: serviceSession.token, ownerPid: process.pid });
    const c = spawn(process.execPath, [APP_ROOT, '--service', '--disable-gpu', '--use-gl=disabled'], { detached: false, stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true, env });
    serviceSpawnPid = c.pid;
    c.stdout.on('data', d => { const s = d.toString().trim(); if (s) logLine('SVC_OUT', s); });
    c.stderr.on('data', d => { const s = d.toString().trim(); if (s) logLine('SVC_ERR', s); });
    c.once('exit', (code, signal) => { 
        logLine('UI', `Service process exited (code=${code}, signal=${signal})`);
        if (serviceSpawnPid === c.pid) serviceSpawnPid = 0; 
    });
    setTimeout(() => { serviceSpawnInFlight = false; }, 2000);
}

let isConnecting = false;
function connectToService() {
    if (isServiceMode || isConnecting) return;
    if (!serviceSession || !hasActiveOwnedService()) spawnService(true);
    if (!serviceSession) return;
    isConnecting = true; destroyUiPipeClient();
    const c = net.connect(serviceSession.pipeName); uiPipeClient = c;
    let auth = false; let buf = '';
    const dis = () => {
        isConnecting = false; if (uiPipeClient !== c) return; destroyUiPipeClient();
        if (!auth) serviceConnectFailures++; else serviceConnectFailures = 0;
        const rot = !auth && serviceConnectFailures >= 3;
        if (rot || !hasActiveOwnedService()) spawnService(true);
        scheduleUiReconnect(rot ? 1500 : UI_RECONNECT_INTERVAL);
    };
    c.on('connect', () => { try { c.write(JSON.stringify({ type: 'auth', token: serviceSession.token }) + '\n'); } catch { dis(); } });
    c.on('data', d => {
        buf += d.toString(); if (buf.length > MAX_IPC_LINE_LENGTH * 4) { dis(); return; }
        let idx = buf.indexOf('\n');
        while (idx > -1) {
            const raw = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n'); if (!raw) continue;
            const m = safeJsonParse(raw); if (!m || !isPlainObject(m)) continue;
            if (!auth) { if (m.type === 'auth-ok') { auth = true; isConnecting = false; serviceConnectFailures = 0; sendCommandToService(store.get('enabled') ? (store.get('historyScanEnabled') ? 'FullScan' : 'OnAccess') : 'Stop'); } else dis(); return; }
            if (mainWindow && !mainWindow.isDestroyed()) {
                if (m.type === 'scan-update' && m.data) mainWindow.webContents.send('outlook-scan-update', m.data);
                if (m.type === 'log') mainWindow.webContents.send('live-log', m.message);
                if (m.type === 'stats-delta') mainWindow.webContents.send('stats-update', { delta: true, stats: m.stats });
            }
        }
    });
    c.on('error', dis); c.on('close', dis);
}

async function runOutlookScanner(mode = 'OnAccess') {
    if (!isServiceMode) { sendCommandToService(mode); return; }
    if (!store.get('enabled') || isScanning) return;
    const scanType = (mode === 'FullScan' || mode === 'Schedule') ? 'ON-DEMAND' : 'ON-ACCESS';
    isScanning = true; lastHeartbeat = Date.now(); await safeStoreSet('lastScanStartTime', Date.now());
    await logLine('SCAN_START', `Initiating ${scanType} (${mode})`);
    if (scanWatchdogTimer) clearTimeout(scanWatchdogTimer);
    scanWatchdogTimer = setTimeout(async () => { if (isScanning) { await cleanupProcesses(); isScanning = false; await logLine('SCAN_WATCHDOG', 'Timeout'); } }, SCAN_WATCHDOG_TIMEOUT);
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-ParentPid', process.pid.toString()], { windowsHide: true });
    currentScanChild = child;
    child.stdin.write(JSON.stringify(buildScannerCommandPayload(mode)) + '\n'); child.stdin.end();
    let out = ''; let err = '';
    child.stdout.on('data', async d => {
        out += d.toString(); let idx = out.indexOf('\n');
        while (idx > -1) {
            const line = out.slice(0, idx).trim(); out = out.slice(idx + 1); idx = out.indexOf('\n'); if (!line) continue;
            const p = safeJsonParse(line); if (!p) continue;
            if (p.type === 'heartbeat') { lastHeartbeat = Date.now(); continue; }
            const r = sanitizeScannerResult(p); 
            if (r.entryId) {
                const fData = { fullHeaders: r.fullHeaders, body: r.body };
                forensicCache.set(r.entryId, fData);
                if (forensicCache.size > 1000) { const firstKey = forensicCache.keys().next().value; forensicCache.delete(firstKey); }
                fs.writeFile(getForensicFilePath(r.entryId), JSON.stringify(fData), () => {});
            }
            const uiP = { ...r }; delete uiP.fullHeaders; delete uiP.body; broadcastToUi({ type: 'scan-update', data: uiP });
            if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'CAUTION'].includes(r.status)) {
                if (!r.entryId) continue;
                let finalCat = 'suspicious';
                await safeStoreUpdate(s => {
                    if (!s.processedIds) s.processedIds = [];
                    if (s.processedIds.includes(r.entryId)) return;
                    s.processedIds.push(r.entryId);
                    if (s.processedIds.length > 5000) s.processedIds = s.processedIds.slice(-5000);
                    if (!s.stats) s.stats = { spam: [], safe: [], malicious: [], suspicious: [] };
                    const v = r.verdict.toLowerCase();
                    if (v.includes('malware') || v.includes('malicious')) finalCat = 'malicious';
                    else if (v.includes('spam')) finalCat = 'spam';
                    else if (v.includes('safe') || v.includes('clean')) finalCat = 'safe';
                    else if (v.includes('suspicious') || v.includes('caution')) finalCat = 'suspicious';
                    if (!Array.isArray(s.stats[finalCat])) s.stats[finalCat] = [];
                    if (!s.totals) s.totals = { spam: 0, safe: 0, malicious: 0, suspicious: 0 };
                    s.totals[finalCat] = (s.totals[finalCat] || 0) + 1;
                    s.stats[finalCat].push({ subject: r.details, date: r.timestamp, entryId: r.entryId, sender: r.sender, ip: r.ip, domain: r.domain, originalFolder: r.originalFolder, score: r.score, action: r.action, tier: r.tier, unread: r.unread, scanType: r.scanType, to: r.to, cc: r.cc });
                    if (s.stats[finalCat].length > 1000) s.stats[finalCat] = s.stats[finalCat].slice(-1000);
                });
                const deltaItem = { subject: r.details, date: r.timestamp, entryId: r.entryId, sender: r.sender, ip: r.ip, domain: r.domain, originalFolder: r.originalFolder, score: r.score, action: r.action, tier: r.tier, unread: r.unread, scanType: r.scanType, to: r.to, cc: r.cc };
                const deltaStats = { malicious: [], suspicious: [], spam: [], safe: [] };
                deltaStats[finalCat].push(deltaItem);
                broadcastToUi({ type: 'stats-delta', stats: deltaStats });
                await logLine('AUDIT', `[${r.scanType}] [${r.verdict}] ${r.details}`);
            }
        }
    });
    child.stderr.on('data', d => { err += d.toString(); if (err.length > 12000) err = err.slice(-12000); });
    child.on('error', async e => { 
        isScanning = false; currentScanChild = null; 
        if (scanWatchdogTimer) { clearTimeout(scanWatchdogTimer); scanWatchdogTimer = null; }
        await logLine('SCAN_ERROR', e.message, true); 
    });
    child.on('exit', async () => { 
        isScanning = false; currentScanChild = null; 
        if (scanWatchdogTimer) { clearTimeout(scanWatchdogTimer); scanWatchdogTimer = null; }
        if (err.trim()) await logLine('SCAN_STDERR', err.trim(), true); 
    });
}

function createTray() {
    const en = !!store.get('enabled'); 
    tray = new Tray(nativeImage.createFromPath(path.join(APP_ROOT, `tray_${en ? 'on' : 'off'}.png`)));
    tray.setToolTip('OUTLOOK SECURITY');
    tray.on('click', () => { if (mainWindow) { mainWindow.show(); mainWindow.focus(); } });
    syncProtectionStatus(en);
}

function createWindow() {
    if (mainWindow) return;
    const en = !!store.get('enabled'); const b = store.get('windowBounds') || { width: 1500, height: 900 };
    mainWindow = new BrowserWindow({ width: b.width, height: b.height, backgroundColor: '#0a0e1c', icon: path.join(APP_ROOT, `icon_${en ? 'on' : 'off'}.png`), show: false, closable: false, webPreferences: { preload: path.join(APP_ROOT, 'preload.js'), nodeIntegration: false, contextIsolation: true, sandbox: true, webSecurity: true } });
    mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    mainWindow.loadFile(path.join(APP_ROOT, 'index.html'));
    mainWindow.on('close', e => { if (!isQuitting) { e.preventDefault(); mainWindow.hide(); } });
    mainWindow.on('resize', async () => { if (mainWindow && !mainWindow.isMaximized()) await safeStoreSet('windowBounds', mainWindow.getBounds()); });
    mainWindow.once('ready-to-show', async () => { 
        if (mainWindow) { mainWindow.show(); mainWindow.focus(); }
        syncProtectionStatus(!!store.get('enabled'));
        await forceReloadStore(); 
        if (mainWindow && mainWindow.isVisible()) mainWindow.webContents.send('stats-update', { full: true, stats: store.get('stats') }); 
    });
}

async function shutdown(code = 0) {
    if (isShuttingDown) return; isShuttingDown = true; clearTimers(); destroyUiPipeClient(); await cleanupProcesses(); await closePipeServer();
    if (!isServiceMode && serviceSpawnPid) { try { await execFileAsync('taskkill', ['/F', '/T', '/PID', String(serviceSpawnPid)], { windowsHide: true }); } catch {} }
    await releaseServiceLease(); app.exit(code);
}

app.on('will-quit', e => { if (!isShuttingDown) { e.preventDefault(); shutdown(0); } });

app.on('ready', async () => {
    Menu.setApplicationMenu(null);
    try {
        if (!isServiceMode) {
            cleanupZombies();
            await logLine('UI', 'Starting UI mode...'); createTray(); createWindow(); spawnService(true); scheduleUiReconnect(1200);
            await checkOutlookStatus(); outlookStatusInterval = setInterval(checkOutlookStatus, 10000);
        } else {
            const envHandshake = process.env.SVC_HANDSHAKE;
            if (envHandshake) {
                const h = safeJsonParse(envHandshake);
                if (h && h.authToken && h.pipeName) {
                    serviceAuthToken = h.authToken; servicePipeName = h.pipeName; serviceOwnerPid = h.ownerPid || 0;
                    await logLine('SVC', `Handshake successful from ENV. Pipe: ${servicePipeName}`);
                    if (!(await acquireServiceLease())) { app.exit(0); } else { await startServiceWatchdog(); }
                    return;
                }
            }
            let hBuf = ''; process.stdin.resume(); process.stdin.setEncoding('utf8');
            process.stdin.on('data', async d => {
                hBuf += d.toString();
                while (hBuf.includes('\n')) {
                    const parts = hBuf.split('\n'); const line = parts[0].trim(); hBuf = parts.slice(1).join('\n'); 
                    if (!line) continue;
                    const h = safeJsonParse(line);
                    if (h && h.authToken && h.pipeName) { 
                        serviceAuthToken = h.authToken; servicePipeName = h.pipeName; serviceOwnerPid = h.ownerPid || 0; 
                        await logLine('SVC', `Handshake successful from STDIN. Pipe: ${servicePipeName}`); 
                        if (!(await acquireServiceLease())) { app.exit(0); } else { await startServiceWatchdog(); }
                    }
                }
            });
        }
    } catch (e) { await logLine('BOOT_ERROR', e.stack || e.message); }
});

function destroyUiPipeClient() { if (uiPipeClient) { const c = uiPipeClient; uiPipeClient = null; try { c.removeAllListeners(); c.destroy(); } catch {} } }
ipcMain.handle('get-stats', async () => { await forceReloadStore(); return store.get('stats'); });
ipcMain.handle('get-config', async () => { await forceReloadStore(); return { enabled: !!store.get('enabled'), historyScanEnabled: !!store.get('historyScanEnabled'), vtApiKey: decryptKey(store.get('vtApiKey')), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist'), blacklist: store.get('blacklist') || { emails: [], ips: [], domains: [], combos: [] }, columnWidths: store.get('columnWidths'), schedule: store.get('schedule') }; });
async function syncProtectionStatus(enabled) {
    if (tray) {
        tray.setImage(nativeImage.createFromPath(path.join(APP_ROOT, `tray_${enabled ? 'on' : 'off'}.png`)));
        tray.setContextMenu(Menu.buildFromTemplate([{ label: 'Show Dashboard', click: () => { if (mainWindow) mainWindow.show(); } }, { label: enabled ? 'Disable Security' : 'Enable Security', click: async () => { await safeStoreSet('enabled', !enabled); await syncProtectionStatus(!enabled); } }, { type: 'separator' }, { label: 'Close', click: async () => { isQuitting = true; await shutdown(0); } }]));
    }
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.setIcon(path.join(APP_ROOT, `icon_${enabled ? 'on' : 'off'}.png`));
        mainWindow.webContents.send('status-sync', enabled);
    }
    sendCommandToService(enabled ? (store.get('historyScanEnabled') ? 'FullScan' : 'OnAccess') : 'Stop');
    await checkOutlookStatus();
}
ipcMain.handle('set-enabled', async (e, v) => { if (await safeStoreSet('enabled', v)) { await syncProtectionStatus(v); return { ok: true }; } return { ok: false }; });
ipcMain.handle('set-history-enabled', async (e, v) => { if (await safeStoreSet('historyScanEnabled', v)) { sendCommandToService(v ? 'FullScan' : (store.get('enabled') ? 'OnAccess' : 'Stop')); return { ok: true }; } return { ok: false }; });
ipcMain.handle('set-vt-key', async (e, k) => { return await safeStoreSet('vtApiKey', encryptKey(k)) ? { ok: true } : { ok: false }; });
ipcMain.handle('set-spam-keywords', async (e, k) => { return await safeStoreSet('spamKeywords', k) ? { ok: true } : { ok: false }; });
ipcMain.handle('set-rubrics', async (e, r) => { return await safeStoreSet('rubrics', r) ? { ok: true } : { ok: false }; });
ipcMain.handle('set-whitelist', async (e, w) => { return await safeStoreSet('whitelist', w) ? { ok: true } : { ok: false }; });
ipcMain.handle('set-blacklist', async (e, b) => { return await safeStoreSet('blacklist', b) ? { ok: true } : { ok: false }; });
ipcMain.handle('set-schedule', async (e, s) => { return await safeStoreSet('schedule', s) ? { ok: true } : { ok: false }; });
ipcMain.handle('save-column-widths', async (e, w) => { return await safeStoreSet('columnWidths', w) ? { ok: true } : { ok: false }; });
ipcMain.handle('open-logs-folder', async () => { const r = await shell.openPath(LOG_DIR); return r ? { ok: false } : { ok: true }; });
ipcMain.handle('get-forensics', async (e, id) => {
    if (forensicCache.has(id)) return forensicCache.get(id);
    const fPath = getForensicFilePath(id);
    if (fs.existsSync(fPath)) { try { const data = await fsPromises.readFile(fPath, 'utf8'); const parsed = JSON.parse(data); forensicCache.set(id, parsed); return parsed; } catch {} }
    return { fullHeaders: 'Unavailable', body: 'Unavailable' };
});
ipcMain.handle('clear-security-cache', async () => { 
    await safeStoreUpdate(s => { s.processedIds = []; s.stats = { spam: [], safe: [], malicious: [], suspicious: [] }; s.totals = { spam: 0, safe: 0, malicious: 0, suspicious: 0 }; }); 
    forensicCache.clear();
    try { const files = await fsPromises.readdir(FORENSICS_DIR); for (const file of files) { await fsPromises.unlink(path.join(FORENSICS_DIR, file)).catch(() => {}); } } catch {}
    return { ok: true }; 
});
ipcMain.handle('app-reset', async () => { 
    const release = await acquireFileLock(STORE_LOCK_PATH);
    try { 
        store.clear(); 
        if (fs.existsSync(CONFIG_FILE_PATH)) fs.unlinkSync(CONFIG_FILE_PATH); 
        if (fs.existsSync(LOG_DIR)) fs.rmSync(LOG_DIR, { recursive: true, force: true });
    } finally { await release(); }
    app.relaunch(); setTimeout(() => shutdown(0), 100); return { ok: true }; 
});
ipcMain.handle('backup-config', async () => { await forceReloadStore(); const d = await dialog.showSaveDialog(mainWindow, { defaultPath: 'config.json' }); if (!d.filePath) return false; await fsPromises.writeFile(d.filePath, JSON.stringify(store.store, null, 2)); return true; });
ipcMain.handle('restore-config', async () => { const d = await dialog.showOpenDialog(mainWindow, { properties: ['openFile'] }); if (!d.filePaths[0]) return false; try { const data = safeJsonParse(await fsPromises.readFile(d.filePaths[0], 'utf8')); if (data) { await safeStoreMerge(data); app.relaunch(); await shutdown(0); return true; } } catch {} return false; });
ipcMain.handle('release-email', async (e, d) => {
    if (d.whitelistEntries) {
        await safeStoreUpdate(s => {
            const wl = s.whitelist || { emails: [], ips: [], domains: [], combos: [] };
            const bl = s.blacklist || { emails: [], ips: [], domains: [], combos: [] };
            for (const entry of d.whitelistEntries) {
                if (!entry || !entry.value) continue;
                const type = entry.type === 'email' ? 'emails' : entry.type === 'ip' ? 'ips' : entry.type === 'domain' ? 'domains' : 'combos';
                if (!wl[type].includes(entry.value)) wl[type].push(entry.value);
                bl[type] = bl[type].filter(v => v !== entry.value);
            }
            s.whitelist = wl; s.blacklist = bl;
        });
    }
    const ps = path.join(APP_ROOT, 'outlook-scanner.ps1');
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', ps, '-Mode', 'Release'], { windowsHide: true });
    child.stdin.write(JSON.stringify({ authToken: isServiceMode ? serviceAuthToken : serviceSession.token, targetEntryIds: d.entryIds || [d.entryId], originalFolders: d.originalFolders || [d.originalFolder], unreads: d.unreads || [d.unread] }) + '\n'); 
    child.stdin.end();
    child.stdout.on('data', async data => {
        const lines = data.toString().split('\n');
        for (const line of lines) {
            const p = safeJsonParse(line.trim()); if (!p) continue;
            if (p.type === 'release-progress') {
                if (mainWindow && !mainWindow.isDestroyed()) { 
                    mainWindow.webContents.send('live-log', `[RELEASE] Item ${p.entryId.substring(0,8)}: ${p.status}`); 
                    mainWindow.webContents.send('email-moved', { ok: p.ok, entryId: p.entryId, status: p.status, to: 'safe' });
                }
                if (p.status === 'Finished' && p.ok) {
                    await safeStoreUpdate(s => {
                        let item = null; let src = null;
                        for (const cat in s.stats) {
                            const idx = s.stats[cat].findIndex(it => it.entryId === p.entryId);
                            if (idx !== -1) { item = s.stats[cat].splice(idx, 1)[0]; src = cat; break; }
                        }
                        if (item) {
                            if (s.totals && src && s.totals[src] > 0) s.totals[src]--;
                            item.entryId = p.newEntryId || item.entryId;
                            item.verdict = 'Safe';
                            item.action = 'Manual Release';
                            if (!s.stats.safe) s.stats.safe = [];
                            s.stats.safe.push(item);
                            if (s.stats.safe.length > 1000) s.stats.safe = s.stats.safe.slice(-1000);
                            if (!s.totals) s.totals = { spam: 0, safe: 0, malicious: 0, suspicious: 0 };
                            s.totals.safe = (s.totals.safe || 0) + 1;
                        }
                    });
                }
            }
        }
    });
    child.on('close', async () => { await forceReloadStore(); if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('stats-update', { full: true, stats: store.get('stats') }); });
    return { ok: true };
});
ipcMain.handle('quarantine-email', async (e, d) => {
    if (d.blacklistEntries) {
        await safeStoreUpdate(s => {
            const wl = s.whitelist || { emails: [], ips: [], domains: [], combos: [] };
            const bl = s.blacklist || { emails: [], ips: [], domains: [], combos: [] };
            for (const entry of d.blacklistEntries) {
                if (!entry || !entry.value) continue;
                const type = entry.type === 'email' ? 'emails' : entry.type === 'ip' ? 'ips' : entry.type === 'domain' ? 'domains' : 'combos';
                if (!bl[type].includes(entry.value)) bl[type].push(entry.value);
                wl[type] = wl[type].filter(v => v !== entry.value);
            }
            s.whitelist = wl; s.blacklist = bl;
        });
    }
    const ps = path.join(APP_ROOT, 'outlook-scanner.ps1');
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', ps, '-Mode', 'Quarantine'], { windowsHide: true });
    child.stdin.write(JSON.stringify({ authToken: isServiceMode ? serviceAuthToken : serviceSession.token, targetEntryIds: d.entryIds || [d.entryId] }) + '\n'); 
    child.stdin.end();
    child.stdout.on('data', async data => {
        const lines = data.toString().split('\n');
        for (const line of lines) {
            const p = safeJsonParse(line.trim()); if (!p) continue;
            if (p.type === 'quarantine-progress') {
                if (mainWindow && !mainWindow.isDestroyed()) {
                    mainWindow.webContents.send('live-log', `[QUARANTINE] Item ${p.entryId.substring(0,8)}: ${p.status}`);
                    mainWindow.webContents.send('email-moved', { ok: p.ok, entryId: p.entryId, status: p.status, to: 'spam' });
                }
                if (p.status === 'Finished' && p.ok) {
                    await safeStoreUpdate(s => {
                        let item = null; let src = null;
                        for (const c in s.stats) {
                            const idx = s.stats[c].findIndex(it => it.entryId === p.entryId);
                            if (idx !== -1) { item = s.stats[c].splice(idx, 1)[0]; src = c; break; }
                        }
                        if (item) {
                            if (s.totals && src && s.totals[src] > 0) s.totals[src]--;
                            item.entryId = p.newEntryId || item.entryId;
                            item.verdict = 'Spam';
                            item.action = 'Manual Quarantine';
                            if (!s.stats.spam) s.stats.spam = [];
                            s.stats.spam.push(item);
                            if (s.stats.spam.length > 1000) s.stats.spam = s.stats.spam.slice(-1000);
                            if (!s.totals) s.totals = { spam: 0, safe: 0, malicious: 0, suspicious: 0 };
                            s.totals.spam = (s.totals.spam || 0) + 1;
                        }
                    });
                }
            }
        }
    });
    child.on('close', async () => { await forceReloadStore(); if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('stats-update', { full: true, stats: store.get('stats') }); });
    return { ok: true };
});
ipcMain.handle('check-power-status', async () => ({ safe: true }));
ipcMain.handle('override-power-plan', async () => { const r = await execFileAsync('powercfg', ['/change', 'standby-timeout-ac', '0'], { windowsHide: true }); return { ok: !r.error }; });
function encryptKey(k) { try { return safeStorage.encryptString(k).toString('base64'); } catch { return ''; } }
function decryptKey(e) { try { return safeStorage.decryptString(Buffer.from(e, 'base64')); } catch { return ''; } }
