const path = require('node:path');
const fs = require('node:fs');
const fsPromises = fs.promises;
const net = require('node:net');
const crypto = require('node:crypto');
const { spawn, execFile } = require('node:child_process');
const electron = require('electron');
const { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, safeStorage, shell } = electron;

const isServiceMode = process.argv.includes('--service');
const APP_ROOT = __dirname;
const LOG_DIR = path.join(APP_ROOT, 'logs');
const FORENSICS_DIR = path.join(LOG_DIR, 'forensics');
const LOG_FILE = path.join(LOG_DIR, 'security.log');

[LOG_DIR, FORENSICS_DIR].forEach(d => { if (!fs.existsSync(d)) try { fs.mkdirSync(d, { recursive: true }); } catch {} });

function logToFile(msg) {
    const ts = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
    try { fs.appendFileSync(LOG_FILE, `[${ts}] ${msg}\n`); } catch {}
}

const USER_DATA = app.getPath('userData');
const Store = require('electron-store');
let store = null;
if (isServiceMode) {
    store = new Store({
        cwd: USER_DATA, name: 'config', clearInvalidConfig: true,
        defaults: { processedIds: [], stats: { spam: [], safe: [], malicious: [], suspicious: [] }, enabled: false, vtApiKey: '', spamKeywords: ['viagra', 'lottery', 'urgent', 'bitcoin'], rubrics: { weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }, toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }, spamThresholdPercent: 50 }, whitelist: { emails: [], ips: [], domains: [], combos: [] }, blacklist: { emails: [], ips: [], domains: [], combos: [] } }
    });
}

let mainWindow = null, tray = null, isQuitting = false, isEnabled = false;
let uiPipeClient = null, serviceSession = null, serviceSpawnInFlight = false;
let pipeServer = null, activeConnections = new Set(), isScanning = false, currentScanChild = null;
let psWorker = null;

let statsBuffer = { malicious: [], suspicious: [], spam: [], safe: [] };
let bufferTimer = null;

function broadcastToUi(data) {
    const msg = JSON.stringify(data) + '\n';
    if (isServiceMode) {
        activeConnections.forEach(s => { try { s.write(msg); } catch { activeConnections.delete(s); } });
    } else if (mainWindow && !mainWindow.isDestroyed()) {
        if (data.type === 'scan-update') mainWindow.webContents.send('outlook-scan-update', data.data);
        if (data.type === 'status-sync') { isEnabled = !!data.enabled; updateTrayState(); mainWindow.webContents.send('status-sync', data.enabled); }
        if (data.type === 'stats-update') mainWindow.webContents.send('stats-update', data.data);
        if (data.type === 'outlook-status') mainWindow.webContents.send('outlook-status', data.running);
    }
}

function flushStats() {
    if (!isServiceMode) return;
    const hasData = Object.values(statsBuffer).some(a => a.length > 0);
    if (hasData) {
        const currentStats = store.get('stats') || { malicious: [], suspicious: [], spam: [], safe: [] };
        for (const cat in statsBuffer) {
            currentStats[cat] = [...(currentStats[cat] || []), ...statsBuffer[cat]].slice(-1000);
            statsBuffer[cat] = [];
        }
        store.set('stats', currentStats);
        broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
    }
    bufferTimer = null;
}

function getPsWorker() {
    if (psWorker && !psWorker.killed) return psWorker;
    psWorker = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-Mode', 'Worker', '-ParentPid', process.pid.toString()], { windowsHide: true });
    let buf = '';
    psWorker.stdout.on('data', d => {
        buf += d.toString(); let idx = buf.indexOf('\n');
        while (idx > -1) {
            const line = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
            try { const p = JSON.parse(line); if (p) broadcastToUi({ type: 'scan-update', data: p }); } catch {}
        }
    });
    return psWorker;
}

async function runOutlookScanner() {
    if (!isServiceMode || !store.get('enabled') || isScanning) return;
    isScanning = true;
    currentScanChild = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-ParentPid', process.pid.toString()], { windowsHide: true });
    currentScanChild.stdin.write(JSON.stringify({ mode: 'OnAccess', processedIds: store.get('processedIds'), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist'), blacklist: store.get('blacklist'), vtKey: store.get('vtApiKey') ? safeStorage.decryptString(Buffer.from(store.get('vtApiKey'), 'base64')) : '' }) + '\n');
    let buf = '';
    currentScanChild.stdout.on('data', d => {
        buf += d.toString(); let idx = buf.indexOf('\n');
        while (idx > -1) {
            const line = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
            try {
                const p = JSON.parse(line); if (!p || p.type === 'heartbeat') continue;
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'MONITORING'].includes(p.status)) {
                    if (p.status !== 'MONITORING') {
                        const cat = p.verdict.toLowerCase().includes('malicious') ? 'malicious' : (p.verdict.toLowerCase().includes('spam') ? 'spam' : 'safe');
                        const pIds = store.get('processedIds') || []; if (!pIds.includes(p.originalEntryId)) store.set('processedIds', [...pIds, p.originalEntryId].slice(-10000));
                        statsBuffer[cat].push(p);
                        if (!bufferTimer) bufferTimer = setTimeout(flushStats, 500);
                        if (p.fullHeaders || p.body) {
                            const fPath = path.join(FORENSICS_DIR, `${crypto.createHash('sha256').update(String(p.entryId)).digest('hex')}.json`);
                            fsPromises.writeFile(fPath, JSON.stringify({ fullHeaders: Buffer.from(p.fullHeaders || '', 'base64').toString(), body: Buffer.from(p.body || '', 'base64').toString() })).catch(() => {});
                        }
                    }
                }
                broadcastToUi({ type: 'scan-update', data: p });
            } catch {}
        }
    });
    currentScanChild.on('exit', () => isScanning = false);
}

function startService() {
    if (serviceSession.ownerPid) setInterval(() => { try { process.kill(serviceSession.ownerPid, 0); } catch { process.exit(0); } }, 5000);
    let buf = '';
    pipeServer = net.createServer(s => {
        let auth = false;
        s.on('data', d => {
            buf += d.toString(); let idx = buf.indexOf('\n');
            while (idx > -1) {
                const raw = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
                try {
                    const m = JSON.parse(raw); if (!m) continue;
                    if (!auth) { if (m.type === 'auth' && m.token === serviceSession.token) { auth = true; activeConnections.add(s); s.write(JSON.stringify({ type: 'status-sync', enabled: !!store.get('enabled'), stats: store.get('stats'), config: store.store }) + '\n'); } else s.destroy(); continue; }
                    if (m.type === 'store-get') s.write(JSON.stringify({ type: 'store-data', rid: m.rid, key: m.key, value: m.key === '' ? store.store : store.get(m.key) }) + '\n');
                    if (m.type === 'store-set') { store.set(m.key, m.value); if (m.key === 'enabled') { broadcastToUi({ type: 'status-sync', enabled: !!m.value }); if (m.value) runOutlookScanner(); else if (currentScanChild) currentScanChild.kill(); } }
                    if (m.type === 'cmd') { if (m.payload === 'Reset') { store.clear(); process.exit(0); } if (['Release', 'Quarantine'].includes(m.payload)) getPsWorker().stdin.write(JSON.stringify({ action: m.payload, ...m.data }) + '\n'); }
                } catch {}
            }
        });
        s.on('close', () => activeConnections.delete(s));
    }).listen(serviceSession.pipeName);
    if (store.get('enabled')) runOutlookScanner();
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
                if (r.type === 'store-data' && r.rid && reqHandlers.has(r.rid)) { const resolve = reqHandlers.get(r.rid); reqHandlers.delete(r.rid); resolve(r.value); }
                else if (r.type === 'status-sync' && r.stats) { broadcastToUi({ type: 'stats-update', data: { full: true, stats: r.stats } }); broadcastToUi(r); }
                else broadcastToUi(r);
            } catch {}
        }
    });
}

function spawnService() {
    if (isServiceMode || serviceSpawnInFlight) return;
    serviceSpawnInFlight = true;
    serviceSession = { pipeName: `\\\\.\\pipe\\mos_${process.pid}`, token: crypto.randomBytes(32).toString('hex'), ownerPid: process.pid };
    const env = { ...process.env, SVC_HANDSHAKE: JSON.stringify(serviceSession) };
    delete env.ELECTRON_RUN_AS_NODE;
    spawn(process.execPath, [APP_ROOT, '--service'], { detached: true, windowsHide: true, env });
    setTimeout(() => { uiPipeClient = net.connect(serviceSession.pipeName, () => { uiPipeClient.write(JSON.stringify({ type: 'auth', token: serviceSession.token }) + '\n'); setupPipeClient(); }); }, 2000);
}

app.on('ready', () => {
    Menu.setApplicationMenu(null);
    if (isServiceMode) { const h = JSON.parse(process.env.SVC_HANDSHAKE); if (h) { serviceSession = h; startService(); } }
    else {
        if (!app.requestSingleInstanceLock()) { app.quit(); return; }
        tray = new Tray(nativeImage.createFromPath(path.join(APP_ROOT, 'tray_off.png')));
        tray.on('click', () => { if(mainWindow) { if(mainWindow.isVisible()) mainWindow.hide(); else mainWindow.show(); } });
        updateTrayState();
        mainWindow = new BrowserWindow({ width: 1500, height: 900, backgroundColor: '#0a0e1c', show: false, webPreferences: { preload: path.join(APP_ROOT, 'preload.js'), contextIsolation: true, sandbox: true } });
        mainWindow.loadFile('index.html');
        mainWindow.on('close', e => { if (!isQuitting) { e.preventDefault(); mainWindow.hide(); } });
        mainWindow.once('ready-to-show', () => mainWindow.show());
        spawnService();
    }
});

function updateTrayState() {
    if (isServiceMode || !tray) return;
    const iconName = isEnabled ? 'tray_on.png' : 'tray_off.png';
    tray.setImage(nativeImage.createFromPath(path.join(APP_ROOT, iconName)));
    tray.setContextMenu(Menu.buildFromTemplate([{ label: 'Show Dashboard', click: () => mainWindow.show() }, { label: isEnabled ? '🛡️ Security: ACTIVE' : '⚠️ Security: DISABLED', enabled: false }, { label: isEnabled ? 'Stop Protection' : 'Start Protection', click: () => uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: !isEnabled }) + '\n') }, { type: 'separator' }, { label: 'Exit Application', click: () => { isQuitting = true; app.quit(); } }]));
}

const pipeReq = (m) => new Promise(resolve => { if (!uiPipeClient) return resolve(null); const rid = crypto.randomBytes(8).toString('hex'); reqHandlers.set(rid, resolve); uiPipeClient.write(JSON.stringify({ ...m, rid }) + '\n'); });

ipcMain.handle('get-config', () => pipeReq({ type: 'store-get', key: '' }).then(v => v || {}));
ipcMain.handle('get-stats', () => pipeReq({ type: 'store-get', key: 'stats' }).then(v => v || { malicious: [], suspicious: [], spam: [], safe: [] }));
ipcMain.handle('get-forensics', (e, id) => { const fPath = path.join(FORENSICS_DIR, `${crypto.createHash('sha256').update(String(id)).digest('hex')}.json`); return fs.existsSync(fPath) ? JSON.parse(fs.readFileSync(fPath, 'utf8')) : { fullHeaders: 'N/A', body: 'N/A' }; });
ipcMain.handle('set-enabled', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-vt-key', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'vtApiKey', value: safeStorage.encryptString(v).toString('base64') }) + '\n'); return { ok: true }; });
ipcMain.handle('set-spam-keywords', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'spamKeywords', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-rubrics', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'rubrics', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-whitelist', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'whitelist', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-blacklist', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'blacklist', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('release-email', (e, d) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Release', data: d }) + '\n'); return { ok: true }; });
ipcMain.handle('quarantine-email', (e, d) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Quarantine', data: d }) + '\n'); return { ok: true }; });
ipcMain.handle('open-logs-folder', () => shell.openPath(LOG_DIR));
ipcMain.handle('app-reset', () => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Reset' }) + '\n'); setTimeout(() => { app.relaunch(); app.exit(); }, 1000); });
