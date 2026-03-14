const path = require('node:path');
const fs = require('node:fs');
const fsPromises = fs.promises;
const net = require('node:net');
const crypto = require('node:crypto');
const { spawn, execFile } = require('node:child_process');
const electron = require('electron');
const { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, safeStorage, shell, dialog } = electron;

const isServiceMode = process.argv.includes('--service');
const APP_ROOT = __dirname;
const LOG_DIR = path.join(APP_ROOT, 'logs');
const FORENSICS_DIR = path.join(LOG_DIR, 'forensics');
const LOG_FILE = path.join(LOG_DIR, 'security.log');

[LOG_DIR, FORENSICS_DIR].forEach(d => { if (!fs.existsSync(d)) try { fs.mkdirSync(d, { recursive: true }); } catch {} });

function logToFile(msg) {
    const ts = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
    try { 
        fs.appendFileSync(LOG_FILE, `[${ts}] ${msg}\n`); 
        broadcastToUi({ type: 'live-log', message: msg });
    } catch {}
}

const USER_DATA = app.getPath('userData');
const Store = require('electron-store');
const configStore = new Store({
    cwd: USER_DATA, name: 'config', clearInvalidConfig: true,
    defaults: { 
        enabled: false, 
        vtApiKey: '', 
        spamKeywords: ['viagra', 'lottery', 'urgent', 'bitcoin', 'sex', 'pussy', 'ass', 'סקס', 'תחת', 'כוס', 'זין', 'cock', 'dick', 'horny'], 
        rubrics: { 
            weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }, 
            toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }, 
            spamThresholdPercent: 50 
        }, 
        whitelist: { emails: [], ips: [], domains: [], combos: [] }, 
        blacklist: { emails: [], ips: [], domains: [], combos: [] },
        launchAtStartup: true
    }
});

const dataStore = new Store({
    cwd: USER_DATA, name: 'data', clearInvalidConfig: true,
    defaults: { processedIds: [], releasedFingerprints: [], stats: { spam: [], safe: [], malicious: [], suspicious: [] } }
});

let mainWindow = null, tray = null, isQuitting = false, isEnabled = !!configStore.get('enabled');
let uiPipeClient = null, serviceSession = null, serviceSpawnInFlight = false;
let pipeServer = null, activeConnections = new Set(), isScanning = false, currentScanChild = null;
let psWorker = null;

let configCache = {};
let statsCache = { malicious: [], suspicious: [], spam: [], safe: [] };

function updateCaches(data) {
    if (data.config) configCache = data.config;
    if (data.stats) statsCache = data.stats;
}


let watchdogTimer = null;
let lastHeartbeat = Date.now();


let statsBuffer = { malicious: [], suspicious: [], spam: [], safe: [] };
let bufferTimer = null;

function broadcastToUi(data) {
    if (isServiceMode) {
        const msg = JSON.stringify(data) + '\n';
        activeConnections.forEach(s => { try { s.write(msg); } catch { activeConnections.delete(s); } });
    } else {
        if (data.type === 'status-sync') { isEnabled = !!data.enabled; updateTrayState(); }
        if (mainWindow && !mainWindow.isDestroyed()) {
            if (data.type === 'scan-update') mainWindow.webContents.send('outlook-scan-update', data.data);
            if (data.type === 'status-sync') mainWindow.webContents.send('status-sync', data.enabled);
            if (data.type === 'stats-update') mainWindow.webContents.send('stats-update', data.data);
            if (data.type === 'outlook-status') mainWindow.webContents.send('outlook-status', data.running);
            if (data.type === 'live-log') mainWindow.webContents.send('live-log', data.message);
        }
    }
}

function flushStats() {
    if (!isServiceMode) return;
    const hasData = Object.values(statsBuffer).some(a => a.length > 0);
    if (hasData) {
        const currentStats = dataStore.get('stats') || { malicious: [], suspicious: [], spam: [], safe: [] };
        const newFingerprints = new Set();
        for (const cat in statsBuffer) {
            statsBuffer[cat].forEach(item => {
                const fid = item.fingerprint || item.entryId || item.originalEntryId;
                if (fid) newFingerprints.add(fid);
            });
        }
        for (const cat in currentStats) {
            currentStats[cat] = currentStats[cat].filter(item => {
                const fid = item.fingerprint || item.entryId || item.originalEntryId;
                return !newFingerprints.has(fid);
            });
        }
        for (const cat in statsBuffer) {
            const combined = [...(currentStats[cat] || []), ...statsBuffer[cat]];
            const seen = new Set();
            currentStats[cat] = combined.filter(item => {
                const fid = item.fingerprint || item.entryId || item.originalEntryId;
                if (!fid || seen.has(fid)) return false;
                seen.add(fid); return true;
            }).slice(-1000);
            statsBuffer[cat] = [];
        }
        dataStore.set('stats', currentStats);
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
            try { 
                const p = JSON.parse(line); 
                if (!p) continue;
                if (p.type === 'store-data') {
                    broadcastToUi(p); 
                    continue;
                }
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'ERROR'].includes(p.status)) {
                    broadcastToUi({ type: 'scan-update', data: p }); 
                }
            } catch {}
        }
    });
    return psWorker;
}

async function ensureOutlookRunning() {
    return new Promise(resolve => {
        execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe'], (err, stdout) => {
            if (stdout.toLowerCase().includes('outlook.exe')) {
                resolve(true);
            } else {
                logToFile('Outlook not detected. Attempting to start minimized...');
                
                spawn('cmd.exe', ['/c', 'start /min outlook.exe'], { detached: true, stdio: 'ignore' }).unref();
                
                
                let attempts = 0;
                const check = setInterval(() => {
                    execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe'], (err, stdout) => {
                        attempts++;
                        if (stdout.toLowerCase().includes('outlook.exe') || attempts > 10) {
                            clearInterval(check);
                            resolve(true);
                        }
                    });
                }, 1000);
            }
        });
    });
}

async function runOutlookScanner() {
    if (!isServiceMode) {
        logToFile('Attempted to run scanner in UI mode. Redirecting to service...');
        return;
    }
    if (!configStore.get('enabled') || isScanning) return;
    
    await ensureOutlookRunning();
    
    isScanning = true;
    lastHeartbeat = Date.now();
    if (watchdogTimer) clearInterval(watchdogTimer);
    watchdogTimer = setInterval(() => { 
        if (Date.now() - lastHeartbeat > 180000) { 
            logToFile('Watchdog: Scanner hung. Hard killing.'); 
            broadcastToUi({ type: 'outlook-status', running: false });
            if (currentScanChild) currentScanChild.kill('SIGKILL'); 
            isScanning = false; 
            runOutlookScanner(); 
        } 
    }, 10000);
    currentScanChild = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-ParentPid', process.pid.toString()], { windowsHide: true });
    
    const vtKeyEnc = configStore.get('vtApiKey');
    let vtKeyDec = '';
    if (vtKeyEnc) { try { vtKeyDec = safeStorage.decryptString(Buffer.from(vtKeyEnc, 'base64')); } catch {} }

    currentScanChild.stdin.write(JSON.stringify({ 
        mode: 'OnAccess', 
        processedIds: dataStore.get('processedIds'), 
        releasedFingerprints: dataStore.get('releasedFingerprints'),
        spamKeywords: configStore.get('spamKeywords'), 
        rubrics: configStore.get('rubrics'), 
        whitelist: configStore.get('whitelist'), 
        blacklist: configStore.get('blacklist'), 
        vtKey: vtKeyDec
    }) + '\n');
    let buf = '';
    currentScanChild.stdout.on('data', d => {
        buf += d.toString(); let idx = buf.indexOf('\n');
        while (idx > -1) {
            const line = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
            try {
                const p = JSON.parse(line); if (!p) continue; 
                if (p.type === 'heartbeat') { 
                    lastHeartbeat = Date.now(); 
                    broadcastToUi({ type: 'outlook-status', running: true });
                    continue; 
                }
                if (p.type === 'store-update') {
                    if (p.key === 'releasedFingerprints') {
                        const current = dataStore.get('releasedFingerprints') || [];
                        if (!current.includes(p.value)) dataStore.set('releasedFingerprints', [...current, p.value].slice(-5000));
                    }
                    continue;
                }
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'MONITORING', 'INFO', 'ERROR'].includes(p.status)) {
                    logToFile(`Scan Event [${p.status}]: ${p.details || ''} (Sender: ${p.sender || 'N/A'}, IP: ${p.ip || 'N/A'}, Score: ${p.score || 0}%, Tier: ${p.tier || 'N/A'})`);
                    if (p.status !== 'MONITORING' && p.status !== 'INFO' && p.status !== 'ERROR') {
                        const cat = p.verdict.toLowerCase().includes('malicious') ? 'malicious' : (p.verdict.toLowerCase().includes('spam') ? 'spam' : 'safe');
                        const pIds = dataStore.get('processedIds') || [];
                        const fid = p.fingerprint || p.originalEntryId;
                        if (!pIds.includes(fid)) dataStore.set('processedIds', [...pIds, fid].slice(-10000));
                        statsBuffer[cat].push(p);
                        if (!bufferTimer) bufferTimer = setTimeout(flushStats, 500);
                        if (p.fullHeaders || p.body) {
                            const fPath = path.join(FORENSICS_DIR, `${crypto.createHash('sha256').update(String(p.entryId)).digest('hex')}.json`);
                            fsPromises.writeFile(fPath, JSON.stringify({ fullHeaders: Buffer.from(p.fullHeaders || '', 'base64').toString(), body: Buffer.from(p.body || '', 'base64').toString() })).catch(() => {});
                        }
                    }
                    broadcastToUi({ type: 'scan-update', data: p });
                }
            } catch {}
        }
    });
    currentScanChild.on('exit', () => { isScanning = false; if(watchdogTimer) clearInterval(watchdogTimer); });
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
                    if (!auth) { if (m.type === 'auth' && m.token === serviceSession.token) { auth = true; activeConnections.add(s); s.write(JSON.stringify({ type: 'status-sync', enabled: !!configStore.get('enabled'), stats: dataStore.get('stats'), config: configStore.store }) + '\n'); } else s.destroy(); continue; }
                    if (m.type === 'store-get') {
                        let val;
                        if (m.key === '') val = configStore.store;
                        else if (['stats', 'processedIds'].includes(m.key)) val = dataStore.get(m.key);
                        else val = configStore.get(m.key);
                        s.write(JSON.stringify({ type: 'store-data', rid: m.rid, key: m.key, value: val }) + '\n');
                    }
                    if (m.type === 'store-set') { 
                        if (['stats', 'processedIds'].includes(m.key)) dataStore.set(m.key, m.value);
                        else configStore.set(m.key, m.value);
                        
                        if (m.key === 'enabled') { 
                            broadcastToUi({ type: 'status-sync', enabled: !!m.value, stats: dataStore.get('stats') }); 
                            if (m.value) { isScanning = false; runOutlookScanner(); }
                            else if (currentScanChild) { currentScanChild.kill('SIGKILL'); isScanning = false; }
                        } else if (['rubrics', 'spamKeywords', 'whitelist', 'blacklist', 'vtApiKey'].includes(m.key)) {
                            if (configStore.get('enabled')) {
                                logToFile(`Security policy updated (${m.key}). Restarting scanner...`);
                                if (currentScanChild) {
                                    currentScanChild.removeAllListeners('exit');
                                    currentScanChild.kill('SIGKILL');
                                }
                                isScanning = false;
                                runOutlookScanner();
                            }
                        }
                    }
                    if (m.type === 'cmd') { 
                        if (m.payload === 'Reset') { 
                            configStore.clear(); 
                            dataStore.clear();
                            try {
                                if (fs.existsSync(LOG_FILE)) fs.unlinkSync(LOG_FILE);
                                if (fs.existsSync(FORENSICS_DIR)) {
                                    const files = fs.readdirSync(FORENSICS_DIR);
                                    for (const f of files) {
                                        const p = path.join(FORENSICS_DIR, f);
                                        if (fs.statSync(p).isFile()) fs.unlinkSync(p);
                                    }
                                }
                            } catch {}
                            process.exit(0); 
                        } 
                        if (['Release', 'Quarantine', 'Delete', 'Check-Existence'].includes(m.payload)) {
                            getPsWorker().stdin.write(JSON.stringify({ action: m.payload, rid: m.rid, data: m.data }) + '\n'); 
                        }
                    }
                    } catch {}
                    }
                    });
                    s.on('close', () => activeConnections.delete(s));
                    }).listen(serviceSession.pipeName);    if (configStore.get('enabled')) runOutlookScanner();
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
                        configCache = r.value;
                        if (configCache.vtApiKey) { try { configCache.vtApiKey = safeStorage.decryptString(Buffer.from(configCache.vtApiKey, 'base64')); } catch { configCache.vtApiKey = ''; } }
                    }
                    if (r.key === 'stats') statsCache = r.value;
                    if (r.rid && reqHandlers.has(r.rid)) { const resolve = reqHandlers.get(r.rid); reqHandlers.delete(r.rid); resolve(r.value); }
                } else if (r.type === 'status-sync') {
                    if (r.stats) statsCache = r.stats;
                    if (r.config) {
                        configCache = r.config;
                        if (configCache.vtApiKey) { try { configCache.vtApiKey = safeStorage.decryptString(Buffer.from(configCache.vtApiKey, 'base64')); } catch { configCache.vtApiKey = ''; } }
                    }
                    broadcastToUi({ type: 'stats-update', data: { full: true, stats: statsCache } });
                    broadcastToUi(r);
                } else {
                    broadcastToUi(r);
                }
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
    if (isServiceMode) { 
        if (process.env.SVC_HANDSHAKE && process.env.SVC_HANDSHAKE !== 'undefined') {
            try {
                const h = JSON.parse(process.env.SVC_HANDSHAKE); 
                if (h) { serviceSession = h; startService(); }
            } catch { app.quit(); }
        } else {
            // Service started directly without handshake, usually by scheduler
            startService();
        }
    }
    else {
        if (!app.requestSingleInstanceLock()) { app.quit(); return; }
        const icon = nativeImage.createFromPath(path.join(APP_ROOT, 'tray_off.png')).resize({ width: 16, height: 16 });
        tray = new Tray(icon);
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
    const windowIconName = isEnabled ? 'icon_on.png' : 'icon_off.png';
    const icon = nativeImage.createFromPath(path.join(APP_ROOT, iconName)).resize({ width: 16, height: 16 });
    tray.setImage(icon);
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.setIcon(nativeImage.createFromPath(path.join(APP_ROOT, windowIconName)));
    }
    tray.setContextMenu(Menu.buildFromTemplate([
        { label: 'Show Dashboard', click: () => mainWindow.show() },
        { label: isEnabled ? 'Security: ACTIVE' : 'Security: DISABLED', enabled: false },
        { label: isEnabled ? 'Stop Protection' : 'Start Protection', click: () => uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: !isEnabled }) + '\n') },
        { type: 'separator' },
        { label: 'Exit Application', click: () => { isQuitting = true; app.quit(); } }
    ]));
}

const pipeReq = (m) => new Promise(resolve => { 
    if (!uiPipeClient) return resolve(null); 
    const rid = crypto.randomBytes(8).toString('hex'); 
    const timeout = setTimeout(() => {
        if (reqHandlers.has(rid)) {
            reqHandlers.delete(rid);
            logToFile(`IPC request timeout for ${m.type} ${m.key || m.payload || ''}`);
            resolve(null);
        }
    }, 5000);
    reqHandlers.set(rid, (val) => {
        clearTimeout(timeout);
        resolve(val);
    }); 
    uiPipeClient.write(JSON.stringify({ ...m, rid }) + '\n'); 
});

ipcMain.handle('get-config', () => {
    if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-get', key: '' }) + '\n'); 
    let res = (configCache && Object.keys(configCache).length > 0) ? { ...configCache } : { ...configStore.store };
    if (res.vtApiKey && !configCache.vtApiKey) { try { res.vtApiKey = safeStorage.decryptString(Buffer.from(res.vtApiKey, 'base64')); } catch { res.vtApiKey = ''; } }
    return res;
});
ipcMain.handle('get-stats', () => {
    if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-get', key: 'stats' }) + '\n');
    return statsCache || { malicious: [], suspicious: [], spam: [], safe: [] };
});
ipcMain.handle('get-forensics', (e, id) => { const fPath = path.join(FORENSICS_DIR, `${crypto.createHash('sha256').update(String(id)).digest('hex')}.json`); return fs.existsSync(fPath) ? JSON.parse(fs.readFileSync(fPath, 'utf8')) : { fullHeaders: 'N/A', body: 'N/A' }; });
ipcMain.handle('set-enabled', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-history-enabled', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'historyScanEnabled', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-vt-key', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'vtApiKey', value: safeStorage.encryptString(v).toString('base64') }) + '\n'); return { ok: true }; });
ipcMain.handle('set-spam-keywords', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'spamKeywords', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-rubrics', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'rubrics', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-whitelist', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'whitelist', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-blacklist', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'blacklist', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('save-column-widths', (e, v) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'columnWidths', value: v }) + '\n'); return { ok: true }; });
ipcMain.handle('set-startup', (e, v) => {
    if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'launchAtStartup', value: v }) + '\n');
    app.setLoginItemSettings({
        openAtLogin: v,
        path: process.execPath,
        args: [APP_ROOT, '--service']
    });
    return { ok: true };
});
ipcMain.handle('release-email', (e, d) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Release', data: d }) + '\n'); return { ok: true }; });
ipcMain.handle('quarantine-email', (e, d) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Quarantine', data: d }) + '\n'); return { ok: true }; });
ipcMain.handle('delete-email', (e, d) => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Delete', data: d }) + '\n'); return { ok: true }; });
ipcMain.handle('verify-existence', async (e, d) => {
    if (!uiPipeClient || !d.items || d.items.length === 0) return { removedCount: 0 };
    const rid = crypto.randomBytes(8).toString('hex');
    const probeCategory = d.items[0].category; // All items in a probe are same category
    
    return new Promise(resolve => {
        const timeout = setTimeout(() => {
            reqHandlers.delete(rid);
            resolve({ removedCount: 0 });
        }, 35000);
        
        reqHandlers.set(rid, (val) => {
            clearTimeout(timeout);
            if (val && val.removed && val.removed.length > 0) {
                const currentStats = dataStore.get('stats') || { malicious: [], suspicious: [], spam: [], safe: [] };
                const removedIds = new Set(val.removed.map(r => r.entryId));
                
                if (currentStats[probeCategory]) {
                    currentStats[probeCategory] = currentStats[probeCategory].filter(i => !removedIds.has(i.entryId));
                }
                
                dataStore.set('stats', currentStats);
                broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
                
                // If items were MOVED (found elsewhere), they will be picked up by the next scan cycle
                // because we didn't add their fingerprints to processedIds yet (or they are new IDs)
                
                resolve({ removedCount: val.removed.length });
            } else {
                resolve({ removedCount: 0 });
            }
        });
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Check-Existence', rid, data: d }) + '\n');
    });
});
ipcMain.handle('open-logs-folder', () => shell.openPath(LOG_DIR));
ipcMain.handle('app-reset', () => { if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Reset' }) + '\n'); setTimeout(() => { app.relaunch(); app.exit(); }, 1000); });

ipcMain.handle('export-config', async () => {
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Security Configuration',
        defaultPath: path.join(app.getPath('downloads'), 'outlook-security-config.json'),
        filters: [{ name: 'JSON Files', extensions: ['json'] }]
    });
    if (!filePath) return { canceled: true };
    const cfg = { ...configStore.store };
    if (cfg.vtApiKey) { try { cfg.vtApiKey = safeStorage.decryptString(Buffer.from(cfg.vtApiKey, 'base64')); } catch {} }
    
    const exportData = {
        vtApiKey: cfg.vtApiKey,
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
        title: 'Import Security Configuration',
        filters: [{ name: 'JSON Files', extensions: ['json'] }],
        properties: ['openFile']
    });
    if (!filePaths || filePaths.length === 0) return { canceled: true };
    try {
        const content = fs.readFileSync(filePaths[0], 'utf8');
        const data = JSON.parse(content);
        
        const keys = ['vtApiKey', 'spamKeywords', 'rubrics', 'whitelist', 'blacklist'];
        for (const k of keys) {
            if (data[k] !== undefined) {
                let val = data[k];
                if (k === 'vtApiKey' && val) { val = safeStorage.encryptString(val).toString('base64'); }
                configStore.set(k, val);
                if (uiPipeClient) uiPipeClient.write(JSON.stringify({ type: 'store-set', key: k, value: val }) + '\n');
            }
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
});

