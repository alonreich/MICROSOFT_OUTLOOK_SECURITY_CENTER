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
const LOG_FILE = path.join(LOG_DIR, 'microsoft_outlook_security.log');

[LOG_DIR, FORENSICS_DIR].forEach(d => { if (!fs.existsSync(d)) try { fs.mkdirSync(d, { recursive: true }); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } } });
if (!fs.existsSync(LOG_FILE)) try { fs.writeFileSync(LOG_FILE, `[${new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '')}] Security Center: Initialization Success. Monitoring is active.\n`); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }

function logToFile(msg, level = "INFO") {
    const ts = new Date().toISOString().replace(/T/, ' ').replace(/\..+/, '');
    const logLine = `[${ts}] [${level}] ${msg}\n`;
    try { 
        if (fs.existsSync(LOG_FILE)) {
            const stats = fs.statSync(LOG_FILE);
            if (stats.size > 5 * 1024 * 1024) {
                fs.renameSync(LOG_FILE, LOG_FILE + '.1');
            }
        }
        fs.appendFileSync(LOG_FILE, logLine); 
        broadcastToUi({ type: 'live-log', message: `[${level}] ${msg}` });
    } catch(err) { console.error(err); }
}

const USER_DATA = app.getPath('userData');
const Store = require('electron-store');
const configStore = new Store({
    cwd: USER_DATA, name: 'config', clearInvalidConfig: true,
    defaults: { 
        enabled: false, 
        vtApiKey: '', 
        spamKeywords: ['viagra', 'lottery', 'urgent', 'bitcoin', 'winner', 'unpaid', 'invoice', 'payment', 'account', 'verify', 'security', 'update', 'action'], 
        rubrics: { 
            weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }, 
            toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }, 
            spamThresholdPercent: 50 
        }, 
        whitelist: { emails: [], ips: [], domains: [], combos: [] }, 
        blacklist: { emails: [], ips: [], domains: [], combos: [] },
        launchAtStartup: true,
        scanningSpeed: 50,
        historyScanEnabled: false
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
let statsBuffer = { malicious: [], suspicious: [], spam: [], safe: [] };
let bufferTimer = null;
let watchdogTimer = null;
let lastHeartbeat = Date.now();
let configCache = {};
let statsCache = null;

function broadcastToUi(msg) {
    if (isServiceMode) {
        const raw = JSON.stringify(msg) + "\n";
        activeConnections.forEach(s => { try { s.write(raw); } catch(e){} });
    } else if (mainWindow && mainWindow.webContents) {
        if (msg.type === 'scan-update') {
            mainWindow.webContents.send("outlook-scan-update", msg.data);
            if (msg.data && msg.data.status && ['THREAT BLOCKED', 'SPAM FILTERED', 'Finished'].includes(msg.data.status)) {
                const engineDetails = msg.data.tier ? ` [Engines: ${msg.data.tier}]` : "";
                logToFile(`Scan Result [${msg.data.status}]: ${msg.data.details}${engineDetails} (${msg.data.sender || 'N/A'})`);
            }
        }
        else if (msg.type === 'status-sync') mainWindow.webContents.send("status-sync", msg.enabled);
        else if (msg.type === 'stats-update') mainWindow.webContents.send("stats-update", msg.data);
        else if (msg.type === 'live-log') mainWindow.webContents.send("live-log", msg.message);
        else if (msg.type === 'outlook-status') mainWindow.webContents.send("outlook-status", msg.running);
        else if (msg.type === 'duplicate-update') mainWindow.webContents.send("duplicate-update", msg);
        else mainWindow.webContents.send("from-main", msg);
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
            }).slice(-1000);
            statsBuffer[cat] = [];
        }
        dataStore.set('stats', currentStats);
        broadcastToUi({ type: 'stats-update', data: { full: true, stats: currentStats } });
    }
    bufferTimer = null;
}

// Forensic Cleanup: Remove files that are not in the current stats or buffer
async function cleanupForensics() {
    try {
        const stats = dataStore.get('stats') || { malicious: [], suspicious: [], spam: [], safe: [] };
        const activeFingerprints = new Set();
        const cats = ['malicious', 'suspicious', 'spam', 'safe'];
        cats.forEach(cat => {
            (stats[cat] || []).concat(statsBuffer[cat] || []).forEach(item => {
                const fid = item.fingerprint || item.entryId;
                if (fid) activeFingerprints.add(crypto.createHash('sha256').update(String(fid)).digest('hex'));
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
    if (psWorker && !psWorker.killed) return psWorker;
    logToFile('Spawning Security Engine Worker process...');
    psWorker = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-Mode', 'Worker', '-ParentPid', process.pid.toString()], { windowsHide: true });
    let buf = '';
    psWorker.stdout.on('data', d => {
        buf += d.toString(); 
        let idx = buf.indexOf('\n');
        while (idx > -1) {
            const line = buf.slice(0, idx).trim(); 
            buf = buf.slice(idx + 1); 
            idx = buf.indexOf('\n');
            if (!line) continue;
            try { 
                const p = JSON.parse(line); 
                if (!p) continue;
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
                    continue;
                }
                if (['Finished', 'THREAT BLOCKED', 'SPAM FILTERED', 'MONITORING', 'INFO', 'ERROR'].includes(p.status)) {
                    if (p.status === 'INFO' || p.status === 'MONITORING') {
                        logToFile(`Worker: ${p.details || ''}`);
                    } else if (p.status === 'ERROR') {
                        logToFile(`Worker Error: ${p.details || ''}`, 'ERROR');
                    } else {
                        logToFile(`Worker Result [${p.status}]: ${p.details || ''}`);
                    }
                    broadcastToUi({ type: 'scan-update', data: p }); 
                }
            } catch (err) { 
                // Silent catch for parsing noise
            }
        }
    });
    psWorker.on('exit', (code) => {
        logToFile(`Security Engine Worker process exited with code ${code}`);
    });
    return psWorker;
}

async function ensureOutlookRunning() {
    return new Promise(resolve => {
        execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe'], (err, stdout) => {
            if (stdout.toLowerCase().includes('outlook.exe')) {
                resolve(true);
            } else {
                logToFile('Outlook not detected. Probing Registry for installation path...');
                let outlookPath = 'outlook.exe';
                try {
                    const regPath = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\outlook.exe';
                    const cp = require('child_process');
                    const regOut = cp.execSync(`reg query "${regPath}" /ve`).toString();
                    const match = regOut.match(/REG_SZ\s+(.+)/);
                    if (match && match[1]) { outlookPath = match[1].trim(); }
                } catch (e) {
                    logToFile('Registry probe failed. Falling back to default PATH search.', 'WARN');
                }

                logToFile(`Launching Outlook: ${outlookPath}`);
                spawn(outlookPath, ['/min'], { detached: true, stdio: 'ignore' }).unref();
                
                let attempts = 0;
                const check = setInterval(() => {
                    execFile('tasklist', ['/FI', 'IMAGENAME eq outlook.exe'], (err, stdout) => {
                        attempts++;
                        if (stdout.toLowerCase().includes('outlook.exe')) {
                            clearInterval(check);
                            logToFile('Outlook successfully launched and detected.');
                            resolve(true);
                        } else if (attempts > 30) {
                            clearInterval(check);
                            logToFile('Outlook launch timeout reached.', 'ERROR');
                            resolve(false);
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
    if (!configStore.get('enabled')) return;
    if (isScanning) {
        logToFile('Scanner already running. Skipping duplicate start.');
        return;
    }
    
    await ensureOutlookRunning();
    await cleanupForensics();

    isScanning = true;
    lastHeartbeat = Date.now();
    if (watchdogTimer) clearInterval(watchdogTimer);
    watchdogTimer = setInterval(() => { 
        if (Date.now() - lastHeartbeat > 180000) { 
            logToFile('Watchdog: Scanner hung. Hard killing.', 'ERROR'); 
            broadcastToUi({ type: 'outlook-status', running: false });
            if (currentScanChild) {
                currentScanChild.removeAllListeners('exit');
                currentScanChild.kill('SIGKILL'); 
            }
            isScanning = false; 
            runOutlookScanner(); 
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
        buf += d.toString(); 
        let idx = buf.indexOf('\n');
        while (idx > -1) {
            const line = buf.slice(0, idx).trim(); 
            buf = buf.slice(idx + 1); 
            idx = buf.indexOf('\n');
            if (!line) continue;
            try {
                const p = JSON.parse(line); 
                if (!p) continue; 
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
                    if (p.status === 'INFO' || p.status === 'MONITORING') {
                        logToFile(`Engine: ${p.details || ''}`);
                    } else if (p.status === 'ERROR') {
                        logToFile(`Engine Error: ${p.details || ''}`, 'ERROR');
                    } else {
                        logToFile(`Scan Result [${p.status}]: ${p.details || ''} (${p.sender || 'N/A'})`);
                    }
                    
                    if (p.status !== 'MONITORING' && p.status !== 'INFO' && p.status !== 'ERROR') {
                        const cat = p.verdict.toLowerCase().includes('malicious') ? 'malicious' : (p.verdict.toLowerCase().includes('spam') ? 'spam' : 'safe');
                        
                        // Use fingerprint as the unique key for tracking what was already scanned
                        const fid = p.fingerprint || p.entryId || p.originalEntryId;
                        if (fid) {
                            const pIds = dataStore.get('processedIds') || [];
                            if (!pIds.includes(fid)) {
                                dataStore.set('processedIds', [...pIds, fid].slice(-10000));
                            }
                        }

                        statsBuffer[cat].push(p);
                        if (!bufferTimer) bufferTimer = setTimeout(flushStats, 500);
                        if (p.fullHeaders || p.body) {
                            const forensicId = p.entryId || p.originalEntryId || p.fingerprint;
                            const fHash = crypto.createHash('sha256').update(String(forensicId)).digest('hex');
                            const fPath = path.join(FORENSICS_DIR, `${fHash}.json`);
                            fsPromises.writeFile(fPath, JSON.stringify({ fullHeaders: Buffer.from(p.fullHeaders || '', 'base64').toString(), body: Buffer.from(p.body || '', 'base64').toString() })).catch(() => {});
                        }
                    }
                    broadcastToUi({ type: 'scan-update', data: p });
                }
            } catch (err) { 
                // Silent catch for line parsing noise
            }
        }
    });

    currentScanChild.on('exit', (code) => { 
        logToFile(`Security Engine process exited with code ${code}`);
        isScanning = false; 
        if(watchdogTimer) clearInterval(watchdogTimer); 
    });
}

let restartTimer = null;
function requestScannerRestart(reason) {
    if (restartTimer) clearTimeout(restartTimer);
    
    // Optimization: If the engine is already running and we're just updating policies,
    // push the data live instead of a hard restart.
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
                vtKey: vtKeyDec
            }) + '\n');
            return;
        } catch (e) {
            logToFile(`Live Injection Failed: ${e.message}. Falling back to hard restart.`, 'WARN');
        }
    }

    restartTimer = setTimeout(() => {
        if (!configStore.get('enabled')) return;
        logToFile(`Hard Engine Restart triggered [Reason: ${reason}]`);
        if (currentScanChild) {
            currentScanChild.removeAllListeners('exit');
            currentScanChild.kill('SIGKILL');
        }
        isScanning = false;
        runOutlookScanner();
    }, 1000);
}

function startService() {
    if (serviceSession && serviceSession.ownerPid) setInterval(() => { try { process.kill(serviceSession.ownerPid, 0); } catch { process.exit(0); } }, 5000);
    let buf = '';
    pipeServer = net.createServer(s => {
        let auth = false;
        s.on('data', d => {
            buf += d.toString(); let idx = buf.indexOf('\n');
            while (idx > -1) {
                const raw = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1); idx = buf.indexOf('\n');
                try {
                    const m = JSON.parse(raw); if (!m) continue;
                    if (!auth) { if (serviceSession && m.type === 'auth' && m.token === serviceSession.token) { auth = true; activeConnections.add(s); s.write(JSON.stringify({ type: 'status-sync', enabled: !!configStore.get('enabled'), stats: dataStore.get('stats'), config: configStore.store }) + '\n'); } else s.destroy(); continue; }
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
                        
                        if (m.key === 'enabled' || m.key === 'historyScanEnabled') { 
                            broadcastToUi({ type: 'status-sync', enabled: !!configStore.get('enabled'), stats: dataStore.get('stats') }); 
                            if (configStore.get('enabled')) { 
                                requestScannerRestart(m.key);
                            }
                            else if (currentScanChild) { 
                                currentScanChild.removeAllListeners('exit');
                                currentScanChild.kill('SIGKILL'); 
                                isScanning = false; 
                            }
                        } else if (m.key === 'scanningSpeed') {
                            if (currentScanChild) {
                                currentScanChild.stdin.write(JSON.stringify({ type: 'config-update', scanningSpeed: m.value }) + '\n');
                            }
                        } else if (['rubrics', 'spamKeywords', 'whitelist', 'blacklist', 'vtApiKey'].includes(m.key)) {
                            requestScannerRestart(m.key);
                        }
                    }
                    if (m.type === 'cmd') { 
                        if (m.payload === 'Reset') { 
                            try {
                                if (currentScanChild) currentScanChild.kill('SIGKILL');
                                if (psWorker) psWorker.kill('SIGKILL');
                            } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }
                            
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
                            } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }
                            process.exit(0); 
                        } 
                        if (m.payload === 'Release' || m.payload === 'Quarantine' || m.payload === 'Delete' || m.payload === 'Check-Existence' || m.payload === 'DuplicateScan') {
                            logToFile(`DEBUG: Service received [${m.payload}] command. Forwarding to Security Worker...`);
                            const worker = getPsWorker();
                            if (worker && worker.stdin) {
                                worker.stdin.write(JSON.stringify({ action: m.payload, rid: m.rid, ...m.data }) + '\n'); 
                            } else {
                                logToFile(`ERROR: Security Worker stdin is not available. Command dropped.`, 'ERROR');
                            }
                        }
                    }
                    } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }
                    }
                    });
                    s.on('close', () => activeConnections.delete(s));
                    }).listen(serviceSession.pipeName, () => {
                        logToFile('Security Service IPC Layer: ACTIVE');
                        if (configStore.get('enabled')) runOutlookScanner();
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
                        configCache = r.value;
                        if (configCache.vtApiKey) { try { configCache.vtApiKey = safeStorage.decryptString(Buffer.from(configCache.vtApiKey, 'base64')); } catch { configCache.vtApiKey = ''; } }
                    }
                    if (r.key === 'stats') statsCache = r.value;
                    if (r.rid && reqHandlers.has(r.rid)) { const resolve = reqHandlers.get(r.rid); reqHandlers.delete(r.rid); resolve(r.value); }
                } else if (r.type === 'stats-update') {
                    if (r.data && r.data.full) statsCache = r.data.stats;
                    broadcastToUi(r);
                } else if (r.type === 'status-sync') {
                    if (r.stats) statsCache = r.stats;
                    if (r.config) {
                        configCache = r.config;
                        if (configCache.vtApiKey) { try { configCache.vtApiKey = safeStorage.decryptString(Buffer.from(configCache.vtApiKey, 'base64')); } catch { configCache.vtApiKey = ''; } }
                    }
                    isEnabled = r.enabled;
                    updateTrayState();
                    broadcastToUi({ type: 'stats-update', data: { full: true, stats: statsCache } });
                    broadcastToUi(r);
                } else {
                    broadcastToUi(r);
                }
            } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }
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
    
    const shouldStart = configStore.get('launchAtStartup');
    if (shouldStart !== undefined) {
        app.setLoginItemSettings({
            openAtLogin: shouldStart,
            path: process.execPath,
            args: [APP_ROOT, '--service']
        });
    }

    // Auto-migration: Ensure new default keywords are added if they don't exist
    const currentKeywords = configStore.get('spamKeywords') || [];
    const defaultKeywords = ['viagra', 'lottery', 'urgent', 'bitcoin', 'sex', 'pussy', 'ass', 'סקס', 'תחת', 'כוס', 'זין', 'cock', 'dick', 'horny'];
    let keywordsChanged = false;
    defaultKeywords.forEach(kw => {
        if (!currentKeywords.includes(kw)) {
            currentKeywords.push(kw);
            keywordsChanged = true;
        }
    });
    if (keywordsChanged) configStore.set('spamKeywords', currentKeywords);

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

ipcMain.handle('get-config', async () => {
    if (uiPipeClient) {
        const res = await pipeReq({ type: 'store-get', key: '' });
        if (res) {
            configCache = res;
            let cfg = { ...configCache };
            if (cfg.vtApiKey) { try { cfg.vtApiKey = safeStorage.decryptString(Buffer.from(cfg.vtApiKey, 'base64')); } catch { cfg.vtApiKey = ''; } }
            return cfg;
        }
    }
    let res = (configCache && Object.keys(configCache).length > 0) ? { ...configCache } : { ...configStore.store };
    if (res.vtApiKey && !configCache.vtApiKey) { try { res.vtApiKey = safeStorage.decryptString(Buffer.from(res.vtApiKey, 'base64')); } catch { res.vtApiKey = ''; } }
    return res;
});
ipcMain.handle('pause-duplicate-scan', () => {
    try { fs.writeFileSync(path.join(APP_ROOT, '.dup_pause'), '1'); return { ok: true }; }
    catch (e) { return { ok: false, error: e.message }; }
});

ipcMain.handle('resume-duplicate-scan', () => {
    const p = path.join(APP_ROOT, '.dup_pause');
    try { if (fs.existsSync(p)) fs.unlinkSync(p); return { ok: true }; }
    catch (e) { return { ok: false, error: e.message }; }
});

ipcMain.handle('get-stats', async () => {
    if (uiPipeClient) {
        const res = await pipeReq({ type: 'store-get', key: 'stats' });
        if (res) {
            statsCache = res;
            return statsCache;
        }
    }
    return statsCache || { malicious: [], suspicious: [], spam: [], safe: [] };
});
ipcMain.handle('get-forensics', (e, id) => { 
    // Always hash the ID provided by the UI (which is entryId)
    const fHash = crypto.createHash('sha256').update(String(id)).digest('hex');
    const fPath = path.join(FORENSICS_DIR, `${fHash}.json`); 
    if (fs.existsSync(fPath)) {
        return JSON.parse(fs.readFileSync(fPath, 'utf8'));
    }
    return { fullHeaders: 'N/A', body: 'N/A' }; 
});
ipcMain.handle('set-enabled', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'enabled', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-history-enabled', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'historyScanEnabled', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-vt-key', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'vtApiKey', value: safeStorage.encryptString(v).toString('base64') }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-spam-keywords', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'spamKeywords', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-rubrics', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'rubrics', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-whitelist', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'whitelist', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-blacklist', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'blacklist', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('save-column-widths', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'columnWidths', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-scanning-speed', (e, v) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'scanningSpeed', value: v }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('set-startup', (e, v) => {
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'store-set', key: 'launchAtStartup', value: v }) + '\n');
        app.setLoginItemSettings({ openAtLogin: v, path: process.execPath, args: [APP_ROOT, '--service'] });
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});
ipcMain.handle('release-email', (e, d) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Release', data: d }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('quarantine-email', (e, d) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Quarantine', data: d }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
ipcMain.handle('delete-email', (e, d) => { if (uiPipeClient) { uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Delete', data: d }) + '\n'); return { ok: true }; } return { ok: false, error: 'Service initializing' }; });
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

ipcMain.handle('scan-duplicates', async () => {
    logToFile('UI REQUEST: Initialize Duplicate Email Detection sequence.');
    if (uiPipeClient) {
        if (!psWorker || psWorker.killed) {
            logToFile('RECOVERY: Security Worker found in inactive state. Re-spawning for Duplicate Scan...', 'WARN');
            getPsWorker();
        }
        
        logToFile('COMMAND: Sending [DuplicateScan] action to Security Worker.');
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'DuplicateScan' }) + '\n');
        return { ok: true };
    }
    logToFile('ERROR: IPC Pipe not established. Duplicate Scan request rejected.', 'ERROR');
    return { ok: false, error: 'Service initializing' };
});

ipcMain.handle('delete-duplicates', async (e, d) => {
    logToFile(`UI Request: Deleting ${d.entryIds?.length || 0} duplicate emails...`);
    if (uiPipeClient) {
        uiPipeClient.write(JSON.stringify({ type: 'cmd', payload: 'Delete', data: d }) + '\n');
        return { ok: true };
    }
    return { ok: false, error: 'Service initializing' };
});

ipcMain.handle('export-config', async () => {
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Security Configuration',
        defaultPath: path.join(app.getPath('downloads'), 'outlook-security-config.json'),
        filters: [{ name: 'JSON Files', extensions: ['json'] }]
    });
    if (!filePath) return { canceled: true };
    const cfg = { ...configStore.store };
    if (cfg.vtApiKey) { try { cfg.vtApiKey = safeStorage.decryptString(Buffer.from(cfg.vtApiKey, 'base64')); } catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } } }
    
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

