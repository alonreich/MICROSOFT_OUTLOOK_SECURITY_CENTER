const electron = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const { shell } = require('electron');
const { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, dialog } = electron;
app.commandLine.appendSwitch('disable-gpu-shader-disk-cache');
app.commandLine.appendSwitch('disable-http-cache');
app.commandLine.appendSwitch('disable-gpu-cache');
app.commandLine.appendSwitch('disk-cache-size', '1');
app.commandLine.appendSwitch('media-cache-size', '1');
const { execFile } = require('node:child_process');
const APP_ROOT = __dirname;
const LOG_DIR = path.join(APP_ROOT, 'logs');
const DEBUG_LOG_PATH = path.join(LOG_DIR, 'debug.log');
function logLine(section, details = '') {
  const now = new Date();
  const timestamp = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0') + ' ' + String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0') + ':' + String(now.getSeconds()).padStart(2, '0');
  const payload = `[${timestamp}] [${section}] ${details}\n`;
  try { 
    if (!fs.existsSync(LOG_DIR)) { fs.mkdirSync(LOG_DIR, { recursive: true }); }
    fs.appendFileSync(DEBUG_LOG_PATH, payload, { encoding: 'utf8', flag: 'a' }); 
  } catch (e) { console.error('Logging failed:', e); }
}
logLine('SESSION_START', 'MICROSOFT OUTLOOK SECURITY CENTER application launched.');
process.on('uncaughtException', (error) => { logLine('CRITICAL_EXCEPTION', error.stack || error.message); });
process.on('unhandledRejection', (reason) => { logLine('UNHANDLED_REJECTION', reason.stack || String(reason)); });
const electronStoreModule = require('electron-store');
const Store = electronStoreModule.default || electronStoreModule;
const store = new Store({
  projectName: 'microsoft-outlook-addon',
  name: 'security-data',
  defaults: {
    processedIds: [],
    stats: { spam: [], clean: [], malicious: [], suspicious: [] },
    enabled: true,
    historyScanEnabled: false,
    vtApiKey: "80a58ac4dbf037bebb6190a350160f451932a4a3cd56085c34e5b6483e058b98",
    spamKeywords: ["viagra", "lottery", "urgent", "inheritance", "winner", "prize", "verify your account", "bitcoin", "investment"],
    rubrics: { pointsSystem: true, bodyAnalysis: true, friendlySender: true, wordDecoding: true, phraseMatching: true, threshold: 5 },
    windowBounds: { width: 1500, height: 900 },
    whitelist: { emails: [], ips: [], domains: [], combos: [] },
    columnWidths: { subject: '2fr', date: '100px', time: '80px', ip: '120px', verdict: '100px', action: '150px', reasoning: '1fr' },
    schedule: { enabled: false, datetime: "" }
  },
});
let mainWindow = null;
let tray = null;
let isQuitting = false;
let isScanning = false;
let schedulerTimer = null;
let currentScanChild = null;
function startScheduler() {
    if (schedulerTimer) clearInterval(schedulerTimer);
    schedulerTimer = setInterval(() => {
        const sched = store.get('schedule');
        if (!sched || !sched.enabled) return;
        const now = new Date();
        const target = new Date(sched.datetime);
        if (now >= target && !isScanning) {
            logLine('SCHEDULED_SCAN', `Triggering one-time scheduled scan for: ${sched.datetime}`);
            store.set('schedule', { enabled: false, datetime: sched.datetime });
            runOutlookScanner('FullScan');
        }
    }, 30000);
}
function updateAppIcons(enabled) {
    const iconSuffix = enabled ? 'on' : 'off';
    const appIcon = path.join(APP_ROOT, `icon_${iconSuffix}.png`);
    const trayIcon = path.join(APP_ROOT, `tray_${iconSuffix}.png`);
    if (mainWindow) mainWindow.setIcon(nativeImage.createFromPath(appIcon));
    if (tray) tray.setImage(nativeImage.createFromPath(trayIcon));
}
function runOutlookScanner(mode = 'OnAccess') {
  if (!store.get('enabled') && mode === 'OnAccess') { logLine('SCAN_SKIP', 'Protection is disabled.'); return; }
  if (mode === 'FullScan' && !store.get('historyScanEnabled')) { logLine('SCAN_HALT', 'History scan disabled.'); return; }
  if (isScanning) { logLine('SCAN_SKIP', 'Scanner already running.'); return; }
  isScanning = true;
  logLine('SCAN_START', `Mode: ${mode}`);
  const psScript = path.join(APP_ROOT, 'outlook-scanner.ps1');
  const exchangeFile = path.join(app.getPath('temp'), `outlook_security_exchange_${Date.now()}.json`);
  const configExchange = { mode: mode, processedIds: store.get('processedIds'), vtApiKey: store.get('vtApiKey'), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist') };
  try { 
      fs.writeFileSync(exchangeFile, JSON.stringify(configExchange), 'utf8'); 
      logLine('EXCHANGE_CREATE', exchangeFile);
  } catch (e) { logLine('CRITICAL_ERROR', 'Exchange failed: ' + e.message); isScanning = false; return; }
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psScript, '-ExchangeFile', exchangeFile];
  const child = execFile('powershell.exe', args, { windowsHide: true }, (error, stdout, stderr) => {
    isScanning = false; currentScanChild = null;
    try { if (fs.existsSync(exchangeFile)) fs.unlinkSync(exchangeFile); } catch(e) {}
    if (error) logLine('SCAN_ERROR', error.message);
    logLine('SCAN_COMPLETE', mode);
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('scan-finished', mode);
  });
  currentScanChild = child;
  child.stdout.on('data', (data) => {
    data.trim().split('\n').forEach(line => {
      try {
        const result = JSON.parse(line);
        if (['Finished', 'THREAT BLOCKED', 'CAUTION', 'SPAM FILTERED'].includes(result.status)) {
          const currentIds = store.get('processedIds');
          if (!currentIds.includes(result.entryId)) {
            currentIds.push(result.entryId); store.set('processedIds', currentIds);
            const stats = store.get('stats'); const cat = result.verdict.toLowerCase();
            if (stats[cat]) { 
                stats[cat].push({ subject: result.details, date: result.timestamp, entryId: result.entryId, sender: result.sender, ip: result.ip, domain: result.domain, originalFolder: result.originalFolder, fullHeaders: result.fullHeaders, score: result.score, action: result.action, tier: result.tier }); 
                store.set('stats', stats); 
            }
            const reasonNote = result.tier ? ` | Reason: ${result.tier}` : '';
            const actionNote = result.action ? ` | Action: ${result.action}` : '';
            const ipNote = result.ip && result.ip !== 'N/A' ? ` | IP: ${result.ip}` : '';
            logLine('SCAN_RESULT', `[${result.verdict}] ${result.details}${reasonNote}${actionNote}${ipNote}`);
          }
        }
        if (mainWindow && !mainWindow.isDestroyed()) { mainWindow.webContents.send('outlook-scan-update', result); mainWindow.webContents.send('stats-update', store.get('stats')); }
      } catch (e) {}
    });
  });
}
function createTray() {
  const isEnabled = store.get('enabled');
  tray = new Tray(nativeImage.createFromPath(path.join(APP_ROOT, `tray_${isEnabled ? 'on' : 'off'}.png`)));
  const updateTrayMenu = () => {
    const currentEnabled = store.get('enabled');
    tray.setContextMenu(Menu.buildFromTemplate([
        { label: 'Show Security Hub', click: () => { if (mainWindow) mainWindow.show(); } },
        { label: currentEnabled ? 'DISABLE PROTECTION' : 'ENABLE MICROSOFT OUTLOOK PROTECTION', click: () => { 
            const newState = !currentEnabled; store.set('enabled', newState); 
            updateAppIcons(newState); 
            if (mainWindow) {
                mainWindow.webContents.send('status-sync', newState);
                mainWindow.webContents.send('outlook-scan-update', { status: newState ? "Enabled" : "Disabled", details: `Protection turned ${newState ? "ON" : "OFF"} from Tray.` });
            }
            updateTrayMenu(); 
        } },
        { type: 'separator' }, { label: 'Exit Application', click: () => { isQuitting = true; if(mainWindow) mainWindow.setClosable(true); app.quit(); } }
    ]));
  };
  tray.setToolTip('MICROSOFT OUTLOOK SECURITY CENTER');
  updateTrayMenu();
  tray.on('click', () => { if (mainWindow) mainWindow.show(); });
}
function createWindow() {
  const isEnabled = store.get('enabled');
  mainWindow = new BrowserWindow({ ...store.get('windowBounds'), backgroundColor: '#0a0e1c', icon: path.join(APP_ROOT, `icon_${isEnabled ? 'on' : 'off'}.png`), show: false, closable: true, webPreferences: { preload: path.join(APP_ROOT, 'preload.js'), nodeIntegration: false, contextIsolation: true } });
  mainWindow.loadFile(path.join(APP_ROOT, 'index.html'));
  mainWindow.setMenuBarVisibility(false);
  mainWindow.on('close', (event) => { if (!isQuitting) { event.preventDefault(); mainWindow.hide(); return false; } });
  mainWindow.on('minimize', (event) => { event.preventDefault(); mainWindow.hide(); });
  mainWindow.on('resize', () => { if(!mainWindow.isMaximized()) store.set('windowBounds', mainWindow.getBounds()); });
  mainWindow.on('move', () => { store.set('windowBounds', mainWindow.getBounds()); });
  mainWindow.once('ready-to-show', () => { 
      mainWindow.show(); 
      logLine('APP_READY', 'Dashboard loaded');
      if (store.get('historyScanEnabled')) runOutlookScanner('FullScan'); else runOutlookScanner('OnAccess'); 
  });
}
app.on('ready', () => { createTray(); createWindow(); startScheduler(); });
ipcMain.handle('get-stats', () => store.get('stats'));
ipcMain.handle('get-config', () => ({ enabled: store.get('enabled'), historyScanEnabled: store.get('historyScanEnabled'), vtApiKey: store.get('vtApiKey'), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist'), columnWidths: store.get('columnWidths'), schedule: store.get('schedule') }));
ipcMain.on('set-enabled', (event, val) => { logLine('USER_ACTION', `Protection set ${val ? 'ON' : 'OFF'}`); store.set('enabled', val); updateAppIcons(val); });
ipcMain.on('set-history-enabled', (event, val) => { 
    logLine('USER_ACTION', `History scan set ${val ? 'ON' : 'OFF'}`); 
    store.set('historyScanEnabled', val); 
    if(val) { runOutlookScanner('FullScan'); }
    else if (currentScanChild) { logLine('USER_ACTION', 'History scan halted by user.'); currentScanChild.kill(); isScanning = false; currentScanChild = null; }
});
ipcMain.on('set-vt-key', (event, key) => { logLine('USER_ACTION', 'VT Key updated'); store.set('vtApiKey', key); });
ipcMain.on('set-spam-keywords', (event, keywords) => { logLine('USER_ACTION', 'Spam keywords updated'); store.set('spamKeywords', keywords); });
ipcMain.on('set-rubrics', (event, rubrics) => { logLine('USER_ACTION', 'Spam rubrics updated'); store.set('rubrics', rubrics); });
ipcMain.on('set-whitelist', (event, whitelist) => { logLine('USER_ACTION', 'Whitelist updated'); store.set('whitelist', whitelist); });
ipcMain.on('set-schedule', (event, val) => { logLine('USER_ACTION', `Scan schedule updated: ${val.enabled ? 'ON' : 'OFF'} at ${val.datetime}`); store.set('schedule', val); startScheduler(); });
ipcMain.on('save-column-widths', (event, widths) => { store.set('columnWidths', widths); });
ipcMain.on('open-logs-folder', () => { shell.openPath(LOG_DIR); });
ipcMain.on('app-reset', () => { logLine('USER_ACTION', 'App reset initiated'); store.clear(); app.relaunch(); app.exit(); });
ipcMain.handle('backup-config', async () => {
    const config = { vtApiKey: store.get('vtApiKey'), spamKeywords: store.get('spamKeywords'), rubrics: store.get('rubrics'), whitelist: store.get('whitelist'), schedule: store.get('schedule'), processedIds: store.get('processedIds') };
    const { filePath } = await dialog.showSaveDialog(mainWindow, { title: 'Export Security Configuration', defaultPath: path.join(app.getPath('downloads'), `outlook_security_backup_${Date.now()}.json`), filters: [{ name: 'JSON Files', extensions: ['json'] }] });
    if (filePath) { fs.writeFileSync(filePath, JSON.stringify(config, null, 2), 'utf8'); logLine('USER_ACTION', `Configuration backed up to: ${filePath}`); return true; }
    return false;
});
ipcMain.handle('restore-config', async () => {
    const { filePaths } = await dialog.showOpenDialog(mainWindow, { title: 'Import Security Configuration', filters: [{ name: 'JSON Files', extensions: ['json'] }], properties: ['openFile'] });
    if (filePaths && filePaths[0]) {
        try {
            const data = JSON.parse(fs.readFileSync(filePaths[0], 'utf8'));
            if (data.vtApiKey !== undefined) store.set('vtApiKey', data.vtApiKey);
            if (data.spamKeywords) store.set('spamKeywords', data.spamKeywords);
            if (data.rubrics) store.set('rubrics', data.rubrics);
            if (data.whitelist) store.set('whitelist', data.whitelist);
            if (data.schedule) store.set('schedule', data.schedule);
            if (data.processedIds) store.set('processedIds', data.processedIds);
            logLine('USER_ACTION', `Configuration restored from: ${filePaths[0]}`);
            app.relaunch(); app.exit();
            return true;
        } catch (e) { logLine('CRITICAL_ERROR', `Restore failed: ${e.message}`); return false; }
    }
    return false;
});
ipcMain.on('release-email', (event, { entryId, whitelistEntry, originalFolder }) => {
    logLine('USER_ACTION', `Attempting confirmed release for ID: ${entryId.substring(0,10)}...`);
    execFile('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(APP_ROOT, 'outlook-scanner.ps1'), '-Mode', 'Release', '-TargetEntryId', entryId, '-OriginalFolder', originalFolder || ""], { windowsHide: true }, (error, stdout) => {
        let success = !error; let errorMessage = "";
        if (stdout) { try { const res = JSON.parse(stdout); if (res.status === 'Error') { success = false; errorMessage = res.message; } else { logLine('APP_EVENT', `Physical Move Verified: ${res.message}`); } } catch(e) {} }
        if (success) {
            if (whitelistEntry) {
                const wl = store.get('whitelist'); const stats = store.get('stats'); let addedValue = "";
                if (whitelistEntry.type === 'email') { if (!wl.emails.includes(whitelistEntry.value)) { wl.emails.push(whitelistEntry.value); addedValue = whitelistEntry.value; } }
                else if (whitelistEntry.type === 'ip') { if (!wl.ips.includes(whitelistEntry.value)) { wl.ips.push(whitelistEntry.value); addedValue = whitelistEntry.value; } }
                else if (whitelistEntry.type === 'domain') { if (!wl.domains.includes(whitelistEntry.value)) { wl.domains.push(whitelistEntry.value); addedValue = whitelistEntry.value; } }
                else if (whitelistEntry.type === 'combo') { const exists = wl.combos.some(c => c.ip === whitelistEntry.ip && c.domain === whitelistEntry.domain); if (!exists) { wl.combos.push({ ip: whitelistEntry.ip, domain: whitelistEntry.domain }); addedValue = `${whitelistEntry.domain}:${whitelistEntry.ip}`; } }
                if (addedValue) { store.set('whitelist', wl); logLine('APP_EVENT', `Whitelist Updated: ${whitelistEntry.type.toUpperCase()} -> ${addedValue}`); }
                const itemsToMove = [];
                ['malicious', 'suspicious', 'spam'].forEach(cat => {
                    const remaining = [];
                    stats[cat].forEach(item => {
                        let match = false;
                        if (whitelistEntry.type === 'email' && item.sender === whitelistEntry.value) match = true;
                        else if (whitelistEntry.type === 'ip' && item.ip === whitelistEntry.value) match = true;
                        else if (whitelistEntry.type === 'domain' && item.domain === whitelistEntry.value) match = true;
                        else if (whitelistEntry.type === 'combo' && item.ip === whitelistEntry.ip && item.domain === whitelistEntry.domain) match = true;
                        else if (item.entryId === entryId) match = true;
                        if (match) { item.action = "Restored & Whitelisted"; item.tier = `User Action: ${whitelistEntry.type.toUpperCase()}`; itemsToMove.push(item); }
                        else { remaining.push(item); }
                    });
                    stats[cat] = remaining;
                });
                itemsToMove.forEach(item => { if (!stats.clean.some(c => c.entryId === item.entryId)) stats.clean.push(item); });
                store.set('stats', stats); if (mainWindow) mainWindow.webContents.send('stats-update', stats);
            }
            logLine('APP_EVENT', `Release transaction completed: ${entryId.substring(0,10)}...`);
            event.reply('email-released-success', entryId);
        } else {
            logLine('RELEASE_ERROR', `Physical Restore Blocked: ${errorMessage || error.message}`);
            event.reply('email-released-error', { entryId, message: errorMessage || "Outlook blocked the move operation. Is the item open in another window?" });
        }
    });
});
ipcMain.handle('check-power-status', async () => { return new Promise((resolve) => { execFile('powercfg', ['/query', 'SCHEME_CURRENT', 'SUB_SLEEP', 'STANDBYIDLE'], { windowsHide: true }, (error, stdout) => { if (error) return resolve({ safe: true }); resolve({ safe: !stdout.includes('Current AC Power Setting Index: 0x00000000') === false }); }); }); });
ipcMain.on('override-power-plan', () => { logLine('USER_ACTION', 'Override power plan'); execFile('powercfg', ['/change', 'standby-timeout-ac', '0'], { windowsHide: true }); execFile('powercfg', ['/change', 'hibernate-timeout-ac', '0'], { windowsHide: true }); });
