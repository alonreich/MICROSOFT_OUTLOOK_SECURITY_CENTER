const { contextBridge, ipcRenderer } = require('electron');

const invoke = (channel, payload) => ipcRenderer.invoke(channel, payload);

contextBridge.exposeInMainWorld('securityApi', {
    getStats: () => invoke('get-stats'),
    getConfig: () => invoke('get-config'),
    setEnabled: value => invoke('set-enabled', value),
    setHistoryEnabled: value => invoke('set-history-enabled', value),
    setVTKey: value => invoke('set-vt-key', value),
    setSpamKeywords: value => invoke('set-spam-keywords', value),
    setRubrics: value => invoke('set-rubrics', value),
    setWhitelist: value => invoke('set-whitelist', value),
    setBlacklist: value => invoke('set-blacklist', value),
    setSchedule: value => invoke('set-schedule', value),
    saveColumnWidths: value => invoke('save-column-widths', value),
    openLogsFolder: () => invoke('open-logs-folder'),
    resetApp: () => invoke('app-reset'),
    clearSecurityCache: () => invoke('clear-security-cache'),
    releaseEmail: value => invoke('release-email', value),
    checkPowerStatus: () => invoke('check-power-status'),
    overridePowerPlan: () => invoke('override-power-plan'),
    backupConfig: () => invoke('backup-config'),
    restoreConfig: () => invoke('restore-config'),
    cleanupListeners: () => {
        ipcRenderer.removeAllListeners('outlook-scan-update');
        ipcRenderer.removeAllListeners('status-sync');
        ipcRenderer.removeAllListeners('stats-update');
        ipcRenderer.removeAllListeners('live-log');
        ipcRenderer.removeAllListeners('email-released');
    },
    onOutlookScanUpdate: callback => ipcRenderer.on('outlook-scan-update', (event, data) => callback(data)),
    onStatusSync: callback => ipcRenderer.on('status-sync', (event, value) => callback(value)),
    onStatsUpdate: callback => ipcRenderer.on('stats-update', (event, data) => callback(data)),
    onLiveLog: callback => ipcRenderer.on('live-log', (event, message) => callback(message)),
    onEmailReleased: callback => ipcRenderer.on('email-released', (event, data) => callback(data)),
    quarantineEmail: (data) => invoke('quarantine-email', data)
});
