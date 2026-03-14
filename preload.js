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
    saveColumnWidths: value => invoke('save-column-widths', value),
    setStartup: value => invoke('set-startup', value),
    exportConfig: () => invoke('export-config'),
    importConfig: () => invoke('import-config'),
    openLogsFolder: () => invoke('open-logs-folder'),
    resetApp: () => invoke('app-reset'),
    releaseEmail: value => invoke('release-email', value),
    cleanupListeners: () => {
        ipcRenderer.removeAllListeners('outlook-scan-update');
        ipcRenderer.removeAllListeners('status-sync');
        ipcRenderer.removeAllListeners('stats-update');
        ipcRenderer.removeAllListeners('live-log');
        ipcRenderer.removeAllListeners('email-moved');
        ipcRenderer.removeAllListeners('outlook-status');
    },
    onOutlookScanUpdate: callback => ipcRenderer.on('outlook-scan-update', (event, data) => callback(data)),
    onStatusSync: callback => ipcRenderer.on('status-sync', (event, value) => callback(value)),
    onStatsUpdate: callback => ipcRenderer.on('stats-update', (event, data) => callback(data)),
    onLiveLog: callback => ipcRenderer.on('live-log', (event, message) => callback(message)),
    onOutlookStatus: callback => ipcRenderer.on('outlook-status', (event, running) => callback(running)),
    onEmailMoved: callback => ipcRenderer.on('email-moved', (event, data) => callback(data)),
    quarantineEmail: (data) => invoke('quarantine-email', data),
    deleteEmail: (data) => invoke('delete-email', data),
    verifyExistence: (data) => invoke('verify-existence', data),
    getForensics: (id) => invoke('get-forensics', id)
});

