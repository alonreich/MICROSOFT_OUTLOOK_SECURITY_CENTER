const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('securityApi', {
  getStats: () => ipcRenderer.invoke('get-stats'),
  getConfig: () => ipcRenderer.invoke('get-config'),
  setEnabled: (val) => ipcRenderer.send('set-enabled', val),
  setHistoryEnabled: (val) => ipcRenderer.send('set-history-enabled', val),
  setVTKey: (key) => ipcRenderer.send('set-vt-key', key),
  setSpamKeywords: (keywords) => ipcRenderer.send('set-spam-keywords', keywords),
  setRubrics: (rubrics) => ipcRenderer.send('set-rubrics', rubrics),
  setWhitelist: (whitelist) => ipcRenderer.send('set-whitelist', whitelist),
  saveColumnWidths: (widths) => ipcRenderer.send('save-column-widths', widths),
  releaseEmail: (data) => ipcRenderer.send('release-email', data),
  checkPowerStatus: () => ipcRenderer.invoke('check-power-status'),
  overridePowerPlan: () => ipcRenderer.send('override-power-plan'),
  openLogsFolder: () => ipcRenderer.send('open-logs-folder'),
  resetApp: () => ipcRenderer.send('app-reset'),
  reportNetworkStatus: (status) => ipcRenderer.send('network-status', status),
  setSchedule: (val) => ipcRenderer.send('set-schedule', val),
  onEmailReleased: (callback) => {
    const listener = (_, id) => callback(id);
    ipcRenderer.on('email-released-success', listener);
    return () => ipcRenderer.removeListener('email-released-success', listener);
  },
  onEmailReleasedError: (callback) => {
    const listener = (_, data) => callback(data);
    ipcRenderer.on('email-released-error', listener);
    return () => ipcRenderer.removeListener('email-released-error', listener);
  },
  onOutlookScanUpdate: (callback) => {
    const listener = (_, data) => callback(data);
    ipcRenderer.on('outlook-scan-update', listener);
    return () => ipcRenderer.removeListener('outlook-scan-update', listener);
  },
  onStatsUpdate: (callback) => {
    const listener = (_, data) => callback(data);
    ipcRenderer.on('stats-update', listener);
    return () => ipcRenderer.removeListener('stats-update', listener);
  },
  onScanFinished: (callback) => {
    const listener = (_, mode) => callback(mode);
    ipcRenderer.on('scan-finished', listener);
    return () => ipcRenderer.removeListener('scan-finished', listener);
  },
  onStatusSync: (callback) => {
    const listener = (_, enabled) => callback(enabled);
    ipcRenderer.on('status-sync', listener);
    return () => ipcRenderer.removeListener('status-sync', listener);
  }
});
