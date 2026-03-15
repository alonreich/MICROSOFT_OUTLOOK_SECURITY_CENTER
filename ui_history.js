
document.getElementById('history-toggle').onclick = async () => { 
    const cfg = await window.securityApi.getConfig(); 
    const newState = !cfg.historyScanEnabled;
    
    await window.securityApi.setHistoryEnabled(newState); 
    document.getElementById('history-toggle').classList.toggle('active', newState); 
    
    if (newState) {
        window.addLog("HISTORY SCAN MODE: Enabled. The engine will retrospectively audit ALL emails (Read & Unread).");
    } else {
        window.addLog("ON-ACCESS SCAN MODE: Enabled. The engine will now only audit new incoming unread emails.");
    }
    
    if (cfg.enabled) {
        window.addLog("Restarting security engine to apply scan mode changes...");
    }
};

(async () => {
    const cfg = await window.securityApi.getConfig();
    if (cfg.historyScanEnabled) {
        document.getElementById('history-toggle').classList.add('active');
    }
})();

