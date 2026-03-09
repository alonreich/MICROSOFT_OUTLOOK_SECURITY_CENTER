
document.getElementById('history-toggle').onclick = async () => { 
    const cfg = await window.securityApi.getConfig(); 
    
    
    if (window.securityApi.setHistoryEnabled) {
        await window.securityApi.setHistoryEnabled(!cfg.historyScanEnabled); 
        document.getElementById('history-toggle').classList.toggle('active', !cfg.historyScanEnabled); 
    } else {
        
        document.getElementById('history-toggle').classList.toggle('active');
        window.addLog("History Scan is not yet implemented in the backend engine.");
    }
};

