(function() {
    const api = window.securityApi;
    document.getElementById('reset-btn').onclick = async () => {
        if (confirm('CRITICAL: This will wipe all security configurations, blacklists, and scan history. The application will then restart. Proceed?')) {
            window.addLog('Initiating System Hard Reset...');
            await api.resetApp();
        }
    };
})();

