window.SecurityUI = window.SecurityUI || {};
SecurityUI.Reset = (function() {
    const api = window.securityApi;

    async function nuclear() {
        if (!window.showConfirmModal) {
            if (!confirm('CRITICAL SCORCHED EARTH RESET: This will permanently wipe ALL settings, API keys, blacklists, whitelists, and scan history. Proceed?')) return;
            if (!confirm('FINAL WARNING: This cannot be undone. Wipe everything and restart?')) return;
            window.addLog('Initiating Scorched Earth Factory Reset...');
            await api.resetApp('scorched');
            return;
        }

        window.showConfirmModal(
            'SCORCHED EARTH FACTORY RESET',
            'This action will permanently wipe ALL security configurations, VirusTotal API keys, blacklists, forensic snapshots, whitelists, and scan database. The Windows startup run key will also be uninstalled. Do you wish to proceed?',
            async () => {
                window.addLog('Initiating Scorched Earth Factory Reset...');
                if (window.showNotification) window.showNotification('SCORCHED EARTH WIPE INITIATED: All configuration and data erased. Application will restart now...', true);
                await api.resetApp('scorched');
            }
        );
    }

    async function safeUpgrade() {
        if (!window.showConfirmModal) {
            if (!confirm('SAFE UPGRADE / CACHE CLEANUP: This refreshes the database and clears cached mail history while PRESERVING all your settings, VirusTotal API keys, rules, and whitelists. Proceed?')) return;
            window.addLog('Initiating Safe Upgrade / Cache Refresh...');
            await api.resetApp('safe');
            return;
        }

        window.showConfirmModal(
            'SAFE UPGRADE & CACHE CLEANUP',
            'This safely clears the cached email incident list and refreshes the database, while completely preserving your settings, VirusTotal API keys, custom whitelists, and security engines. The application will restart clean. Proceed?',
            async () => {
                window.addLog('Initiating Safe Upgrade / Cache Refresh...');
                if (window.showNotification) window.showNotification('SAFE UPGRADE INITIATED: Resetting cache while preserving your settings. Application restarting...', true);
                await api.resetApp('safe');
            }
        );
    }

    return { 
        nuclear: nuclear,
        safeUpgrade: safeUpgrade
    };
})();

window.nuclearReset = SecurityUI.Reset.nuclear;
window.safeUpgrade = SecurityUI.Reset.safeUpgrade;
