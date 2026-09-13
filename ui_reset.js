window.SecurityUI = window.SecurityUI || {};
SecurityUI.Reset = (function() {
    const api = window.securityApi;

    async function nuclear() {
        if (!window.showConfirmModal) {
            if (!confirm('CRITICAL: This will wipe all security configurations, blacklists, and scan history. Proceed?')) return;
            if (!confirm('FINAL WARNING: This will permanently delete ALL forensics and whitelists. Proceed?')) return;
            window.addLog('Initiating System Hard Reset...');
            await api.resetApp();
            return;
        }

        window.showConfirmModal(
            'FACTORY RESET: CRITICAL WARNING',
            'This action will permanently wipe ALL security configurations, blacklists, forensic snapshots, whitelists, and scan history. This process cannot be undone and the application will restart. Do you wish to proceed?',
            async () => {
                window.addLog('Initiating System Hard Reset...');
                window.showNotification('SYSTEM WIPE INITIATED: All data is being erased. Application will restart now...', true);
                await api.resetApp();
            }
        );
    }

    return { nuclear: nuclear };
})();

window.nuclearReset = SecurityUI.Reset.nuclear;
