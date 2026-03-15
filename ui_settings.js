(function() {
    const api = window.securityApi;
    window.syncSettingsUI = async () => {
        const cfg = await api.getConfig(); 
        const vt = document.getElementById('vt-api-key');
        if (vt) vt.value = cfg.vtApiKey || ''; 
        const wle = document.getElementById('wl-emails');
        if (wle) wle.value = (cfg.whitelist?.emails || []).join('\n');
        const ble = document.getElementById('bl-emails');
        if (ble) ble.value = (cfg.blacklist?.emails || []).join('\n');
        const wli = document.getElementById('wl-ips');
        if (wli) wli.value = (cfg.whitelist?.ips || []).join('\n');
        const bli = document.getElementById('bl-ips');
        if (bli) bli.value = (cfg.blacklist?.ips || []).join('\n');
        const wld = document.getElementById('wl-domains');
        if (wld) wld.value = (cfg.whitelist?.domains || []).join('\n');
        const bld = document.getElementById('bl-domains');
        if (bld) bld.value = (cfg.blacklist?.domains || []).join('\n');
        const wlc = document.getElementById('wl-combos');
        if (wlc) wlc.value = (cfg.whitelist?.combos || []).join('\n');
        const blc = document.getElementById('bl-combos');
        if (blc) blc.value = (cfg.blacklist?.combos || []).join('\n');
        const sk = document.getElementById('spam-keywords');
        if (sk) sk.value = (cfg.spamKeywords || []).join('\n');
        const la = document.getElementById('launch-at-startup');
        if (la) la.checked = !!cfg.launchAtStartup;
        
        const speed = cfg.scanningSpeed !== undefined ? cfg.scanningSpeed : 50;
        const slider = document.getElementById('scanning-speed-slider');
        if (slider) slider.value = speed;
        const valText = document.getElementById('scanning-speed-val');
        if (valText) valText.textContent = speed + '%';
    };

    document.getElementById('scanning-speed-slider').oninput = (e) => {
        const val = e.target.value;
        document.getElementById('scanning-speed-val').textContent = val + '%';
    };
    document.getElementById('scanning-speed-slider').onchange = (e) => {
        api.setScanningSpeed(parseInt(e.target.value));
        window.showNotification(`Scanning Engine Speed set to ${e.target.value}%`);
    };

    document.getElementById('reset-performance-btn').onclick = async () => {
        const slider = document.getElementById('scanning-speed-slider');
        slider.value = 50;
        document.getElementById('scanning-speed-val').textContent = '50%';
        await api.setScanningSpeed(50);
        window.showNotification('Performance profile reset to balanced defaults.');
    };

    document.getElementById('settings-btn').onclick = async () => { 
        document.getElementById('settings-modal').style.display = 'flex';
        // Reset to first tab
        document.querySelector('.tab-btn[data-tab="tab-general"]').click();
        await window.syncSettingsUI();
    };

    document.getElementById('toggle-vt-visibility').onclick = () => {
        const el = document.getElementById('vt-api-key');
        el.type = el.type === 'password' ? 'text' : 'password';
    };

    const handleIO = (e) => {
        const btn = e.target;
        const targetId = btn.dataset.target;
        const action = btn.dataset.action;
        if (action === 'export') {
            const val = document.getElementById(targetId).value;
            const blob = new Blob([val.split('\n').filter(s=>s.trim()).join(',')], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${targetId}.csv`;
            a.click();
            URL.revokeObjectURL(url);
            window.showNotification(`Exported ${targetId} to CSV.`);
        } else {
            const inp = document.createElement('input');
            inp.type = 'file';
            inp.accept = '.csv,.txt';
            inp.onchange = (ie) => {
                const file = ie.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (re) => {
                    const content = re.target.result;
                    const items = content.split(/[,\n\r]+/).map(s=>s.trim()).filter(Boolean);
                    document.getElementById(targetId).value = items.join('\n');
                    window.showNotification(`Imported ${items.length} items from CSV.`);
                };
                reader.readAsText(file);
            };
            inp.click();
        }
    };

    document.querySelectorAll('.csv-io').forEach(b => b.onclick = handleIO);

    document.getElementById('export-keywords-btn').onclick = () => {
        const text = document.getElementById('spam-keywords').value;
        const blob = new Blob([text.split('\n').filter(s=>s.trim()).join(',')], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spam_keywords.csv';
        a.click();
        URL.revokeObjectURL(url);
        window.showNotification('Exported spam keywords to CSV.');
    };

    document.getElementById('import-keywords-btn').onclick = () => document.getElementById('import-keywords-input').click();
    document.getElementById('import-keywords-input').onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (re) => {
            const content = re.target.result;
            const keywords = content.split(/[,\n\r]+/).map(s=>s.trim()).filter(Boolean);
            document.getElementById('spam-keywords').value = keywords.join('\n');
            window.showNotification(`Imported ${keywords.length} keywords.`);
        };
        reader.readAsText(file);
    };

    document.getElementById('export-app-config').onclick = async () => {
        const res = await api.exportConfig();
        if (res.success) {
            window.showNotification('Full system configuration exported successfully.');
        }
    };

    document.getElementById('import-app-config').onclick = async () => {
        const res = await api.importConfig();
        if (res.success) {
            window.showNotification('System configuration imported. Reloading UI...');
            await window.syncSettingsUI();
        } else if (res.error) {
            window.showNotification('Config Import Error: ' + res.error, true);
        }
    };

    document.getElementById('save-settings').onclick = async () => {
        const vt = document.getElementById('vt-api-key').value;
        const kw = document.getElementById('spam-keywords').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const wle = document.getElementById('wl-emails').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const ble = document.getElementById('bl-emails').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const wli = document.getElementById('wl-ips').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const bli = document.getElementById('bl-ips').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const wld = document.getElementById('wl-domains').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const bld = document.getElementById('bl-domains').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const wlc = document.getElementById('wl-combos').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const blc = document.getElementById('bl-combos').value.split('\n').map(s=>s.trim()).filter(Boolean);
        const startup = document.getElementById('launch-at-startup').checked;
        
        await api.setVTKey(vt); 
        await api.setSpamKeywords(kw);
        await api.setWhitelist({ emails: wle, ips: wli, domains: wld, combos: wlc });
        await api.setBlacklist({ emails: ble, ips: bli, domains: bld, combos: blc });
        await api.setStartup(startup);
        
        window.showNotification('Security policy and system settings saved.');
        document.getElementById('settings-modal').style.display = 'none';
    };

    document.getElementById('close-settings').onclick = () => document.getElementById('settings-modal').style.display = 'none';
    
    // Connect Nuclear Reset
    const resetBtn = document.getElementById('reset-btn');
    if (resetBtn) resetBtn.onclick = () => {
        if (typeof window.nuclearReset === 'function') {
            window.nuclearReset();
        }
    };
})();
