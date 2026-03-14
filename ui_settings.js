document.body.insertAdjacentHTML('beforeend', `
<div id="settings-modal" class="modal">
    <div class="modal-content" style="width: 950px; height: 85vh;">
        <div class="modal-header" style="display: flex; align-items: center; justify-content: center; gap: 15px;">
            <svg style="width:24px;height:24px;fill:var(--accent)" viewBox="0 0 24 24"><path d="M12,15.5A3.5,3.5 0 0,1 8.5,12A3.5,3.5 0 0,1 12,8.5A3.5,3.5 0 0,1 15.5,12A3.5,3.5 0 0,1 12,15.5M19.43,12.97C19.47,12.65 19.5,12.33 19.5,12C19.5,11.67 19.47,11.35 19.43,11.03L21.54,9.37C21.73,9.22 21.78,8.95 21.66,8.73L19.66,5.27C19.54,5.05 19.27,4.97 19.05,5.05L16.56,5.9C16.04,5.53 15.48,5.22 14.87,5.03L14.5,2.35C14.46,2.12 14.27,1.95 14.04,1.95H10.04C9.81,1.95 9.62,2.12 9.58,2.35L9.21,5.03C8.6,5.22 8.04,5.53 7.52,5.9L5.03,5.05C4.81,4.97 4.54,5.05 4.42,5.27L2.42,8.73C2.3,8.95 2.35,9.22 2.54,9.37L4.65,11.03C4.61,11.35 4.58,11.67 4.58,12C4.58,12.33 4.61,12.65 4.65,12.97L2.54,14.63C2.35,14.78 2.3,15.05 2.42,15.27L4.42,18.73C4.54,18.95 4.81,19.03 5.03,18.95L7.52,18.1C8.04,18.47 8.6,18.78 9.21,18.97L9.58,21.65C9.62,21.88 9.81,22.05 10.04,22.05H14.04C14.27,22.05 14.46,21.88 14.5,21.65L14.87,18.97C15.48,18.78 16.04,18.47 16.56,18.1L19.05,18.95C19.27,19.03 19.54,18.95 19.66,18.73L21.66,15.27C21.78,15.05 21.73,14.78 21.54,14.63L19.43,12.97Z" /></svg>
            SYSTEM CONFIGURATION & SECURITY POLICY
        </div>
        <div style="overflow-y: auto; flex: 1; padding-right: 15px; margin-bottom: 10px;">
            <div class="section-divider" data-label="External Intelligence"></div>
            <div class="setting-group">
                <label><svg style="width:14px;height:14px;fill:currentColor" viewBox="0 0 24 24"><path d="M7,14L12,19L17,14H14V5H10V14H7Z" /></svg>VirusTotal Enterprise API Key</label>
                <div style="position:relative">
                    <input type="password" id="vt-api-key" placeholder="Enter your 64-character API key..." style="width:100%; padding-right:40px;">
                    <div id="toggle-vt-visibility" style="position:absolute; right:12px; top:50%; transform:translateY(-50%); cursor:pointer; color:var(--muted);"><svg style="width:18px;height:18px;fill:currentColor" viewBox="0 0 24 24"><path d="M12,9A3,3 0 0,0 9,12A3,3 0 0,0 12,15A3,3 0 0,0 15,12A3,3 0 0,0 12,9M12,17A5,5 0 0,1 7,12A5,5 0 0,1 12,7A5,5 0 0,1 17,12A5,5 0 0,1 12,17M12,4.5C7,4.5 2.73,7.61 1,12C2.73,16.39 7,19.5 12,19.5C17,19.5 21.27,16.39 23,12C21.27,7.61 17,4.5 12,4.5Z" /></svg></div>
                </div>
                <div style="font-size:0.6rem; color:var(--muted); margin-top:5px; font-style:italic;">Used for real-time IP reputation and attachment analysis.</div>
            </div>

            <div class="section-divider" data-label="System Behavior"></div>
            <div style="display:flex; align-items:center; gap:12px; background:rgba(142, 117, 255, 0.05); padding:15px; border-radius:10px; border:1px solid var(--border);">
                <input type="checkbox" id="launch-at-startup" style="width:18px; height:18px; cursor:pointer;">
                <label for="launch-at-startup" style="cursor:pointer; flex:1; font-size:0.8rem; font-weight:700; color:#fff; text-transform:none; letter-spacing:0;">Run Automatically When Windows Starts</label>
            </div>

            <div class="section-divider" data-label="Trust & Reputation Management"></div>
            <div class="setting-grid">
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--ok)" viewBox="0 0 24 24"><path d="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" /></svg>Trusted Senders (Emails)</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="wl-emails" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="wl-emails" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="wl-emails" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--danger)" viewBox="0 0 24 24"><path d="M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z" /></svg>Blocked Senders (Emails)</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="bl-emails" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="bl-emails" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="bl-emails" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
            </div>
            <div class="setting-grid">
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--ok)" viewBox="0 0 24 24"><path d="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" /></svg>Trusted IP Addresses</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="wl-ips" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="wl-ips" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="wl-ips" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--danger)" viewBox="0 0 24 24"><path d="M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z" /></svg>Blocked IP Addresses</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="bl-ips" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="bl-ips" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="bl-ips" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
            </div>
            <div class="setting-grid">
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--ok)" viewBox="0 0 24 24"><path d="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" /></svg>Trusted Domains</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="wl-domains" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="wl-domains" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="wl-domains" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--danger)" viewBox="0 0 24 24"><path d="M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z" /></svg>Blocked Domains</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="bl-domains" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="bl-domains" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="bl-domains" placeholder="One per line..." style="height:200px;"></textarea>
                </div>
            </div>
            <div class="setting-grid">
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--ok)" viewBox="0 0 24 24"><path d="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" /></svg>Trusted IP|Domain Combos</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="wl-combos" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="wl-combos" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="wl-combos" placeholder="IP|domain.com (One per line)..." style="height:200px;"></textarea>
                </div>
                <div class="setting-group">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <label><svg style="width:14px;height:14px;fill:var(--danger)" viewBox="0 0 24 24"><path d="M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z" /></svg>Blocked IP|Domain Combos</label>
                        <div style="display:flex; gap:4px;">
                            <button class="btn-step csv-io" data-action="import" data-target="bl-combos" style="font-size:0.5rem; padding:2px 4px;">IMP</button>
                            <button class="btn-step csv-io" data-action="export" data-target="bl-combos" style="font-size:0.5rem; padding:2px 4px;">EXP</button>
                        </div>
                    </div>
                    <textarea id="bl-combos" placeholder="IP|domain.com (One per line)..." style="height:200px;"></textarea>
                </div>
            </div>

            <div class="section-divider" data-label="Content Filtering"></div>
            <div class="setting-group">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label><svg style="width:14px;height:14px;fill:currentColor" viewBox="0 0 24 24"><path d="M11,18H13V16H11V18M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,20C7.59,20 4,15.73 4,12C4,8.27 7.59,4 12,4C16.41,4 20,8.27 20,12C20,15.73 16.41,20 12,20M12,6A4,4 0 0,0 8,10H10A2,2 0 0,1 12,8A2,2 0 0,1 14,10C14,12 11,11.75 11,15H13C13,12.75 16,12.5 16,10A4,4 0 0,0 12,6Z" /></svg>Global Spam Keywords</label>
                    <div style="display:flex; gap:8px;">
                        <input type="file" id="import-keywords-input" accept=".csv,.txt" style="display:none">
                        <button id="import-keywords-btn" class="btn-step" style="font-size:0.6rem; padding:4px 8px; height:auto; width:auto;">IMPORT CSV</button>
                        <button id="export-keywords-btn" class="btn-step" style="font-size:0.6rem; padding:4px 8px; height:auto; width:auto;">EXPORT CSV</button>
                    </div>
                </div>
                <textarea id="spam-keywords" placeholder="Enter keywords to flag as spam (one per line)..." style="height:300px;"></textarea>
                <div style="font-size:0.6rem; color:var(--muted); margin-top:5px; font-style:italic;">Case-insensitive matching in subject and body.</div>
            </div>
        </div>
        <div style="display:flex; justify-content: center; gap:15px; margin-top:15px; padding-top:15px; border-top:1px solid var(--border);">
            <button id="export-app-config" class="btn-ui" style="font-size:0.7rem;">EXPORT CONFIGURATION</button>
            <button id="import-app-config" class="btn-ui" style="font-size:0.7rem;">IMPORT CONFIGURATION</button>
        </div>
        <div style="display:flex; justify-content: center; gap:20px; padding-top:15px; margin-top:15px; border-top:1px solid var(--border);">
            <button id="save-settings" class="btn-ui success" style="width:180px; height:45px; font-size:0.9rem;">SAVE CHANGES</button>
            <button id="close-settings" class="btn-ui danger" style="width:180px; height:45px; font-size:0.9rem;">DISCARD</button>
        </div>
    </div>
</div>
`);

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
    };

    document.getElementById('settings-btn').onclick = async () => { 
        document.getElementById('settings-modal').style.display = 'flex';
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
        };
        reader.readAsText(file);
    };

    document.getElementById('export-app-config').onclick = async () => {
        const res = await api.exportConfig();
        if (res.success) {
            alert('Configuration exported successfully.');
        }
    };

    document.getElementById('import-app-config').onclick = async () => {
        const res = await api.importConfig();
        if (res.success) {
            alert('Configuration imported successfully. Refreshing settings...');
            const cfg = await api.getConfig(); 
            document.getElementById('vt-api-key').value = cfg.vtApiKey || ''; 
            document.getElementById('wl-emails').value = (cfg.whitelist?.emails || []).join('\n');
            document.getElementById('bl-emails').value = (cfg.blacklist?.emails || []).join('\n');
            document.getElementById('wl-ips').value = (cfg.whitelist?.ips || []).join('\n');
            document.getElementById('bl-ips').value = (cfg.blacklist?.ips || []).join('\n');
            document.getElementById('wl-domains').value = (cfg.whitelist?.domains || []).join('\n');
            document.getElementById('bl-domains').value = (cfg.blacklist?.domains || []).join('\n');
            document.getElementById('wl-combos').value = (cfg.whitelist?.combos || []).join('\n');
            document.getElementById('bl-combos').value = (cfg.blacklist?.combos || []).join('\n');
            document.getElementById('spam-keywords').value = (cfg.spamKeywords || []).join('\n');
            document.getElementById('launch-at-startup').checked = !!cfg.launchAtStartup;
        } else if (res.error) {
            alert('Error importing configuration: ' + res.error);
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
        
        document.getElementById('settings-modal').style.display = 'none';
    };

    document.getElementById('close-settings').onclick = () => document.getElementById('settings-modal').style.display = 'none';
})();
