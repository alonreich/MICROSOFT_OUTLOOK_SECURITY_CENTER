// Centralized Sanitization & HTML Entity Encoder Fallback
window.escapeHtml = window.escapeHtml || function(str) {
    if (str === null || str === undefined) return '';
    if (typeof str !== 'string') str = String(str);
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
};
window.escapeHTML = window.escapeHtml;

window.SecurityUI = window.SecurityUI || {};
SecurityUI.Settings = (function() {
    const api = window.securityApi;

    async function sync() {
        const config = await api.getConfig();
        let startupInfo = { configEnabled: !!config.launchAtStartup, systemActive: !!config.launchAtStartup };
        try {
            if (api.checkStartup) {
                startupInfo = await api.checkStartup();
            }
        } catch (e) {}

        const startupActive = startupInfo.systemActive !== undefined ? startupInfo.systemActive : !!config.launchAtStartup;
        const hasVtKey = !!(config.vtApiKey && config.vtApiKey !== 'MASKED_FOR_SECURITY');

        const body = document.getElementById('settings-body');
        body.innerHTML = `
            <!-- WINDOWS BOOT & SYSTEM INTEGRATION SECTION -->
            <div style="background: var(--surface-low); padding: 18px 20px; border-radius: 8px; border: 1px solid var(--border); margin-bottom: 24px;">
                <div style="display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <div style="font-weight: 600; font-size: 0.95rem; color: var(--text); display: flex; align-items: center; gap: 8px;">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="var(--accent)"><path d="M3 12V6.75l6-1.5v6.75H3zm7 0V5l11-1.5V12h-11zm-7 1.25h6V20l-6-1.5v-5.25zm7 0h11v8.5L10 20.5v-7.25z"/></svg>
                            Run on Windows Startup (Boot)
                        </div>
                        <div style="font-size: 0.8rem; color: var(--muted); margin-top: 4px;">
                            Automatically launch DeskGuard in background minimized to systray when Windows starts up.
                        </div>
                        <div id="startup-status-badge" style="margin-top: 8px;">
                            <span class="badge ${startupActive ? 'badge-safe' : 'badge-spam'}" style="font-size: 0.72rem; padding: 4px 10px;">
                                ${startupActive ? '● Windows Run Key Active (Auto-Start Enabled)' : '○ Auto-Start Disabled (Manual Launch Only)'}
                            </span>
                        </div>
                    </div>
                    <label class="switch" style="margin-left: 20px;">
                        <input type="checkbox" id="cfg-startup" ${startupActive ? 'checked' : ''} onchange="SecurityUI.Settings.onStartupToggle(this)">
                        <span class="slider"></span>
                    </label>
                </div>
            </div>

            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
                <div>
                    <h3 style="margin-top: 0; font-size: 0.9rem; text-transform: uppercase; letter-spacing: 0.5px; color: var(--accent);">Protection Engine</h3>
                    <div style="margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; background: var(--surface-low); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <div>
                            <div style="font-weight: 600; font-size: 0.85rem;">Active Protection</div>
                            <div style="font-size: 0.75rem; color: var(--muted);">Real-time MAPI mailbox surveillance</div>
                        </div>
                        <label class="switch">
                            <input type="checkbox" id="cfg-enabled" ${config.enabled ? 'checked' : ''}>
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="margin-bottom: 20px;">
                        <label for="cfg-speed">Mailbox Crawl Pace</label>
                        <input type="range" id="cfg-speed" min="10" max="100" value="${config.scanningSpeed || 50}">
                        <div style="display: flex; justify-content: space-between; font-size: 0.7rem; color: var(--muted); margin-top: 4px;">
                            <span>Gentle (Low CPU)</span>
                            <span>Balanced</span>
                            <span>High Throughput</span>
                        </div>
                    </div>
                    <div style="background: var(--surface-low); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border); margin-bottom: 20px;">
                        <div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 4px; color: var(--text);">Scoring Engines & Sensitivity</div>
                        <div style="font-size: 0.75rem; color: var(--muted); margin-bottom: 12px; line-height: 1.4;">
                            Configure 8 detection engines (DMARC, SPF, DKIM, PTR, Alignment, Heuristics, Body Entropy, RBL) and customize percentage weights.
                        </div>
                        <button type="button" onclick="SecurityUI.Engines.open()" class="btn-ui" style="width: 100%; justify-content: center; background: rgba(0, 120, 212, 0.15); border: 1px solid var(--accent); color: var(--accent); font-weight: 600; font-size: 0.8rem; padding: 8px 12px; border-radius: 6px; cursor: pointer;">
                            Configure Engines & Scoring Weights &rarr;
                        </button>
                    </div>

                    <div style="background: var(--surface-low); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border); margin-bottom: 20px;">
                        <div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 12px; color: var(--text);">Scan Modes & Priority</div>
                        
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
                            <div>
                                <div style="font-size: 0.8rem; font-weight: 600;">On-Access (Real-Time Protection)</div>
                                <div style="font-size: 0.7rem; color: var(--muted);">Scans new emails immediately upon arrival with top priority</div>
                            </div>
                            <label class="switch">
                                <input type="checkbox" id="cfg-onaccess" ${config.onAccessEnabled !== false ? 'checked' : ''}>
                                <span class="slider"></span>
                            </label>
                        </div>

                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
                            <div>
                                <div style="font-size: 0.8rem; font-weight: 600;">On-Demand (Light History Scan)</div>
                                <div style="font-size: 0.7rem; color: var(--muted);">Scans latest 1,000 emails per store from newest to oldest</div>
                            </div>
                            <label class="switch">
                                <input type="checkbox" id="cfg-ondemand" ${config.historyScanEnabled !== false ? 'checked' : ''}>
                                <span class="slider"></span>
                            </label>
                        </div>

                        <div style="display: flex; align-items: center; justify-content: space-between;">
                            <div>
                                <div style="font-size: 0.8rem; font-weight: 600;">Deep History Crawl</div>
                                <div style="font-size: 0.7rem; color: var(--muted);">Crawls beyond 1,000 emails across all PST archives (background)</div>
                            </div>
                            <label class="switch">
                                <input type="checkbox" id="cfg-deephistory" ${config.deepHistoryScanEnabled ? 'checked' : ''}>
                                <span class="slider"></span>
                            </label>
                        </div>
                    </div>
                </div>
                <div>
                    <h3 style="margin-top: 0; font-size: 0.9rem; text-transform: uppercase; letter-spacing: 0.5px; color: var(--accent);">Threat Intelligence</h3>
                    <div style="margin-bottom: 16px;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                            <label for="cfg-vtkey" style="margin-bottom: 0;">VirusTotal API Key</label>
                            <span class="badge ${hasVtKey ? 'badge-safe' : 'badge-spam'}" style="font-size: 0.65rem;">
                                ${hasVtKey ? 'Connected (DPAPI)' : 'Not Connected'}
                            </span>
                        </div>
                        <input type="password" id="cfg-vtkey" value="${window.escapeHtml(config.vtApiKey || '')}" placeholder="Paste 64-character API key...">
                        <div style="font-size: 0.72rem; color: var(--muted); margin-top: 4px;">
                            Get a free key: <a href="https://www.virustotal.com/gui/my-apikey" target="_blank" style="color: var(--accent); text-decoration: underline;">VirusTotal API Keys &rarr;</a>
                        </div>
                    </div>
                    <div style="margin-bottom: 20px;">
                        <label for="cfg-intel-level">Threat Intelligence Level</label>
                        <select id="cfg-intel-level">
                            <option value="0" ${config.threatIntelligenceLevel == 0 ? 'selected' : ''}>Conservative (Local Signatures Only)</option>
                            <option value="1" ${config.threatIntelligenceLevel == 1 ? 'selected' : ''}>Standard (Balanced Local + On-Demand VT)</option>
                            <option value="2" ${config.threatIntelligenceLevel == 2 ? 'selected' : ''}>Aggressive (Cloud Scan All Attachments)</option>
                        </select>
                    </div>
                </div>
            </div>
            <div style="margin-top: 20px; border-top: 1px solid var(--border); padding-top: 20px; display: flex; gap: 10px; flex-wrap: wrap;">
                <button onclick="api.openLogsFolder()">Open System Logs Directory</button>
                <button onclick="api.exportConfig()">Export Security Policy</button>
                <button onclick="api.importConfig()">Import Security Policy</button>
            </div>
        `;

        document.getElementById('save-settings-btn').onclick = async () => {
            const btn = document.getElementById('save-settings-btn');
            const oldText = btn.textContent;
            btn.textContent = 'Saving...';
            btn.disabled = true;

            try {
                const startupEnabled = document.getElementById('cfg-startup').checked;
                const onAccessEnabled = document.getElementById('cfg-onaccess').checked;
                const historyScanEnabled = document.getElementById('cfg-ondemand').checked;
                const deepHistoryScanEnabled = document.getElementById('cfg-deephistory').checked;

                await api.setEnabled(document.getElementById('cfg-enabled').checked);
                await api.setStartup(startupEnabled);
                await api.setScanningSpeed(parseInt(document.getElementById('cfg-speed').value));
                await api.setHistoryEnabled(historyScanEnabled);
                if (api.setScanModes) {
                    await api.setScanModes({ onAccessEnabled, onDemandLimit: 1000, deepHistoryScanEnabled });
                }
                await api.setVTKey(document.getElementById('cfg-vtkey').value);
                await api.setThreatIntelLevel(parseInt(document.getElementById('cfg-intel-level').value));

                window.addLog(`Security policy saved: Boot ${startupEnabled ? '[ON]' : '[OFF]'} | OnAccess ${onAccessEnabled ? '[ON]' : '[OFF]'} | LightScan ${historyScanEnabled ? '[ON]' : '[OFF]'} | DeepHistory ${deepHistoryScanEnabled ? '[ON]' : '[OFF]'}`);
                window.closeSettings();
            } catch (err) {
                console.error("Save Error:", err);
                window.addLog("Error saving settings: " + err.message);
                alert("Failed to save settings. Please check the logs.");
            } finally {
                btn.textContent = oldText;
                btn.disabled = false;
            }
        };
    }

    function onStartupToggle(checkbox) {
        const badge = document.getElementById('startup-status-badge');
        if (!badge) return;
        if (checkbox.checked) {
            badge.innerHTML = `<span class="badge badge-safe" style="font-size: 0.72rem; padding: 4px 10px;">● Windows Run Key Active (Auto-Start Enabled)</span>`;
        } else {
            badge.innerHTML = `<span class="badge badge-spam" style="font-size: 0.72rem; padding: 4px 10px;">○ Auto-Start Disabled (Manual Launch Only)</span>`;
        }
    }

    return { 
        sync: sync,
        onStartupToggle: onStartupToggle
    };
})();

window.syncSettingsUI = SecurityUI.Settings.sync;
