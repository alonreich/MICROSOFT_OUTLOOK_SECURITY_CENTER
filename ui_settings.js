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

    const intelExplanations = {
        0: {
            title: "Level 0: Local PC Only (Maximum Privacy)",
            badge: "Offline / Private",
            badgeClass: "badge-safe",
            summary: "All email security checks happen 100% inside your computer. DeskGuard inspects sender authentication (SPF, DKIM, DMARC), email headers, and heuristic patterns strictly on your PC. No file hashes, sender information, or queries are ever sent over the internet.",
            bestFor: "Best for: High-privacy environments, offline computers, and strict no-cloud policies."
        },
        1: {
            title: "Level 1: Smart Community Defense (Recommended)",
            badge: "Smart Hybrid",
            badgeClass: "badge-safe",
            summary: "DeskGuard checks your email locally first. Only if an email looks suspicious or unverified does it query global threat intelligence (Team Cymru DNS hash lookup and VirusTotal if an API key is connected). Safe emails from regular senders are never checked against the cloud.",
            bestFor: "Best for: 99% of users. Maximum protection, instant speeds, and zero unnecessary data sharing."
        },
        2: {
            title: "Level 2: Maximum Cloud Inspection (Aggressive)",
            badge: "Aggressive",
            badgeClass: "badge-spam",
            summary: "Every incoming attachment, link, and file hash is rigorously checked against 70+ cloud antivirus engines via VirusTotal and community databases, regardless of whether the email looks safe. Delivers zero-tolerance malware protection.",
            bestFor: "Best for: High-risk environments, finance/exec accounts, and users frequently opening unknown files."
        }
    };

    function updateIntelExplainer(val) {
        const explainerEl = document.getElementById('intel-explainer-card');
        if (!explainerEl) return;
        const info = intelExplanations[val] || intelExplanations[1];
        explainerEl.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <div style="font-weight: 700; font-size: 0.8rem; color: var(--text);">${info.title}</div>
                <span class="badge ${info.badgeClass}" style="font-size: 0.65rem;">${info.badge}</span>
            </div>
            <div style="font-size: 0.74rem; color: var(--muted); line-height: 1.45; margin-bottom: 6px;">
                ${info.summary}
            </div>
            <div style="font-size: 0.72rem; color: var(--accent); font-weight: 600;">
                ${info.bestFor}
            </div>
        `;
    }

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
        const currentIntelLevel = (config.threatIntelligenceLevel !== undefined && config.threatIntelligenceLevel !== null) ? parseInt(config.threatIntelligenceLevel) : 1;

        const body = document.getElementById('settings-body');
        body.innerHTML = `
            <!-- WINDOWS BOOT & SYSTEM INTEGRATION SECTION -->
            <div style="background: var(--surface-low); padding: 18px 20px; border-radius: 8px; border: 1px solid var(--border); margin-bottom: 24px;">
                <h3 style="text-align: center; margin-top: 0; margin-bottom: 14px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.8px; color: var(--accent);">Windows Boot &amp; System Integration</h3>
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
                <!-- PROTECTION ENGINE & PERFORMANCE SECTION -->
                <div style="background: var(--surface-low); padding: 18px 20px; border-radius: 8px; border: 1px solid var(--border);">
                    <h3 style="text-align: center; margin-top: 0; margin-bottom: 18px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.8px; color: var(--accent);">Protection Engine &amp; Performance</h3>
                    
                    <div style="margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <div>
                            <div style="font-weight: 600; font-size: 0.85rem;">Active Protection</div>
                            <div style="font-size: 0.75rem; color: var(--muted);">Real-time MAPI mailbox surveillance</div>
                        </div>
                        <label class="switch">
                            <input type="checkbox" id="cfg-enabled" ${config.enabled ? 'checked' : ''}>
                            <span class="slider"></span>
                        </label>
                    </div>

                    <div style="margin-bottom: 20px; background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <label for="cfg-speed" style="font-weight: 600; font-size: 0.85rem; margin-bottom: 6px; display: block;">Emails Scan Speed</label>
                        <input type="range" id="cfg-speed" min="10" max="100" value="${config.scanningSpeed || 50}" style="width: 100%;">
                        <div style="display: flex; justify-content: space-between; font-size: 0.7rem; color: var(--muted); margin-top: 4px;">
                            <span>Gentle (Low CPU)</span>
                            <span>Balanced</span>
                            <span>High Throughput</span>
                        </div>
                    </div>

                    <div style="background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border); margin-bottom: 20px;">
                        <div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 4px; color: var(--text);">Scoring Engines &amp; Sensitivity</div>
                        <div style="font-size: 0.75rem; color: var(--muted); margin-bottom: 12px; line-height: 1.4;">
                            Configure 8 detection engines (DMARC, SPF, DKIM, PTR, Alignment, Heuristics, Body Entropy, RBL) and customize percentage weights.
                        </div>
                        <button type="button" onclick="SecurityUI.Engines.open()" class="btn-ui" style="width: 100%; justify-content: center; background: rgba(0, 120, 212, 0.15); border: 1px solid var(--accent); color: var(--accent); font-weight: 600; font-size: 0.8rem; padding: 8px 12px; border-radius: 6px; cursor: pointer;">
                            Configure Engines &amp; Scoring Weights &rarr;
                        </button>
                    </div>

                    <div style="background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 12px; color: var(--text);">Scan Modes &amp; Priority</div>
                        
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

                <!-- THREAT INTELLIGENCE & CLOUD DETECTION SECTION -->
                <div style="background: var(--surface-low); padding: 18px 20px; border-radius: 8px; border: 1px solid var(--border);">
                    <h3 style="text-align: center; margin-top: 0; margin-bottom: 18px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.8px; color: var(--accent);">Threat Intelligence &amp; Cloud Detection</h3>
                    
                    <div style="margin-bottom: 16px; background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                            <label for="cfg-vtkey" style="margin-bottom: 0; font-weight: 600; font-size: 0.85rem;">VirusTotal API Key</label>
                            <span class="badge ${hasVtKey ? 'badge-safe' : 'badge-spam'}" style="font-size: 0.65rem;">
                                ${hasVtKey ? 'Connected (DPAPI Encrypted)' : 'Not Connected'}
                            </span>
                        </div>
                        <input type="password" id="cfg-vtkey" value="${window.escapeHtml(config.vtApiKey || '')}" placeholder="Paste 64-character API key..." style="width: 100%;">
                        <div style="font-size: 0.72rem; color: var(--muted); margin-top: 6px;">
                            Connect your free VirusTotal key to scan suspicious attachments against 70+ antivirus engines: 
                            <a href="https://www.virustotal.com/gui/my-apikey" target="_blank" style="color: var(--accent); text-decoration: underline;">Get Free VirusTotal Key &rarr;</a>
                        </div>
                    </div>

                    <div style="margin-bottom: 16px; background: var(--surface-high); padding: 14px 16px; border-radius: 6px; border: 1px solid var(--border);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                            <label for="cfg-intel-level" style="margin-bottom: 0; font-weight: 600; font-size: 0.85rem;">Threat Intelligence Level</label>
                            <span title="Threat Intelligence lets DeskGuard consult worldwide threat databases to identify zero-day viruses and phishing links." style="cursor: help; font-size: 0.75rem; color: var(--accent); font-weight: 700; border: 1px solid var(--accent); border-radius: 50%; width: 16px; height: 16px; display: inline-flex; align-items: center; justify-content: center;">?</span>
                        </div>
                        <select id="cfg-intel-level" style="width: 100%; margin-bottom: 10px;" onchange="SecurityUI.Settings.onIntelChange(this.value)">
                            <option value="0" ${currentIntelLevel === 0 ? 'selected' : ''}>Level 0: Local PC Only (Conservative - No Cloud)</option>
                            <option value="1" ${currentIntelLevel === 1 ? 'selected' : ''}>Level 1: Smart Community Defense (Balanced - Recommended)</option>
                            <option value="2" ${currentIntelLevel === 2 ? 'selected' : ''}>Level 2: Maximum Cloud Inspection (Aggressive - Cloud All)</option>
                        </select>
                        <div id="intel-explainer-card" style="background: var(--surface-low); border: 1px solid var(--border); border-radius: 6px; padding: 10px 12px;"></div>
                    </div>
                </div>
            </div>

            <!-- MAINTENANCE, UPGRADE & HARD RESET SECTION -->
            <div style="background: var(--surface-low); padding: 18px 20px; border-radius: 8px; border: 1px solid var(--border); margin-top: 24px;">
                <h3 style="text-align: center; margin-top: 0; margin-bottom: 16px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.8px; color: var(--accent);">Maintenance, Upgrade &amp; Storage Reset</h3>
                <div style="display: flex; gap: 12px; flex-wrap: wrap; justify-content: center;">
                    <button type="button" onclick="api.openLogsFolder()" class="btn-ui" style="padding: 8px 14px;">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                        Open System Logs
                    </button>
                    <button type="button" onclick="api.exportConfig()" class="btn-ui" style="padding: 8px 14px;">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        Export Policy
                    </button>
                    <button type="button" onclick="api.importConfig()" class="btn-ui" style="padding: 8px 14px;">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
                        Import Policy
                    </button>
                    <button type="button" onclick="window.safeUpgrade()" class="btn-ui" style="padding: 8px 14px; background: rgba(0, 120, 212, 0.15); border: 1px solid var(--accent); color: var(--accent); font-weight: 600;">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>
                        Safe Upgrade / Cache Refresh
                    </button>
                    <button type="button" onclick="window.nuclearReset()" class="btn-ui" style="padding: 8px 14px; background: rgba(241, 112, 123, 0.15); border: 1px solid var(--danger); color: var(--danger); font-weight: 600;">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>
                        Scorched Earth Reset
                    </button>
                </div>
            </div>
        `;

        updateIntelExplainer(currentIntelLevel);

        const unsavedInd = document.getElementById('settings-unsaved-indicator');
        if (unsavedInd) unsavedInd.style.display = 'none';

        function markUnsaved() {
            if (unsavedInd) unsavedInd.style.display = 'inline-block';
        }

        ['cfg-startup', 'cfg-enabled', 'cfg-speed', 'cfg-onaccess', 'cfg-ondemand', 'cfg-deephistory', 'cfg-vtkey', 'cfg-intel-level'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('input', markUnsaved);
                el.addEventListener('change', markUnsaved);
            }
        });

        async function doSave() {
            const btnBottom = document.getElementById('save-settings-btn');
            const btnTop = document.getElementById('save-settings-btn-top');
            const oldBottom = btnBottom ? btnBottom.innerHTML : '';
            const oldTop = btnTop ? btnTop.innerHTML : '';
            if (btnBottom) { btnBottom.textContent = 'Saving...'; btnBottom.disabled = true; }
            if (btnTop) { btnTop.textContent = 'Saving...'; btnTop.disabled = true; }

            try {
                const startupEnabled = document.getElementById('cfg-startup').checked;
                const onAccessEnabled = document.getElementById('cfg-onaccess').checked;
                const onDemandChecked = document.getElementById('cfg-ondemand').checked;
                const deepHistoryChecked = document.getElementById('cfg-deephistory').checked;
                const historyScanEnabled = onDemandChecked || deepHistoryChecked;
                const deepHistoryScanEnabled = deepHistoryChecked;

                await api.setEnabled(document.getElementById('cfg-enabled').checked);
                await api.setStartup(startupEnabled);
                await api.setScanningSpeed(parseInt(document.getElementById('cfg-speed').value));
                await api.setHistoryEnabled(historyScanEnabled);
                if (api.setScanModes) {
                    await api.setScanModes({ onAccessEnabled, onDemandLimit: 1000, deepHistoryScanEnabled });
                }
                await api.setVTKey(document.getElementById('cfg-vtkey').value);
                await api.setThreatIntelLevel(parseInt(document.getElementById('cfg-intel-level').value));

                if (unsavedInd) unsavedInd.style.display = 'none';
                window.addLog(`Security policy saved: Boot ${startupEnabled ? '[ON]' : '[OFF]'} | OnAccess ${onAccessEnabled ? '[ON]' : '[OFF]'} | LightScan ${historyScanEnabled ? '[ON]' : '[OFF]'} | DeepHistory ${deepHistoryScanEnabled ? '[ON]' : '[OFF]'}`);
                window.closeSettings();
            } catch (err) {
                console.error("Save Error:", err);
                window.addLog("Error saving settings: " + err.message);
                alert("Failed to save settings. Please check the logs.");
            } finally {
                if (btnBottom) { btnBottom.innerHTML = oldBottom; btnBottom.disabled = false; }
                if (btnTop) { btnTop.innerHTML = oldTop; btnTop.disabled = false; }
            }
        }

        const btnBottom = document.getElementById('save-settings-btn');
        if (btnBottom) btnBottom.onclick = doSave;

        const btnTop = document.getElementById('save-settings-btn-top');
        if (btnTop) btnTop.onclick = doSave;
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
        onStartupToggle: onStartupToggle,
        onIntelChange: updateIntelExplainer
    };
})();

window.syncSettingsUI = SecurityUI.Settings.sync;
