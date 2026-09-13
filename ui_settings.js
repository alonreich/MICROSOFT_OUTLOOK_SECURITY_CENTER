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
        const body = document.getElementById('settings-body');
        body.innerHTML = `
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
                <div>
                    <h3 style="margin-top: 0;">Protection Controls</h3>
                    <div style="margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                        <span>Security Protection Active</span>
                        <label class="switch">
                            <input type="checkbox" id="cfg-enabled" ${config.enabled ? 'checked' : ''}>
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                        <div>
                            <div style="font-weight: 500;">Launch at Startup</div>
                            <div style="font-size: 11px; color: var(--muted, #8b949e); margin-top: 2px;">Start minimized to system tray when Windows boots</div>
                        </div>
                        <label class="switch">
                            <input type="checkbox" id="cfg-startup" ${config.launchAtStartup ? 'checked' : ''}>
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="margin-bottom: 20px;">
                        <label>Scanning Speed (Resource Usage)</label>
                        <input type="range" id="cfg-speed" min="10" max="100" value="${config.scanningSpeed || 50}">
                    </div>
                </div>
                <div>
                    <h3 style="margin-top: 0;">Threat Intelligence</h3>
                    <div style="margin-bottom: 20px;">
                        <label>VirusTotal API Key (Encrypted)</label>
                        <input type="password" id="cfg-vtkey" value="${window.escapeHtml(config.vtApiKey || '')}" placeholder="Paste API key here...">
                    </div>
                    <div style="margin-bottom: 20px;">
                        <label>Intel Threshold</label>
                        <select id="cfg-intel-level">
                            <option value="0" ${config.threatIntelligenceLevel == 0 ? 'selected' : ''}>Conservative (Local Only)</option>
                            <option value="1" ${config.threatIntelligenceLevel == 1 ? 'selected' : ''}>Standard (Balanced)</option>
                            <option value="2" ${config.threatIntelligenceLevel == 2 ? 'selected' : ''}>Aggressive (Cloud Scan All)</option>
                        </select>
                    </div>
                </div>
            </div>
            <div style="margin-top: 20px; border-top: 1px solid var(--border); padding-top: 20px;">
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
                await api.setEnabled(document.getElementById('cfg-enabled').checked);
                await api.setStartup(startupEnabled);
                await api.setScanningSpeed(parseInt(document.getElementById('cfg-speed').value));
                await api.setVTKey(document.getElementById('cfg-vtkey').value);
                await api.setThreatIntelLevel(parseInt(document.getElementById('cfg-intel-level').value));

                window.addLog(`Security policy updated: Startup on boot ${startupEnabled ? '[Enabled]' : '[Disabled]'}`);
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

    return { sync: sync };
})();

window.syncSettingsUI = SecurityUI.Settings.sync;
