window.SecurityUI = window.SecurityUI || {};
SecurityUI.Forensics = (function() {
    const api = window.securityApi;
    const modal = document.getElementById('forensic-modal');
    const fields = document.getElementById('forensic-fields');
    let currentId = null;

    function formatSize(bytes) {
        if (!bytes) return "0 B";
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    window.showForensics = async function(id) {
        currentId = id;
        const data = await api.getForensics(id);
        const item = window.AppStore.stats[window.AppStore.currentCategory].find(x => (x.entryId || x.fingerprint) === id);
        
        fields.innerHTML = `
            <div style="display:grid; grid-template-columns: 150px 1fr; gap: 10px; border-bottom: 1px solid var(--border); padding: 10px 0;">
                <div style="font-weight:bold; color:var(--muted)">SUBJECT</div>
                <div style="color:#fff; font-weight:600">${window.escapeHTML(item?.details || 'N/A')}</div>
            </div>
            <div style="display:grid; grid-template-columns: 150px 1fr; gap: 10px; border-bottom: 1px solid var(--border); padding: 10px 0;">
                <div style="font-weight:bold; color:var(--muted)">SENDER</div>
                <div style="color:var(--accent)">${window.escapeHTML(item?.sender || 'N/A')}</div>
            </div>
            <div style="display:grid; grid-template-columns: 150px 1fr; gap: 10px; border-bottom: 1px solid var(--border); padding: 10px 0;">
                <div style="font-weight:bold; color:var(--muted)">SOURCE IP</div>
                <div style="font-family:monospace">${window.escapeHTML(item?.ip || 'N/A')}</div>
            </div>
            <div style="display:grid; grid-template-columns: 150px 1fr; gap: 10px; border-bottom: 1px solid var(--border); padding: 10px 0;">
                <div style="font-weight:bold; color:var(--muted)">SAFETY SCORE</div>
                <div style="font-weight:900; color:var(--accent)">${Math.round(item?.score || 0)}%</div>
            </div>
            <div style="margin-top:20px; background:rgba(0,0,0,0.2); padding:15px; border-radius:10px; border:1px solid var(--border)">
                <div style="font-weight:bold; color:var(--accent); font-size:0.7rem; margin-bottom:10px; text-transform:uppercase">Forensic Digital Fingerprints (SHA256)</div>
                <div style="display:flex; flex-direction:column; gap:8px; font-family:monospace; font-size:0.7rem;">
                    <div style="display:flex; justify-content:space-between;"><span>Body Hash:</span> <span style="color:var(--muted)">${data.bodyHash || 'NOT CALCULATED'}</span></div>
                    ${(data.attachments || []).map(a => `<div style="display:flex; justify-content:space-between;"><span>File: ${window.escapeHTML(a.name)}</span> <span style="color:var(--muted)">${a.hash}</span></div>`).join('')}
                </div>
            </div>
            <div style="margin-top:20px; display:flex; justify-content:center;">
                <button id="scan-virus-btn" class="btn-ui success" style="width:250px; height:40px; font-size:0.8rem;">
                    <svg style="width:18px;height:18px;margin-right:8px;fill:currentColor" viewBox="0 0 24 24"><path d="M12,2L4.5,20.29L5.21,21L12,18L18.79,21L19.5,20.29L12,2Z" /></svg>
                    SCAN FOR VIRUSES (CLOUD)
                </button>
            </div>
        `;

        document.getElementById('scan-virus-btn').onclick = async () => {
            window.showNotification("INITIATING DEEP VIRUS SCAN...");
            const res = await api.scanForViruses(id);
            if (res.success) {
                window.showNotification("SCAN COMPLETE: DATA LOGGED TO FORENSICS");
                showForensics(id); // Refresh
            } else {
                window.showNotification("SCAN FAILED: " + res.error, true);
            }
        };

        modal.style.display = 'flex';
    };

    document.getElementById('forensic-close').onclick = () => modal.style.display = 'none';
    document.getElementById('forensic-rfc').onclick = async () => {
        const data = await api.getForensics(currentId);
        const win = new window.api.BrowserWindow({ width: 800, height: 600 }); // Note: This needs access to Electron or a new IPC
        // For simplicity, let's just log it for now or use a custom modal
        console.log(data.fullHeaders);
    };
})();
