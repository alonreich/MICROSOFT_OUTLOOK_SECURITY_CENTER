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

window.renderForensics = async (item) => {
    const body = document.getElementById('forensics-body');
    if (!body || !item) return;
    
    body.innerHTML = '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: var(--muted);">Deconstructing forensic snapshots...</div>';

    const data = await window.securityApi.getForensics(item.entryId || item.fingerprint);
    
    let headersHtml = '';
    if (data && data.fullHeaders && data.fullHeaders !== 'N/A') {
        const lines = data.fullHeaders.split('\n');
        headersHtml = lines.map(line => {
            const parts = line.split(':');
            if (parts.length > 1) {
                const key = parts[0].trim();
                const val = parts.slice(1).join(':').trim();
                const keyEsc = window.escapeHtml(key);
                const valEsc = window.escapeHtml(val);
                return `<div style="display: flex; gap: 10px; margin-bottom: 4px; border-bottom: 1px solid #222; padding: 4px 0;">
                    <span style="color: var(--accent); font-weight: 600; min-width: 150px; font-family: monospace; font-size: 0.75rem;">${keyEsc}:</span>
                    <span style="color: var(--text); font-family: monospace; font-size: 0.75rem; word-break: break-all;">${valEsc}</span>
                </div>`;
            }
            return `<div style="color: var(--muted); font-size: 0.75rem;">${window.escapeHtml(line)}</div>`;
        }).join('');
    } else {
        headersHtml = '<div style="color: var(--muted); padding: 20px;">No Header Data Available</div>';
    }

    const safeSubject = window.escapeHtml(item.subject || 'No Subject');
    const safeSender = window.escapeHtml(item.sender || 'Unknown');
    const safeVerdict = window.escapeHtml(item.verdict || 'Safe');
    const safeTier = window.escapeHtml(item.tier || 'Standard Heuristics');
    const verdictClass = (item.verdict || '').toLowerCase().includes('malicious') ? 'badge-malicious' : ((item.verdict || '').toLowerCase().includes('spam') ? 'badge-spam' : 'badge-safe');

    body.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 16px; height: 100%;">
            <div style="background: var(--surface-low); padding: 20px; border: 1px solid var(--border); border-radius: 8px;">
                <h3 style="margin-top: 0; margin-bottom: 16px; color: var(--accent); font-size: 0.9rem; text-transform: uppercase; letter-spacing: 1px;">Security Incident Profile</h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
                    <div>
                        <label style="font-size: 0.65rem; color: var(--muted); font-weight: 700;">SUBJECT</label>
                        <div style="font-weight: 600; font-size: 0.9rem; color: var(--text);">${safeSubject}</div>
                    </div>
                    <div>
                        <label style="font-size: 0.65rem; color: var(--muted); font-weight: 700;">SENDER ORIGIN</label>
                        <div style="font-weight: 600; font-size: 0.9rem; color: var(--text);">${safeSender}</div>
                    </div>
                    <div>
                        <label style="font-size: 0.65rem; color: var(--muted); font-weight: 700;">VERDICT</label>
                        <div><span class="badge ${verdictClass}">${safeVerdict}</span></div>
                    </div>
                    <div>
                        <label style="font-size: 0.65rem; color: var(--muted); font-weight: 700;">DETECTION SOURCE</label>
                        <div style="font-weight: 600; font-family: monospace; font-size: 0.8rem; color: var(--muted);">${safeTier}</div>
                    </div>
                </div>
            </div>

            <div style="flex: 1; display: flex; flex-direction: column; min-height: 0; background: var(--surface-low); border: 1px solid var(--border); border-radius: 8px; overflow: hidden;">
                <div style="display: flex; border-bottom: 1px solid var(--border); background: var(--surface-mid);">
                    <div id="tab-btn-headers" onclick="switchForensicTab('headers')" style="padding: 12px 24px; cursor: pointer; border-bottom: 2px solid var(--accent); color: var(--accent); font-weight: 600; font-size: 0.75rem; letter-spacing: 0.5px;">TRANSPORT HEADERS</div>
                    <div id="tab-btn-body" onclick="switchForensicTab('body')" style="padding: 12px 24px; cursor: pointer; color: var(--muted); font-size: 0.75rem; letter-spacing: 0.5px;">MESSAGE CONTENT</div>
                </div>
                <div id="forensic-tab-headers" style="flex: 1; overflow: auto; padding: 16px; background: #080a11;">
                    ${headersHtml}
                </div>
                <div id="forensic-tab-body" style="display: none; flex: 1; overflow: auto; padding: 20px; background: #080a11; font-family: 'Consolas', monospace; font-size: 0.8rem; color: #a0a0a0; white-space: pre-wrap; line-height: 1.6;"></div>
            </div>
        </div>
    `;

    // Securely populate plain text message body using textContent
    const bodyTabEl = document.getElementById('forensic-tab-body');
    if (bodyTabEl) {
        bodyTabEl.textContent = (data && data.body && data.body !== 'N/A') ? data.body : 'No Body Data Extracted';
    }
};

window.switchForensicTab = (tab) => {
    const hBtn = document.getElementById('tab-btn-headers');
    const bBtn = document.getElementById('tab-btn-body');
    const hTab = document.getElementById('forensic-tab-headers');
    const bTab = document.getElementById('forensic-tab-body');
    if (!hBtn || !bBtn || !hTab || !bTab) return;

    if (tab === 'headers') {
        hBtn.style.borderBottom = '2px solid var(--accent)';
        hBtn.style.color = 'var(--accent)';
        bBtn.style.borderBottom = 'none';
        bBtn.style.color = 'var(--muted)';
        hTab.style.display = 'block';
        bTab.style.display = 'none';
    } else {
        bBtn.style.borderBottom = '2px solid var(--accent)';
        bBtn.style.color = 'var(--accent)';
        hBtn.style.borderBottom = 'none';
        hBtn.style.color = 'var(--muted)';
        bTab.style.display = 'block';
        hTab.style.display = 'none';
    }
};
