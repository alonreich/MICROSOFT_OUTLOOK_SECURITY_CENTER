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

    const data = await window.securityApi.getForensics({
        entryId: item.entryId || item.id,
        originalEntryId: item.originalEntryId,
        fingerprint: item.fingerprint
    });
    
    let rawHeaders = (data && data.fullHeaders && data.fullHeaders !== 'N/A' && data.fullHeaders !== 'Unavailable')
        ? data.fullHeaders
        : (item.fullHeaders || item.headers || '');

    if (rawHeaders && typeof rawHeaders === 'string') {
        const clean = rawHeaders.trim();
        if (/^[A-Za-z0-9+/]+={0,2}$/.test(clean) && clean.length % 4 === 0 && !clean.includes(' ')) {
            try {
                const dec = atob(clean);
                if (dec && dec.length > 0) rawHeaders = dec;
            } catch {}
        }
    }

    if (!rawHeaders || rawHeaders.trim().length === 0) {
        rawHeaders = [
            `X-Delivery-Context: Internal / MAPI Store (No SMTP Transport Headers)`,
            `Message-ID: <${item.fingerprint || item.entryId || item.id}@deskguard.local>`,
            `Date: ${item.date ? item.date + ' ' + (item.time || '') : (item.timestamp || new Date().toUTCString())}`,
            `From: ${item.sender || item.from || 'Unknown'}`,
            `To: ${item.to || item.recipient || 'Internal User'}`,
            item.cc ? `CC: ${item.cc}` : null,
            `Subject: ${item.subject || item.details || '(No Subject)'}`,
            `X-Verdict: ${item.verdict || 'Safe'}`,
            `X-Integrity-Score: ${item.score !== undefined ? item.score : 100}%`
        ].filter(Boolean).join('\n');
    }

    let headersHtml = '';
    const lines = rawHeaders.split('\n');
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

    const safeSubject = window.escapeHtml(item.subject || item.details || 'No Subject');
    const safeSender = window.escapeHtml(item.sender || 'Unknown');
    const safeVerdict = window.escapeHtml(item.verdict || 'Safe');
    const safeTier = window.escapeHtml(item.tier || 'Standard Heuristics');
    const verdictClass = (item.verdict || '').toLowerCase().includes('malicious') ? 'badge-malicious' : ((item.verdict || '').toLowerCase().includes('spam') ? 'badge-spam' : 'badge-safe');

    body.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 16px; height: 100%;">
            <div style="background: var(--surface-low); padding: 20px; border: 1px solid var(--border); border-radius: 8px;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                    <h3 style="margin: 0; color: var(--accent); font-size: 0.9rem; text-transform: uppercase; letter-spacing: 1px;">Security Incident Profile</h3>
                    <button id="forensic-btn-open-outlook" class="primary" style="font-size: 0.75rem; padding: 5px 12px;">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
                        Open in Outlook
                    </button>
                </div>
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
        let bText = (data && data.body && data.body !== 'N/A' && data.body !== 'Unavailable')
            ? data.body
            : (item.body || '');
        if (bText && typeof bText === 'string') {
            const clean = bText.trim();
            if (/^[A-Za-z0-9+/]+={0,2}$/.test(clean) && clean.length % 4 === 0 && !clean.includes(' ')) {
                try {
                    const dec = atob(clean);
                    if (dec && dec.length > 0) bText = dec;
                } catch {}
            }
        }
        bodyTabEl.textContent = (bText && bText.trim().length > 0) ? bText : '(No message body content)';
    }

    const fOpenOutlook = document.getElementById('forensic-btn-open-outlook');
    if (fOpenOutlook) {
        fOpenOutlook.onclick = async () => {
            fOpenOutlook.disabled = true;
            const origText = fOpenOutlook.innerHTML;
            fOpenOutlook.innerHTML = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Launching...`;
            try {
                await window.securityApi.openEmail({
                    entryId: item.entryId || item.id,
                    storeId: item.storeId || ''
                });
            } catch (e) {}
            fOpenOutlook.disabled = false;
            fOpenOutlook.innerHTML = origText;
        };
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
