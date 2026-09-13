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

// --- HIGH-PERFORMANCE VIRTUAL VIEWPORT ENGINE ---
class VirtualList {
    constructor(options) {
        this.container = options.container;
        this.rowHeight = options.rowHeight || 44;
        this.buffer = options.buffer !== undefined ? options.buffer : 5;
        this.createRow = options.createRow;
        this.updateRow = options.updateRow;
        this.emptyRenderer = options.emptyRenderer;
        
        this.items = [];
        this.nodePool = [];
        this.activeNodes = new Map(); // index -> HTMLElement
        this.rafPending = false;
        this.scrollTop = 0;
        
        this.initDOM();
        this.initScrollListener();
    }

    initDOM() {
        this.container.innerHTML = '';
        this.container.style.position = 'relative';
        this.container.style.overflowY = 'auto';
        
        // Spacer defines native scroll height
        this.spacer = document.createElement('div');
        this.spacer.className = 'virtual-spacer';
        this.spacer.style.width = '1px';
        this.spacer.style.pointerEvents = 'none';
        this.spacer.style.opacity = '0';
        this.spacer.style.position = 'absolute';
        this.spacer.style.top = '0';
        this.spacer.style.left = '0';
        
        // Content layer containing visible absolute-positioned rows
        this.contentLayer = document.createElement('div');
        this.contentLayer.className = 'virtual-content';
        this.contentLayer.style.position = 'absolute';
        this.contentLayer.style.top = '0';
        this.contentLayer.style.left = '0';
        this.contentLayer.style.width = '100%';
        this.contentLayer.style.pointerEvents = 'auto';
        this.contentLayer.style.willChange = 'transform';
        
        this.container.appendChild(this.spacer);
        this.container.appendChild(this.contentLayer);
    }

    initScrollListener() {
        this.onScroll = () => {
            this.scrollTop = this.container.scrollTop;
            if (!this.rafPending) {
                this.rafPending = true;
                requestAnimationFrame(() => {
                    this.rafPending = false;
                    this.renderSlice();
                });
            }
        };
        this.container.addEventListener('scroll', this.onScroll, { passive: true });
    }

    setItems(items, resetScroll = false) {
        this.items = items || [];
        if (resetScroll) {
            this.container.scrollTop = 0;
            this.scrollTop = 0;
        } else {
            this.scrollTop = this.container.scrollTop;
        }

        const totalHeight = this.items.length * this.rowHeight;
        this.spacer.style.height = `${totalHeight}px`;

        if (this.items.length === 0) {
            for (const [idx, node] of this.activeNodes.entries()) {
                node.style.display = 'none';
                this.nodePool.push(node);
            }
            this.activeNodes.clear();
            if (this.emptyRenderer) {
                this.contentLayer.innerHTML = this.emptyRenderer();
            } else {
                this.contentLayer.innerHTML = '';
            }
            return;
        }

        // Clean out empty state if present
        if (this.contentLayer.querySelector('.empty-state')) {
            this.contentLayer.innerHTML = '';
            this.nodePool = [];
            this.activeNodes.clear();
        }

        this.renderSlice();
    }

    renderSlice() {
        if (this.items.length === 0) return;

        const viewportHeight = this.container.clientHeight || 500;
        const scrollTop = this.scrollTop;
        const total = this.items.length;

        const startIndex = Math.max(0, Math.floor(scrollTop / this.rowHeight) - this.buffer);
        const endIndex = Math.min(total, Math.ceil((scrollTop + viewportHeight) / this.rowHeight) + this.buffer);

        const newActiveNodes = new Map();
        const neededIndexes = [];

        // Check which needed indexes already have nodes
        for (let i = startIndex; i < endIndex; i++) {
            if (this.activeNodes.has(i)) {
                newActiveNodes.set(i, this.activeNodes.get(i));
            } else {
                neededIndexes.push(i);
            }
        }

        // Evict nodes no longer in viewport slice back to pool
        for (const [idx, node] of this.activeNodes.entries()) {
            if (idx < startIndex || idx >= endIndex) {
                node.style.display = 'none';
                this.nodePool.push(node);
            }
        }

        // Acquire nodes for needed indexes from pool or instantiate new
        for (const i of neededIndexes) {
            let node;
            if (this.nodePool.length > 0) {
                node = this.nodePool.pop();
                node.style.display = '';
            } else {
                node = this.createRow();
                this.contentLayer.appendChild(node);
            }
            newActiveNodes.set(i, node);
        }

        this.activeNodes = newActiveNodes;

        // Diff and position all active nodes
        for (const [i, node] of this.activeNodes.entries()) {
            const y = i * this.rowHeight;
            node.style.transform = `translate3d(0, ${y}px, 0)`;
            this.updateRow(node, this.items[i], i);
        }
    }

    scrollToIndex(index) {
        if (index < 0 || index >= this.items.length) return;
        const targetTop = index * this.rowHeight;
        const viewportHeight = this.container.clientHeight || 500;
        const currentScroll = this.container.scrollTop;

        if (targetTop < currentScroll) {
            this.container.scrollTop = targetTop;
        } else if (targetTop + this.rowHeight > currentScroll + viewportHeight) {
            this.container.scrollTop = targetTop + this.rowHeight - viewportHeight;
        }
    }

    refreshCurrentSlice() {
        for (const [i, node] of this.activeNodes.entries()) {
            if (this.items[i]) {
                this.updateRow(node, this.items[i], i);
            }
        }
    }

    destroy() {
        if (this.onScroll) {
            this.container.removeEventListener('scroll', this.onScroll);
        }
        this.activeNodes.clear();
        this.nodePool = [];
    }
}
window.VirtualList = VirtualList;

// --- INCIDENT LIST CONTROLLER ---
window.incidentVirtualList = null;
window.currentRenderedCategory = null;

function createIncidentRowElement() {
    const div = document.createElement('div');
    div.className = 'list-item incident-grid virtual-row';

    const timeEl = document.createElement('div');
    timeEl.style.fontFamily = 'monospace';
    timeEl.style.fontSize = '0.7rem';

    const subjEl = document.createElement('div');
    subjEl.style.fontWeight = '600';
    subjEl.style.color = 'var(--text)';
    subjEl.style.overflow = 'hidden';
    subjEl.style.textOverflow = 'ellipsis';
    subjEl.style.whiteSpace = 'nowrap';

    const senderEl = document.createElement('div');
    senderEl.style.overflow = 'hidden';
    senderEl.style.textOverflow = 'ellipsis';
    senderEl.style.whiteSpace = 'nowrap';
    senderEl.style.fontSize = '0.75rem';

    const badgeCol = document.createElement('div');
    const badgeSpan = document.createElement('span');
    badgeSpan.className = 'badge';
    badgeCol.appendChild(badgeSpan);

    div.appendChild(timeEl);
    div.appendChild(subjEl);
    div.appendChild(senderEl);
    div.appendChild(badgeCol);

    div._cells = { timeEl, subjEl, senderEl, badgeSpan };
    return div;
}

function updateIncidentRowElement(div, item, index) {
    const cells = div._cells;
    const ts = new Date(item.timestamp).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
    const verdict = item.verdict || 'Pending';
    const vLower = verdict.toLowerCase();
    const verdictClass = vLower.includes('malicious') ? 'badge-malicious' : (vLower.includes('spam') ? 'badge-spam' : 'badge-safe');

    if (cells.timeEl.textContent !== ts) cells.timeEl.textContent = ts;
    const subjText = item.subject || 'No Subject';
    if (cells.subjEl.textContent !== subjText) cells.subjEl.textContent = subjText;
    const senderText = item.sender || 'Unknown';
    if (cells.senderEl.textContent !== senderText) cells.senderEl.textContent = senderText;
    if (cells.badgeSpan.textContent !== verdict) cells.badgeSpan.textContent = verdict;
    const targetBadgeClass = `badge ${verdictClass}`;
    if (cells.badgeSpan.className !== targetBadgeClass) cells.badgeSpan.className = targetBadgeClass;

    const isSelected = window.AppState.selectedId === (item.entryId || item.fingerprint);
    if (isSelected) {
        if (!div.classList.contains('active')) div.classList.add('active');
    } else {
        if (div.classList.contains('active')) div.classList.remove('active');
    }

    div.onclick = () => window.selectIncident(item);
}

window.renderList = (items, category, forceResetScroll = false) => {
    const container = document.getElementById('security-list');
    if (!container) return;

    // Highlight Active Stat Card
    document.querySelectorAll('.stat-card').forEach(c => c.classList.remove('active'));
    const activeCard = document.getElementById(`card-${category}`);
    if (activeCard) activeCard.classList.add('active');

    if (!window.incidentVirtualList) {
        window.incidentVirtualList = new VirtualList({
            container: container,
            rowHeight: 44,
            buffer: 5,
            createRow: createIncidentRowElement,
            updateRow: updateIncidentRowElement,
            emptyRenderer: () => `
                <div class="empty-state">
                    <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="var(--border)" stroke-width="1.5"><circle cx="12" cy="12" r="10"></circle><path d="M12 8v4l3 3"></path></svg>
                    <h3>Clean Environment</h3>
                    <p>No incidents detected in this category.</p>
                </div>`
        });
    }

    const categoryChanged = (window.currentRenderedCategory !== category);
    window.currentRenderedCategory = category;
    const shouldReset = forceResetScroll || categoryChanged;

    // In-memory sort
    const sorted = [...(items || [])];
    const { key, asc } = window.sortOrder;
    sorted.sort((a, b) => {
        let valA = a[key] || '';
        let valB = b[key] || '';
        if (key === 'timestamp') {
            valA = new Date(valA).getTime() || 0;
            valB = new Date(valB).getTime() || 0;
        } else {
            valA = String(valA).toLowerCase();
            valB = String(valB).toLowerCase();
        }
        if (valA < valB) return asc ? -1 : 1;
        if (valA > valB) return asc ? 1 : -1;
        return 0;
    });

    window.incidentVirtualList.setItems(sorted, shouldReset);
};

window.selectIncident = async (item) => {
    window.AppState.selectedId = item.entryId || item.fingerprint;
    if (window.incidentVirtualList) {
        window.incidentVirtualList.refreshCurrentSlice();
    }
    
    const panel = document.getElementById('detail-panel');
    if (!panel) return;

    const verdict = item.verdict || 'Pending';
    const isQuarantined = verdict.toLowerCase().includes('malicious') || verdict.toLowerCase().includes('spam') || verdict.toLowerCase().includes('quarantin');
    const isSafe = verdict.toLowerCase().includes('safe');
    
    const safeSubject = window.escapeHtml(item.subject || 'No Subject');
    const safeSender = window.escapeHtml(item.sender || 'Unknown');
    const safeTier = window.escapeHtml(item.tier || 'Standard Heuristics');
    const safeIp = (!item.ip || item.ip === 'N/A') ? 'Internal Network' : window.escapeHtml(item.ip);
    const safeVerdict = window.escapeHtml(verdict);
    const verdictClass = verdict.toLowerCase().includes('malicious') ? 'badge-malicious' : (verdict.toLowerCase().includes('spam') ? 'badge-spam' : 'badge-safe');

    let actionButtons = `
        <button class="primary" id="btn-open-forensics">Full Forensics</button>
        <button id="btn-vt-scan">VirusTotal Audit</button>
    `;

    if (isQuarantined || window.AppState.currentCategory === 'malicious' || window.AppState.currentCategory === 'spam' || window.AppState.currentCategory === 'suspicious') {
        actionButtons += `<button class="success" id="btn-release">Release to Inbox</button>`;
    }
    if (isSafe || window.AppState.currentCategory === 'safe' || window.AppState.currentCategory === 'suspicious') {
        actionButtons += `<button class="warning" id="btn-quarantine">Quarantine</button>`;
    }
    actionButtons += `<button class="danger" id="btn-delete">Delete Item</button>`;

    panel.innerHTML = `
        <div class="detail-header">
            <h2 style="margin: 0; font-size: 1.1rem; font-weight: 500;">${safeSubject}</h2>
            <div style="margin-top: 8px;"><span class="badge ${verdictClass}">${safeVerdict}</span></div>
        </div>
        <div class="detail-content">
            <div class="detail-meta">
                <label>Sender</label><div style="word-break: break-all;">${safeSender}</div>
                <label>Received</label><div>${window.escapeHtml(new Date(item.timestamp).toLocaleString())}</div>
                <label>IP Origin</label><div style="font-family: monospace; color: ${!item.ip || item.ip === 'N/A' ? 'var(--muted)' : 'var(--accent)'}">${safeIp}</div>
                <label>Score</label><div style="font-weight: 700; color: ${item.score < 50 ? 'var(--danger)' : 'var(--ok)'}">${parseInt(item.score) || 0}% Integrity</div>
                <label>Analysis</label><div style="font-size: 0.75rem; color: var(--muted);">${safeTier}</div>
            </div>
            <div style="margin-bottom: 12px; display: flex; flex-wrap: wrap; gap: 8px;">
                ${actionButtons}
            </div>
            <div class="detail-body" id="detail-body-content">Loading message content...</div>
        </div>
    `;

    // Bind event handlers securely via JS closures
    const btnForensics = document.getElementById('btn-open-forensics');
    if (btnForensics) btnForensics.onclick = () => window.openForensicsModal(item.entryId || item.fingerprint);
    const btnVt = document.getElementById('btn-vt-scan');
    if (btnVt) btnVt.onclick = () => window.cloudScanIncident(item.entryId || '');
    const btnRel = document.getElementById('btn-release');
    if (btnRel) btnRel.onclick = () => window.releaseIncident(item.entryId || '', item.fingerprint || '');
    const btnQuar = document.getElementById('btn-quarantine');
    if (btnQuar) btnQuar.onclick = () => window.quarantineIncident(item.entryId || '', item.fingerprint || '');
    const btnDel = document.getElementById('btn-delete');
    if (btnDel) btnDel.onclick = () => window.deleteIncident(item.entryId || '');

    const forensics = await window.securityApi.getForensics(item.entryId || item.fingerprint);
    const bodyEl = document.getElementById('detail-body-content');
    if (bodyEl) {
        bodyEl.textContent = (forensics && forensics.body && forensics.body !== 'N/A') ? forensics.body : "No body content available for this item (Internal/System Message).";
    }
};

window.handleListKeyNavigation = (e) => {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    const cat = window.AppState.currentCategory;
    const items = (window.incidentVirtualList && window.incidentVirtualList.items) || window.AppState.stats[cat] || [];
    if (items.length === 0) return;
    
    const currentIndex = items.findIndex(i => (i.entryId || i.fingerprint) === window.AppState.selectedId);
    let nextIndex = 0;
    
    if (currentIndex === -1) {
        nextIndex = 0;
    } else if (e.key === 'ArrowDown') {
        nextIndex = Math.min(items.length - 1, currentIndex + 1);
    } else if (e.key === 'ArrowUp') {
        nextIndex = Math.max(0, currentIndex - 1);
    }
    
    if (nextIndex !== currentIndex && items[nextIndex]) {
        e.preventDefault();
        window.selectIncident(items[nextIndex]);
        if (window.incidentVirtualList) {
            window.incidentVirtualList.scrollToIndex(nextIndex);
        }
    }
};

window.openForensicsModal = (id) => {
    const item = Object.values(window.AppState.stats).flat().find(i => (i.entryId || i.fingerprint) === id);
    if (item) window.openForensics(item);
};

window.releaseIncident = async (entryId, fingerprint) => {
    const btn = document.getElementById('btn-release');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Releasing...`;
    }
    try {
        const res = await window.securityApi.releaseEmail({ entryId, fingerprint, targetFolder: 6 });
        if (res && res.success) {
            const panel = document.getElementById('detail-panel');
            if (panel) {
                panel.innerHTML = `
                    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--ok); text-align: center; padding: 20px;">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-bottom: 12px;"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                        <h3 style="margin: 0 0 8px 0; color: var(--text);">Email Restored to Inbox</h3>
                        <p style="color: var(--muted); margin: 0; font-size: 0.85rem;">The message was successfully moved to Outlook Inbox and whitelisted against re-quarantine.</p>
                    </div>`;
            }
        } else {
            alert("Release failed: " + ((res && res.error) || "Outlook could not locate or move the item."));
            if (btn) {
                btn.disabled = false;
                btn.textContent = "Release to Inbox";
            }
        }
    } catch (err) {
        alert("Release error: " + err.message);
        if (btn) {
            btn.disabled = false;
            btn.textContent = "Release to Inbox";
        }
    }
};

window.quarantineIncident = async (entryId, fingerprint) => {
    const btn = document.getElementById('btn-quarantine');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Quarantining...`;
    }
    try {
        const res = await window.securityApi.quarantineEmail({ entryId, fingerprint, targetFolder: 23 });
        if (res && res.success) {
            const panel = document.getElementById('detail-panel');
            if (panel) {
                panel.innerHTML = `
                    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--warn); text-align: center; padding: 20px;">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-bottom: 12px;"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
                        <h3 style="margin: 0 0 8px 0; color: var(--text);">Email Quarantined</h3>
                        <p style="color: var(--muted); margin: 0; font-size: 0.85rem;">The message was successfully moved to Outlook Junk Email.</p>
                    </div>`;
            }
        } else {
            alert("Quarantine failed: " + ((res && res.error) || "Outlook could not locate or move the item."));
            if (btn) {
                btn.disabled = false;
                btn.textContent = "Quarantine";
            }
        }
    } catch (err) {
        alert("Quarantine error: " + err.message);
        if (btn) {
            btn.disabled = false;
            btn.textContent = "Quarantine";
        }
    }
};

window.cloudScanIncident = async (entryId) => {
    const btn = document.getElementById('btn-vt-scan');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Auditing...`;
    }
    try {
        const res = await window.securityApi.scanForViruses(entryId);
        if (btn) {
            btn.disabled = false;
            btn.textContent = "VirusTotal Audit";
        }
        if (res && res.success && res.data) {
            window.showVirusTotalReport(res.data);
        } else {
            alert("VirusTotal audit notice: " + ((res && (res.error || (res.data && res.data.error))) || "No threat data available or item not found in database. Check your VirusTotal API key in Settings."));
        }
    } catch (err) {
        if (btn) {
            btn.disabled = false;
            btn.textContent = "VirusTotal Audit";
        }
        alert("VirusTotal query error: " + err.message);
    }
};

window.showVirusTotalReport = (rep) => {
    const body = document.getElementById('virustotal-body');
    if (!body || !rep) return;
    
    let itemsHtml = '';
    if (rep.items && rep.items.length > 0) {
        itemsHtml = rep.items.map(it => {
            const malCount = parseInt(it.malicious) || 0;
            const suspCount = parseInt(it.suspicious) || 0;
            const isBad = (malCount > 0) || (suspCount > 0);
            const statusColor = isBad ? 'var(--danger)' : ((it.status && it.status.includes('Clean')) ? 'var(--ok)' : 'var(--muted)');
            const safeName = window.escapeHtml(it.name || 'Artifact');
            const safeType = window.escapeHtml(it.type || 'file');
            const safeStatus = window.escapeHtml(it.status || 'Audited');
            const safeHash = window.escapeHtml(it.hash || 'N/A');
            return `
                <div style="background: var(--surface-low); border: 1px solid var(--border); border-radius: 6px; padding: 12px; margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                        <span style="font-weight: 600; font-size: 0.85rem; color: var(--text);">${safeName} (${safeType})</span>
                        <span style="font-weight: 700; font-size: 0.75rem; color: ${statusColor};">${safeStatus}</span>
                    </div>
                    <div style="font-family: monospace; font-size: 0.7rem; color: var(--muted); word-break: break-all;">
                        SHA256: ${safeHash}
                    </div>
                    ${(it.malicious !== undefined) ? `
                    <div style="margin-top: 6px; font-size: 0.75rem; display: flex; gap: 12px;">
                        <span style="color: var(--danger);">Malicious: ${malCount}</span>
                        <span style="color: var(--warn);">Suspicious: ${suspCount}</span>
                    </div>` : ''}
                </div>
            `;
        }).join('');
    } else {
        itemsHtml = '<div style="color: var(--muted); padding: 12px;">No hashable artifacts extracted for this message.</div>';
    }

    const safeSubject = window.escapeHtml(rep.subject || 'Incident Audit Report');
    const safeStatus = window.escapeHtml(rep.status || 'Report Ready');
    const safeSender = window.escapeHtml(rep.sender || 'Unknown');
    const safeIp = window.escapeHtml(rep.ip || 'Internal');
    const threatsCount = parseInt(rep.threatsFound) || 0;

    body.innerHTML = `
        <div style="margin-bottom: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                <h3 style="margin: 0; font-size: 1rem; color: var(--text);">${safeSubject}</h3>
                <span class="badge ${threatsCount > 0 ? 'badge-malicious' : 'badge-safe'}">${safeStatus}</span>
            </div>
            <div style="font-size: 0.75rem; color: var(--muted); display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
                <div>Sender: <span style="color: var(--text);">${safeSender}</span></div>
                <div>IP Origin: <span style="font-family: monospace; color: var(--accent);">${safeIp}</span></div>
                <div>API Integration: <span style="color: ${rep.vtQueried ? 'var(--ok)' : 'var(--warn)'};">${rep.vtQueried ? 'Connected to VirusTotal v3' : 'Local Hashes (No API Key)'}</span></div>
                <div>Threat Detections: <span style="font-weight: 700; color: ${threatsCount > 0 ? 'var(--danger)' : 'var(--ok)'};">${threatsCount}</span></div>
            </div>
        </div>
        <div style="margin-top: 16px;">
            <h4 style="margin: 0 0 10px 0; font-size: 0.8rem; text-transform: uppercase; color: var(--muted); letter-spacing: 0.5px;">Cryptographic Hashes & Signatures</h4>
            ${itemsHtml}
        </div>
    `;

    openModal('virustotal-modal');
};

window.verifyCurrentList = async () => {
    const btn = document.getElementById('verify-list-btn');
    const cat = window.AppState.currentCategory;
    const items = window.AppState.stats[cat] || [];
    if (items.length === 0) {
        alert("No items in " + cat + " to verify.");
        return;
    }
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Verifying...`;
    }
    try {
        const res = await window.securityApi.verifyExistence({ items, category: cat });
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"></path></svg> Verify List`;
        }
        if (res && res.removedCount > 0) {
            alert(`Verification complete: Pruned ${res.removedCount} missing item(s) from Outlook.`);
        } else {
            alert(`Verification complete: All ${items.length} items in ${cat} confirmed present in Outlook.`);
        }
    } catch (err) {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"></path></svg> Verify List`;
        }
        alert("Verification error: " + err.message);
    }
};

window.deleteIncident = (id) => {
    window.showConfirm("Delete Email", "This will permanently delete this email from Outlook. Proceed?", "DELETE", async () => {
        await window.securityApi.deleteEmail({ entryIds: [id] });
    });
};

window.sortOrder = { key: 'timestamp', asc: false };
window.sortList = (key) => {
    if (window.sortOrder.key === key) {
        window.sortOrder.asc = !window.sortOrder.asc;
    } else {
        window.sortOrder.key = key;
        window.sortOrder.asc = true;
    }

    // Update column indicators
    ['timestamp', 'subject', 'sender', 'verdict'].forEach(col => {
        const el = document.getElementById(`sort-${col}`);
        if (el) {
            if (col === window.sortOrder.key) {
                el.textContent = window.sortOrder.asc ? '▲' : '▼';
            } else {
                el.textContent = '';
            }
        }
    });

    const cat = window.AppState.currentCategory;
    const items = window.AppState.stats[cat] || [];
    window.renderList(items, cat, false);
};
