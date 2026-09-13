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
SecurityUI.Duplicates = (function() {
    const api = window.securityApi;
    let duplicateVirtualList = null;
    let duplicateItems = [];
    window.lastDuplicates = [];

    function createDuplicateRow() {
        const div = document.createElement('div');
        div.className = 'list-item duplicate-grid virtual-row';

        const timeEl = document.createElement('div');
        timeEl.style.fontSize = '0.75rem';
        timeEl.style.color = 'var(--muted)';
        timeEl.style.fontFamily = 'monospace';

        const subjEl = document.createElement('div');
        subjEl.style.fontWeight = '500';
        subjEl.style.whiteSpace = 'nowrap';
        subjEl.style.overflow = 'hidden';
        subjEl.style.textOverflow = 'ellipsis';
        subjEl.style.color = 'var(--text)';

        const sizeEl = document.createElement('div');
        sizeEl.style.fontSize = '0.75rem';
        sizeEl.style.color = 'var(--muted)';

        const folderEl = document.createElement('div');
        folderEl.style.fontSize = '0.75rem';
        folderEl.style.whiteSpace = 'nowrap';
        folderEl.style.overflow = 'hidden';
        folderEl.style.textOverflow = 'ellipsis';
        folderEl.style.color = 'var(--muted)';

        const survivorEl = document.createElement('div');
        survivorEl.style.fontSize = '0.75rem';
        survivorEl.style.whiteSpace = 'nowrap';
        survivorEl.style.overflow = 'hidden';
        survivorEl.style.textOverflow = 'ellipsis';
        survivorEl.style.color = 'var(--accent-green, #4ade80)';

        div.appendChild(timeEl);
        div.appendChild(subjEl);
        div.appendChild(sizeEl);
        div.appendChild(folderEl);
        div.appendChild(survivorEl);

        div._cells = { timeEl, subjEl, sizeEl, folderEl, survivorEl };
        return div;
    }

    function updateDuplicateRow(div, item) {
        const cells = div._cells;
        const ts = item.timestamp || '';
        if (cells.timeEl.textContent !== ts) cells.timeEl.textContent = ts;
        const subj = item.subject || 'No Subject';
        if (cells.subjEl.textContent !== subj) cells.subjEl.textContent = subj;
        const sz = item.size ? `${(item.size / 1024).toFixed(1)} KB` : '0 KB';
        if (cells.sizeEl.textContent !== sz) cells.sizeEl.textContent = sz;
        const fld = item.folder || '';
        if (cells.folderEl.textContent !== fld) cells.folderEl.textContent = fld;
        const survFld = item.survivorFolder || (item.survivor && item.survivor.folder) || 'Inbox';
        const survText = `${survFld} (Survivor Preserved)`;
        if (cells.survivorEl.textContent !== survText) cells.survivorEl.textContent = survText;
        div.setAttribute('data-id', item.entryId || '');
    }

    function render(data) {
        const view = document.getElementById('duplicate-view');
        if (!data || !view) return;

        if (data.status === 'Scanned' || data.status === 'Progress' || data.status === 'Found') {
            const scannedCount = data.scanned !== undefined ? data.scanned : (data.current || 0);
            const html = `
                <div style="padding: 40px; text-align: center;">
                    <div class="stat-value" style="color: var(--accent);">${scannedCount}</div>
                    <div class="stat-label">Total Items Scanned</div>
                    <div style="margin-top: 20px; font-size: 0.9rem; color: var(--muted); line-height: 1.6;">
                        ${data.store ? `Scanning: <b>${window.escapeHtml(data.store)}</b> / ${window.escapeHtml(data.currentFolder || data.folder || '')}<br>` : ''}
                        ${data.details ? `<span>${window.escapeHtml(data.details)}</span><br>` : ''}
                    </div>
                    <div style="margin-top: 30px;">
                        <button onclick="window.securityApi.pauseDuplicateScan()">Pause Discovery</button>
                    </div>
                </div>
            `;
            if (view.innerHTML !== html) view.innerHTML = html;
        } else if (data.status === 'Finished') {
            const items = data.items || [];
            duplicateItems = items;
            window.lastDuplicates = items.map(i => i.entryId);
            
            view.innerHTML = `
                <div class="section-header" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-shrink: 0;">
                    <h2 style="margin: 0; font-size: 1.1rem;">${items.length} Redundant Emails Found</h2>
                    <div style="display: flex; gap: 10px;">
                        <button class="primary" onclick="cleanDuplicatesSafe()" ${items.length === 0 ? 'disabled' : ''}>Move to Deleted Items (Safe)</button>
                        <button class="danger" onclick="cleanDuplicatesPurge()" ${items.length === 0 ? 'disabled' : ''}>Permanent Purge</button>
                    </div>
                </div>
                <div class="list-header duplicate-grid" style="flex-shrink: 0;">
                    <div>Date</div>
                    <div>Subject</div>
                    <div>Size</div>
                    <div>Duplicate Location</div>
                    <div>Survivor Status</div>
                </div>
                <div id="duplicate-list" class="list-container" style="flex: 1; min-height: 0; position: relative; overflow-y: auto;">
                </div>
            `;

            const container = document.getElementById('duplicate-list');
            if (container && window.VirtualList) {
                duplicateVirtualList = new window.VirtualList({
                    container: container,
                    rowHeight: 44,
                    buffer: 5,
                    createRow: createDuplicateRow,
                    updateRow: updateDuplicateRow,
                    emptyRenderer: () => `
                        <div class="empty-state">
                            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="var(--border)" stroke-width="1.5"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                            <h3>Clean Mailbox</h3>
                            <p>No redundant or duplicate email copies detected.</p>
                        </div>`
                });
                duplicateVirtualList.setItems(items, true);
            }
        } else if (data.status === 'Paused') {
            view.innerHTML = `
                <div style="padding: 40px; text-align: center;">
                    <h3>Discovery Paused</h3>
                    <p style="color: var(--muted); margin-bottom: 24px;">The duplicate discovery crawler has been temporarily suspended.</p>
                    <button class="primary" onclick="window.securityApi.resumeDuplicateScan()">Resume Discovery</button>
                </div>
            `;
        }
    }

    window.cleanDuplicatesSafe = async () => {
        if (!duplicateItems || duplicateItems.length === 0) return;
        window.showConfirm(
            "Move Duplicates to Deleted Items",
            `This will safely move ${duplicateItems.length} redundant email copies to Deleted Items. Pre-flight validation will ensure that original survivor emails remain untouched in their current folders. Proceed?`,
            null,
            async () => {
                const res = await api.deleteDuplicates({ items: duplicateItems, mode: 'safe' });
                const success = (res && res.successCount !== undefined) ? res.successCount : ((res && res.count) || 0);
                const skipped = (res && res.skippedCount) || 0;
                let msg = `Clean-up Complete:\n- Successfully isolated: ${success} redundant copies.`;
                if (skipped > 0) {
                    msg += `\n- Skipped (Protected): ${skipped} items (survivors missing or modified).`;
                }
                alert(msg);
                window.lastDuplicates = [];
                duplicateItems = [];
                if (duplicateVirtualList) duplicateVirtualList.setItems([], true);
                render({ status: 'Finished', items: [] });
            }
        );
    };

    window.cleanDuplicatesPurge = async () => {
        if (!duplicateItems || duplicateItems.length === 0) return;
        window.showConfirm(
            "Permanent Purge of Duplicates",
            `WARNING: This will permanently delete ${duplicateItems.length} redundant email copies from Outlook. Pre-flight validation will strictly verify that each original survivor exists before any duplicate is deleted. Type PURGE to confirm.`,
            "PURGE",
            async () => {
                const res = await api.deleteDuplicates({ items: duplicateItems, mode: 'purge' });
                const success = (res && res.successCount !== undefined) ? res.successCount : ((res && res.count) || 0);
                const skipped = (res && res.skippedCount) || 0;
                let msg = `Permanent Purge Complete:\n- Successfully deleted: ${success} redundant copies.`;
                if (skipped > 0) {
                    msg += `\n- Skipped (Protected): ${skipped} items (survivors missing or modified).`;
                }
                alert(msg);
                window.lastDuplicates = [];
                duplicateItems = [];
                if (duplicateVirtualList) duplicateVirtualList.setItems([], true);
                render({ status: 'Finished', items: [] });
            }
        );
    };

    // Backward-compatibility alias
    window.deleteDuplicates = window.cleanDuplicatesSafe;

    api.onDuplicateUpdate(data => render(data));

    return {};
})();
