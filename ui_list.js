(function() {
    const api = window.securityApi;
    const store = window.AppStore;

    window.showConfirmModal = function(title, message, onConfirm) {
        const modal = document.getElementById('custom-confirm-modal');
        document.getElementById('confirm-modal-title').textContent = title;
        document.getElementById('confirm-modal-msg').textContent = message;
        
        const yesBtn = document.getElementById('confirm-modal-yes');
        const noBtn = document.getElementById('confirm-modal-no');
        
        yesBtn.onclick = () => { modal.style.display = 'none'; onConfirm(); };
        noBtn.onclick = () => { modal.style.display = 'none'; };
        
        modal.style.display = 'flex';
    }

    async function deleteEmails() {
        const ids = Array.from(store.selectedIds);
        if (ids.length === 0) return;
        
        window.showConfirmModal(
            'PERMANENT DELETION', 
            `Are you sure you want to permanently delete ${ids.length} selected email(s)? This action cannot be undone.`, 
            () => {
                api.deleteEmail({ entryIds: ids });
                const currentCat = store.currentCategory;
                if (store.stats[currentCat]) {
                    store.stats[currentCat] = store.stats[currentCat].filter(i => !ids.includes(i.entryId));
                    store.updateCounterUI();
                }
                store.selectedIds.clear();
                window.renderList();
                window.showNotification(`Successfully deleted ${ids.length} item(s).`);
            }
        );
    }

    window.renderList = function() {
        const cat = store.currentCategory;
        let allItems = (store.stats[cat] || []);

        // Apply Global Sorting
        if (store.sortConfig.key) {
            allItems.sort((a, b) => {
                let vA = a[store.sortConfig.key] || "";
                let vB = b[store.sortConfig.key] || "";
                
                if (store.sortConfig.key === 'score') {
                    vA = parseFloat(vA) || 0;
                    vB = parseFloat(vB) || 0;
                } else if (store.sortConfig.key === 'subject') {
                    vA = (a.subject || a.details || "").toLowerCase();
                    vB = (b.subject || b.details || "").toLowerCase();
                } else {
                    vA = String(vA).toLowerCase();
                    vB = String(vB).toLowerCase();
                }

                if (vA < vB) return store.sortConfig.direction === 'asc' ? -1 : 1;
                if (vA > vB) return store.sortConfig.direction === 'asc' ? 1 : -1;
                return 0;
            });
        }

        const totalPages = Math.max(1, Math.ceil(allItems.length / store.pageSize));
        if (store.currentPage > totalPages) store.currentPage = totalPages;
        const start = (store.currentPage - 1) * store.pageSize;
        const items = allItems.slice(start, start + store.pageSize);
        
        document.getElementById('pg-info').textContent = `Page ${store.currentPage} of ${totalPages} (${allItems.length} items)`;
        document.getElementById('pg-first').disabled = store.currentPage === 1;
        document.getElementById('pg-prev').disabled = store.currentPage === 1;
        document.getElementById('pg-next').disabled = store.currentPage === totalPages;
        document.getElementById('pg-last').disabled = store.currentPage === totalPages;

        const list = document.getElementById('threat-list');
        if (allItems.length === 0) { 
            list.innerHTML = `
                <div class="empty-state">
                    <svg class="empty-icon" viewBox="0 0 24 24"><path d="M12,2L4.5,20.29L5.21,21L12,18L18.79,21L19.5,20.29L12,2Z" /></svg>
                    <div class="empty-text">SYSTEM SECURE: NO THREATS IN THIS CATEGORY</div>
                </div>`; 
        } else {
            const frag = document.createDocumentFragment();
            items.forEach((i, index) => {
                const id = i.entryId || i.fingerprint;
                const ts = i.timestamp || "", [d, t] = ts.includes(' ') ? ts.split(' ') : [ts, ""];
                let sub = i.subject || i.details || "No Subject";
                const movingStyle = i.isMoving ? 'opacity:0.5; pointer-events:none;' : '';
                if (i.isMoving) sub = `<span style="color:var(--accent); font-size:0.6rem; margin-right:8px;">[VERIFYING MOVE...]</span>` + sub;

                const row = document.createElement('div');
                row.className = `list-item ${store.selectedIds.has(id)?'selected':''}`;
                row.dataset.id = id;
                row.dataset.index = index;
                if (movingStyle) row.style.cssText = movingStyle;
                
                row.innerHTML = `
                    <div class="col-check"><input type="checkbox" ${store.selectedIds.has(id)?'checked':''} onclick="event.stopPropagation()"></div>
                    <div title="${window.escapeHTML(sub)}" style="text-align:left; font-weight:600;">${window.escapeHTML(sub)}</div>
                    <div>${window.escapeHTML(d)}</div>
                    <div>${window.escapeHTML(t)}</div>
                    <div style="font-family:monospace; font-size:0.75rem;">${window.escapeHTML(i.ip||'N/A')}</div>
                    <div style="font-family:monospace; color:var(--accent); font-weight:900;">${Math.round(i.score||100)}%</div>
                    <div style="font-weight:700; color:var(--ok);">${window.escapeHTML(i.action||'None')}</div>
                    <div title="${window.escapeHTML(i.tier||'')}" style="color:var(--muted); font-style:italic; font-size:0.7rem;">${window.escapeHTML(i.tier||'Analysis Complete')}</div>
                `;
                
                row.onclick = (e) => { 
                    if (e.shiftKey && store.lastSelectedIndex !== -1) {
                        const start = Math.min(store.lastSelectedIndex, index);
                        const end = Math.max(store.lastSelectedIndex, index);
                        for (let k = start; k <= end; k++) {
                            const item = items[k];
                            store.selectedIds.add(item.entryId || item.fingerprint);
                        }
                    } else {
                        if (!e.ctrlKey) store.selectedIds.clear();
                        if (store.selectedIds.has(id)) store.selectedIds.delete(id);
                        else store.selectedIds.add(id);
                        store.lastSelectedIndex = index;
                    }
                    window.renderList();
                }; 
                
                row.ondblclick = (e) => { 
                    e.preventDefault(); e.stopPropagation(); 
                    window.showForensics(id); 
                };
                
                row.oncontextmenu = (e) => {
                    e.preventDefault();
                    if (!store.selectedIds.has(id)) {
                        store.selectedIds.clear();
                        store.selectedIds.add(id);
                    }
                    const menu = document.getElementById('ctx-menu');
                    document.getElementById('ctx-safe-options').style.display = (store.currentCategory === 'safe') ? 'block' : 'none';
                    document.getElementById('ctx-danger-options').style.display = (store.currentCategory !== 'safe') ? 'block' : 'none';
                    
                    menu.style.display = 'block';
                    // Smart Position (Issue 1)
                    let left = e.pageX, top = e.pageY;
                    if (left + 250 > window.innerWidth) left -= 250;
                    if (top + 300 > window.innerHeight) top -= 300;
                    menu.style.left = left + 'px';
                    menu.style.top = top + 'px';
                };
                frag.appendChild(row);
            });
            list.innerHTML = '';
            list.appendChild(frag);
        }
        
        // Update Sort Indicators
        document.querySelectorAll('.sortable').forEach(h => {
            const icon = h.querySelector('.sort-icon');
            if (icon) {
                if (h.dataset.sort === store.sortConfig.key) {
                    icon.textContent = store.sortConfig.direction === 'asc' ? ' \u25B2' : ' \u25BC';
                    icon.style.color = 'var(--accent)';
                    icon.style.opacity = '1';
                } else {
                    icon.textContent = ' \u2195';
                    icon.style.color = 'var(--muted)';
                    icon.style.opacity = '0.3';
                }
            }
        });
    };

    const headerContainer = document.getElementById('main-header');
    if (headerContainer) {
        headerContainer.onclick = (e) => {
            const header = e.target.closest('.sortable');
            if (!header) return;
            const key = header.dataset.sort;
            if (store.sortConfig.key === key) store.sortConfig.direction = store.sortConfig.direction === 'asc' ? 'desc' : 'asc';
            else { store.sortConfig.key = key; store.sortConfig.direction = 'asc'; }
            window.renderList();
        };
    }

    async function handleSecurityAction(actionType, targetCategory) {
        const ids = Array.from(store.selectedIds);
        if (ids.length === 0) return;
        const currentCat = store.currentCategory;
        const cfg = await api.getConfig();
        const wl = cfg.whitelist || { emails: [], ips: [], domains: [], combos: [] };
        const bl = cfg.blacklist || { emails: [], ips: [], domains: [], combos: [] };
        
        let changed = false;
        const processMap = new Map();
        const categoriesToSearch = ['malicious', 'suspicious', 'spam', 'safe'];

        const criteria = ids.map(id => {
            const item = (store.stats[currentCat] || []).find(i => i.entryId === id);
            if (!item) return null;
            const email = (item.sender || "").match(/<(.+)>$/)?.[1] || item.sender;
            return { email, domain: email.split('@')[1] || "", ip: item.ip || "", combo: `${item.ip}|${email.split('@')[1] || ""}` };
        }).filter(c => c !== null);

        const src = (targetCategory === 'spam') ? wl : bl;
        const dst = (targetCategory === 'spam') ? bl : wl;

        criteria.forEach(c => {
            if (actionType === 'sender') {
                if (!dst.emails.includes(c.email)) { dst.emails.push(c.email); changed = true; }
                src.emails = src.emails.filter(e => e !== c.email);
                categoriesToSearch.forEach(cat => {
                    (store.stats[cat] || []).forEach(i => {
                        const iEmail = (i.sender || "").match(/<(.+)>$/)?.[1] || i.sender;
                        if (iEmail === c.email) processMap.set(i.entryId, { id: i.entryId, cat: cat, folder: i.originalFolder });
                    });
                });
            } else if (actionType === 'domain') {
                if (!dst.domains.includes(c.domain)) { dst.domains.push(c.domain); changed = true; }
                src.domains = src.domains.filter(d => d !== c.domain);
                categoriesToSearch.forEach(cat => {
                    (store.stats[cat] || []).forEach(i => {
                        const iEmail = (i.sender || "").match(/<(.+)>$/)?.[1] || i.sender;
                        if (iEmail.split('@')[1] === c.domain) processMap.set(i.entryId, { id: i.entryId, cat: cat, folder: i.originalFolder });
                    });
                });
            } else if (actionType === 'ip') {
                if (!dst.ips.includes(c.ip)) { dst.ips.push(c.ip); changed = true; }
                src.ips = src.ips.filter(i => i !== c.ip);
                categoriesToSearch.forEach(cat => {
                    (store.stats[cat] || []).forEach(i => { if (i.ip === c.ip) processMap.set(i.entryId, { id: i.entryId, cat: cat, folder: i.originalFolder }); });
                });
            } else if (actionType === 'combo') {
                if (!dst.combos.includes(c.combo)) { dst.combos.push(c.combo); changed = true; }
                src.combos = src.combos.filter(comb => comb !== c.combo);
                categoriesToSearch.forEach(cat => {
                    (store.stats[cat] || []).forEach(i => {
                        const iEmail = (i.sender || "").match(/<(.+)>$/)?.[1] || i.sender;
                        if (`${i.ip}|${iEmail.split('@')[1] || ""}` === c.combo) processMap.set(i.entryId, { id: i.entryId, cat: cat, folder: i.originalFolder });
                    });
                });
            }
        });

        const processList = Array.from(processMap.values());
        const finalIds = processList.map(p => p.id);
        if (targetCategory === 'spam') { api.quarantineEmail({ entryIds: finalIds }); window.showNotification(`Blacklisted ${finalIds.length} item(s).`); }
        else { api.releaseEmail({ entryIds: finalIds, originalFolders: processList.map(p => p.folder) }); window.showNotification(`Whitelisted ${finalIds.length} item(s).`); }

        if (changed) { await api.setWhitelist(wl); await api.setBlacklist(bl); if (typeof window.syncSettingsUI === 'function') window.syncSettingsUI(); }
        store.selectedIds.clear();
    }

    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'a') {
            e.preventDefault();
            const cat = store.currentCategory;
            const items = (store.stats[cat] || []).slice((store.currentPage - 1) * store.pageSize, store.currentPage * store.pageSize);
            items.forEach(i => store.selectedIds.add(i.entryId || i.fingerprint));
            window.renderList();
        }
    });

    document.addEventListener('click', () => { const menu = document.getElementById('ctx-menu'); if(menu) menu.style.display = 'none'; });
    document.body.addEventListener('click', (e) => {
        if (e.target.id === 'ctx-delete-email') deleteEmails();
        else if (e.target.id.startsWith('ctx-spam-')) handleSecurityAction(e.target.id.replace('ctx-spam-', ''), 'spam');
        else if (e.target.id.startsWith('ctx-safe-')) handleSecurityAction(e.target.id.replace('ctx-safe-', ''), 'safe');
    });

    document.getElementById('pg-first').onclick = () => { store.currentPage = 1; window.renderList(); };
    document.getElementById('pg-prev').onclick = () => { if (store.currentPage > 1) { store.currentPage--; window.renderList(); } };
    document.getElementById('pg-next').onclick = () => { const total = Math.ceil((store.stats[store.currentCategory]||[]).length / store.pageSize); if (store.currentPage < total) { store.currentPage++; window.renderList(); } };
    document.getElementById('pg-last').onclick = () => { store.currentPage = Math.max(1, Math.ceil((store.stats[store.currentCategory]||[]).length / store.pageSize)); window.renderList(); };

    document.querySelectorAll('.stat-card').forEach(c => c.onclick = () => { 
        document.querySelectorAll('.stat-card').forEach(s => s.classList.remove('active')); 
        c.classList.add('active'); 
        store.currentCategory = c.id.replace('card-','');
        store.currentPage = 1; 
        window.renderList(); 
    });
})();
