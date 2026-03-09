document.body.insertAdjacentHTML('beforeend', `
<div id="ctx-menu" class="context-menu">
    <div id="ctx-safe-options" style="display:none">
        <div class="context-menu-item" id="ctx-spam-sender" style="color:var(--danger)">Mark as SPAM: Only This Email Sender</div>
        <div class="context-menu-item" id="ctx-spam-domain" style="color:var(--danger)">Mark as SPAM: Only This Domain</div>
        <div class="context-menu-item" id="ctx-spam-ip" style="color:var(--danger)">Mark as SPAM: Only This IP Address</div>
        <div class="context-menu-item" id="ctx-spam-combo" style="color:var(--danger)">Mark as SPAM: Domain & IP Address Combination</div>
    </div>
    <div id="ctx-danger-options" style="display:none">
        <div class="context-menu-item" id="ctx-safe-sender" style="color:var(--ok)">Mark as SAFE: Only This Email Sender</div>
        <div class="context-menu-item" id="ctx-safe-domain" style="color:var(--ok)">Mark as SAFE: Only This Domain</div>
        <div class="context-menu-item" id="ctx-safe-ip" style="color:var(--ok)">Mark as SAFE: Only This IP Address</div>
        <div class="context-menu-item" id="ctx-safe-combo" style="color:var(--ok)">Mark as SAFE: Domain & IP Address Combination</div>
    </div>
</div>
`);

(function() {
    const api = window.securityApi;
    window.selectedIds = new Set();
    let currentPage = 1, pageSize = 50, renderPending = false, lastRenderedJson = "";

    window.renderList = function() {
        if (renderPending) return;
        renderPending = true;
        requestAnimationFrame(() => {
            const cat = document.querySelector('.stat-card.active')?.id.replace('card-','') || 'malicious';
            const allItems = (window.stats[cat] || []).slice().reverse();
            const totalPages = Math.max(1, Math.ceil(allItems.length / pageSize));
            if (currentPage > totalPages) currentPage = totalPages;
            const start = (currentPage - 1) * pageSize;
            const items = allItems.slice(start, start + pageSize);
            
            document.getElementById('pg-info').textContent = `Page ${currentPage} of ${totalPages}`;
            document.getElementById('pg-first').disabled = currentPage === 1;
            document.getElementById('pg-prev').disabled = currentPage === 1;
            document.getElementById('pg-next').disabled = currentPage === totalPages;
            document.getElementById('pg-last').disabled = currentPage === totalPages;

            const hash = JSON.stringify(items.map(i => i.entryId + (window.selectedIds.has(i.entryId) ? '_s' : '')));
            if (hash !== lastRenderedJson) {
                lastRenderedJson = hash;
                const list = document.getElementById('threat-list');
                if (allItems.length === 0) { 
                    list.innerHTML = `<div style="grid-column: 1 / -1; padding:40px; text-align:center; color:var(--muted); font-size:0.8rem; font-weight:800; width: 100%;">NO EMAILS DETECTED</div>`; 
                } else {
                    list.innerHTML = items.map(i => {
                        const ts = i.timestamp || "", [d, t] = ts.includes(' ') ? ts.split(' ') : [ts, ""];
                        const vColor = cat === 'safe' ? 'var(--ok)' : cat === 'malicious' ? 'var(--danger)' : cat === 'suspicious' ? 'var(--warn)' : '#8e75ff';
                        const sub = i.subject || i.details || "No Subject";
                        return `<div class="list-item ${window.selectedIds.has(i.entryId)?'selected':''}" data-id="${window.escapeHTML(i.entryId)}" data-folder="${window.escapeHTML(i.originalFolder||'')}"><div title="${window.escapeHTML(sub)}">${window.escapeHTML(sub)}</div><div>${window.escapeHTML(d)}</div><div>${window.escapeHTML(t)}</div><div>${window.escapeHTML(i.ip||'N/A')}</div><div style="color:${vColor}; font-weight:bold;">${cat.toUpperCase()}</div><div style="font-family:monospace; color:var(--accent); font-weight:900;">${Math.round(i.score||100)}%</div><div>${window.escapeHTML(i.action||'None')}</div><div title="${window.escapeHTML(i.tier||'')}">${window.escapeHTML(i.tier||'')}</div></div>`;
                    }).join('');
                    
                    document.querySelectorAll('.list-item').forEach(el => { 
                        el.onclick = (e) => { 
                            const id = el.dataset.id; 
                            if (!e.ctrlKey) { 
                                document.querySelectorAll('.list-item').forEach(x => x.classList.remove('selected')); 
                                window.selectedIds.clear(); 
                            }
                            if (window.selectedIds.has(id)) { 
                                window.selectedIds.delete(id); 
                                el.classList.remove('selected'); 
                            } else { 
                                window.selectedIds.add(id); 
                                el.classList.add('selected'); 
                            }
                        }; 
                        el.ondblclick = (e) => { 
                            e.preventDefault(); e.stopPropagation(); 
                            window.showForensics(el.dataset.id); 
                        };
                        
                        el.oncontextmenu = (e) => {
                            e.preventDefault();
                            if (!window.selectedIds.has(el.dataset.id)) {
                                document.querySelectorAll('.list-item').forEach(x => x.classList.remove('selected')); 
                                window.selectedIds.clear();
                                window.selectedIds.add(el.dataset.id);
                                el.classList.add('selected');
                            }
                            const menu = document.getElementById('ctx-menu');
                            const cat = document.querySelector('.stat-card.active')?.id.replace('card-','') || 'malicious';
                            
                            document.getElementById('ctx-safe-options').style.display = (cat === 'safe') ? 'block' : 'none';
                            document.getElementById('ctx-danger-options').style.display = (cat !== 'safe') ? 'block' : 'none';

                            menu.style.display = 'block';
                            menu.style.left = e.pageX + 'px';
                            menu.style.top = e.pageY + 'px';
                        };
                    });
                }
            }
            renderPending = false;
        });
    };

    async function handleSecurityAction(actionType, targetCategory) {
        const ids = Array.from(window.selectedIds);
        if (ids.length === 0) return;
        
        const currentCat = document.querySelector('.stat-card.active')?.id.replace('card-','') || 'malicious';
        const cfg = await api.getConfig();
        const whitelist = cfg.whitelist || { emails: [], ips: [], domains: [], combos: [] };
        const blacklist = cfg.blacklist || { emails: [], ips: [], domains: [], combos: [] };
        
        let changed = false;
        const allIdsToProcess = new Set(ids);

        for (const id of ids) {
            const item = window.stats[currentCat].find(i => i.entryId === id);
            if (!item) continue;

            const email = (item.sender || "").match(/<(.+)>$/)?.[1] || item.sender;
            const domain = email.split('@')[1] || "";
            const ip = item.ip || "";
            const combo = `${ip}|${domain}`;

            const source = (targetCategory === 'spam') ? whitelist : blacklist;
            const destination = (targetCategory === 'spam') ? blacklist : whitelist;

            if (actionType === 'sender') {
                if (!destination.emails.includes(email)) { destination.emails.push(email); changed = true; }
                source.emails = source.emails.filter(e => e !== email);
                window.stats[currentCat].forEach(i => { if (((i.sender || "").match(/<(.+)>$/)?.[1] || i.sender) === email) allIdsToProcess.add(i.entryId); });
            } else if (actionType === 'domain') {
                if (!destination.domains.includes(domain)) { destination.domains.push(domain); changed = true; }
                source.domains = source.domains.filter(d => d !== domain);
                window.stats[currentCat].forEach(i => { if (((i.sender || "").match(/<(.+)>$/)?.[1] || i.sender).split('@')[1] === domain) allIdsToProcess.add(i.entryId); });
            } else if (actionType === 'ip') {
                if (!destination.ips.includes(ip)) { destination.ips.push(ip); changed = true; }
                source.ips = source.ips.filter(i => i !== ip);
                window.stats[currentCat].forEach(i => { if (i.ip === ip) allIdsToProcess.add(i.entryId); });
            } else if (actionType === 'combo') {
                if (!destination.combos.includes(combo)) { destination.combos.push(combo); changed = true; }
                source.combos = source.combos.filter(c => c !== combo);
                window.stats[currentCat].forEach(i => { if (`${i.ip}|${((i.sender || "").match(/<(.+)>$/)?.[1] || i.sender).split('@')[1]}` === combo) allIdsToProcess.add(i.entryId); });
            }
        }

        const finalIds = Array.from(allIdsToProcess);
        if (targetCategory === 'spam') {
            api.quarantineEmail({ entryIds: finalIds });
        } else {
            const originalFolders = finalIds.map(id => {
                const node = document.querySelector(`.list-item[data-id="${id}"]`);
                return node ? node.dataset.folder : null;
            });
            api.releaseEmail({ entryIds: finalIds, originalFolders: originalFolders });
        }

        if (changed) {
            await api.setWhitelist(whitelist);
            await api.setBlacklist(blacklist);
            if (typeof window.syncSettingsUI === 'function') window.syncSettingsUI();
        }

        window.selectedIds.clear();
        window.renderList();
    }

    document.addEventListener('click', (e) => {
        const menu = document.getElementById('ctx-menu');
        if(menu) menu.style.display = 'none';
    });

    document.body.addEventListener('click', (e) => {
        if (e.target.id.startsWith('ctx-spam-')) {
            handleSecurityAction(e.target.id.replace('ctx-spam-', ''), 'spam');
        } else if (e.target.id.startsWith('ctx-safe-')) {
            handleSecurityAction(e.target.id.replace('ctx-safe-', ''), 'safe');
        }
    });

    document.getElementById('pg-first').onclick = () => { currentPage = 1; window.renderList(); };
    document.getElementById('pg-prev').onclick = () => { if (currentPage > 1) { currentPage--; window.renderList(); } };
    document.getElementById('pg-next').onclick = () => { const cat = document.querySelector('.stat-card.active')?.id.replace('card-','') || 'malicious'; const total = Math.ceil((window.stats[cat]||[]).length / pageSize); if (currentPage < total) { currentPage++; window.renderList(); } };
    document.getElementById('pg-last').onclick = () => { const cat = document.querySelector('.stat-card.active')?.id.replace('card-','') || 'malicious'; currentPage = Math.max(1, Math.ceil((window.stats[cat]||[]).length / pageSize)); window.renderList(); };

    document.querySelectorAll('.stat-card').forEach(c => c.onclick = () => { 
        document.querySelectorAll('.stat-card').forEach(s => s.classList.remove('active')); 
        c.classList.add('active'); 
        currentPage = 1; 
        lastRenderedJson = ""; 
        window.renderList(); 
    });
})();
