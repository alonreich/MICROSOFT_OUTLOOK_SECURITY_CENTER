window.SecurityUI = window.SecurityUI || {};
SecurityUI.Duplicates = (function() {
    const api = window.securityApi;
    const modal = document.getElementById('duplicate-modal');
    const list = document.getElementById('duplicate-list');
    const scanBtn = document.getElementById('scan-duplicates-btn');
    const pauseBtn = document.getElementById('pause-duplicates-btn');
    const deleteBtn = document.getElementById('delete-duplicates-btn');
    const progressBar = document.getElementById('dup-progress-bar');
    const progressFolder = document.getElementById('dup-progress-folder');
    const progressItem = document.getElementById('dup-progress-item');
    const statFound = document.getElementById('dup-stat-found');
    const statScanned = document.getElementById('dup-stat-scanned');
    const log = document.getElementById('duplicate-log');

    function formatSize(bytes) {
        if (!bytes) return "0 B";
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    function addDupLog(msg, type = 'info') {
        const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
        const s = window.AppStore.duplicateState;
        s.logs.push({ text: line, type });
        if (s.logs.length > 500) s.logs.shift();
        
        const div = document.createElement('div');
        div.textContent = line;
        if (type === 'error') div.style.color = 'var(--danger)';
        if (type === 'warn') div.style.color = 'var(--warn)';
        if (type === 'success') div.style.color = 'var(--ok)';
        log.appendChild(div);
        log.scrollTop = log.scrollHeight;
    }

    function updateButtons() {
        const s = window.AppStore.duplicateState;
        scanBtn.textContent = s.isScanning ? 'SCANNING...' : 'SCAN MAILBOX';
        scanBtn.disabled = s.isScanning;
        scanBtn.style.opacity = s.isScanning ? 0.5 : 1;
        
        pauseBtn.style.display = (s.isScanning || s.isPaused) ? 'block' : 'none';
        pauseBtn.textContent = s.isPaused ? 'RESUME' : 'PAUSE';
        
        const checked = list.querySelectorAll('input[type="checkbox"]:checked').length;
        deleteBtn.disabled = checked === 0;
        deleteBtn.style.opacity = checked === 0 ? 0.5 : 1;
    }

    function restoreUIFromState() {
        const s = window.AppStore.duplicateState;
        statScanned.textContent = s.scannedCount;
        statFound.textContent = s.items.length;
        progressBar.style.width = s.progress + '%';
        
        log.innerHTML = '';
        s.logs.forEach(l => {
            const div = document.createElement('div');
            div.textContent = l.text;
            if (l.type === 'error') div.style.color = 'var(--danger)';
            if (l.type === 'warn') div.style.color = 'var(--warn)';
            if (l.type === 'success') div.style.color = 'var(--ok)';
            log.appendChild(div);
        });
        log.scrollTop = log.scrollHeight;
        
        if (s.items.length > 0 || s.isScanning || s.isPaused) {
            document.getElementById('duplicate-progress-card').style.display = 'block';
            log.style.display = 'block';
        }
        
        renderDuplicateList(s.items);
        updateButtons();
    }

    list.addEventListener('change', (e) => {
        if (e.target.type === 'checkbox') updateButtons();
    });

    api.onDuplicateUpdate(d => {
        const s = window.AppStore.duplicateState;
        
        if (d.status === 'Paused') {
            s.isScanning = false;
            s.isPaused = true;
            addDupLog('SCAN PAUSED: State preserved in memory.', 'warn');
            progressFolder.textContent = 'PAUSED';
            progressItem.textContent = 'Engine is standing by...';
            updateButtons();
            return;
        }

        if (d.status === 'StoreStart') {
            s.storesMeta[d.store] = { totalSize: d.size, totalItems: 0, scannedItems: 0, scannedSize: 0, found: 0 };
            addDupLog(`PROBING STORE: ${d.store} (Size: ${formatSize(d.size)})`);
            return;
        }

        if (d.status === 'StoreMeta') {
            if (s.storesMeta[d.store]) s.storesMeta[d.store].totalItems = d.totalItems;
            addDupLog(`INDEXED: ${d.store} - ${d.totalItems} emails found.`);
            return;
        }

        if (d.status === 'StoreFinish') {
            const meta = s.storesMeta[d.store] || { found: 0 };
            window.showNotification(`FINISHED: ${d.store}\nScanned: ${d.scanned} emails (${formatSize(d.size)})\nDuplicates: ${meta.found}`);
            addDupLog(`COMPLETED: ${d.store}. Found ${meta.found} redundant copies in this store.`, 'success');
            return;
        }

        if (!s.isScanning) return;
        
        if (d.status === 'Scanned') {
            s.scannedCount = d.scanned;
            statScanned.textContent = s.scannedCount;
            statFound.textContent = s.items.length;
            
            const meta = s.storesMeta[d.store];
            if (meta) {
                meta.scannedItems = d.storeScanned;
                meta.scannedSize = d.storeScannedSize;
                const pct = meta.totalItems > 0 ? Math.round((meta.scannedItems / meta.totalItems) * 100) : 0;
                s.progress = pct;
                progressFolder.textContent = `SCANNING: ${d.store} [${pct}%]`;
                progressItem.textContent = `CHECKING (${meta.scannedItems}/${meta.totalItems}): "${d.currentItem || 'Email'}" [${formatSize(meta.scannedSize)} / ${formatSize(meta.totalSize)}]`;
                progressBar.style.width = pct + '%';
            }

        } else if (d.status === 'Found') {
            addDupLog(`Identified Redundant Copy: "${d.current}"`, 'warn');
            if (s.storesMeta[d.store]) s.storesMeta[d.store].found++;
            
            if (d.items) {
                s.items = d.items;
                renderDuplicateList(s.items);
            }
        } else if (d.status === 'Finished') {
            s.isScanning = false;
            s.isPaused = false;
            s.progress = 100;
            addDupLog('FULL MAILBOX DISCOVERY COMPLETE.', 'success');
            progressFolder.textContent = 'SCAN FINISHED';
            progressItem.textContent = `Total Scanned: ${s.scannedCount} emails. Found ${d.items ? d.items.length : s.items.length} duplicates.`;
            progressBar.style.width = '100%';
            if (d.items) s.items = d.items;
            renderDuplicateList(s.items);
            updateButtons();
        }
    });

    function renderDuplicateList(items) {
        if (!items || items.length === 0) {
            list.innerHTML = '<div class="empty-state">NO REDUNDANT COPIES FOUND</div>';
            return;
        }
        list.innerHTML = '';
        items.forEach(i => {
            const row = document.createElement('div');
            row.className = 'list-item';
            row.innerHTML = `
                <div class="col-check"><input type="checkbox" value="${i.entryId}"></div>
                <div title="${window.escapeHTML(i.subject)}">${window.escapeHTML(i.subject)}</div>
                <div title="${window.escapeHTML(i.sender)}">${window.escapeHTML(i.sender)}</div>
                <div>${i.timestamp}</div>
                <div>${formatSize(i.size)}</div>
                <div title="${window.escapeHTML(i.folder)}">${window.escapeHTML(i.folder)}</div>
                <div style="color:var(--warn); font-style:italic;">Identical DNA found in ${i.store}</div>
            `;
            list.appendChild(row);
        });
        updateButtons();
    }

    scanBtn.onclick = async () => {
        const s = window.AppStore.duplicateState;
        if (s.isPaused) {
            await api.resumeDuplicateScan();
            s.isPaused = false;
            s.isScanning = true;
            updateButtons();
            return;
        }
        
        s.isScanning = true;
        s.isPaused = false;
        await api.scanDuplicates();
        updateButtons();
    };

    pauseBtn.onclick = async () => {
        const s = window.AppStore.duplicateState;
        if (s.isPaused) {
            await api.resumeDuplicateScan();
            s.isPaused = false;
            s.isScanning = true;
        } else {
            await api.pauseDuplicateScan();
            s.isScanning = false;
            s.isPaused = true;
        }
        updateButtons();
    };

    document.getElementById('rescan-everything-btn').onclick = () => {
        window.showConfirmModal(
            'RESCAN EVERYTHING',
            'This will erase the current list and restart the scan from the very beginning. Proceed?',
            async () => {
                window.AppStore.duplicateState = {
                    items: [],
                    scannedCount: 0,
                    storesMeta: {},
                    isScanning: true,
                    isPaused: false,
                    progress: 0,
                    logs: []
                };
                list.innerHTML = '';
                progressBar.style.width = '0%';
                document.getElementById('duplicate-progress-card').style.display = 'block';
                log.style.display = 'block';
                log.innerHTML = '';
                addDupLog('Restarting fresh discovery scan...');
                
                await api.resetDuplicateEngine(); 
                await api.scanDuplicates();
                updateButtons();
            }
        );
    };

    document.getElementById('duplicate-detection-btn').onclick = () => {
        modal.style.display = 'flex';
        document.getElementById('settings-modal').style.display = 'none';
        restoreUIFromState();
    };

    document.getElementById('close-duplicate-modal').onclick = () => {
        modal.style.display = 'none';
    };

    deleteBtn.onclick = () => {
        const checkboxes = list.querySelectorAll('input[type="checkbox"]:checked');
        const ids = Array.from(checkboxes).map(cb => cb.value);
        if (ids.length === 0) return;

        window.showConfirmModal(
            'BATCH CLEANUP',
            `Are you sure you want to permanently delete ${ids.length} redundant email copies?`,
            async () => {
                deleteBtn.disabled = true;
                deleteBtn.textContent = 'DELETING...';
                const res = await api.deleteDuplicates({ entryIds: ids });
                if (res.ok) {
                    window.showNotification(`Successfully removed ${ids.length} redundant emails.`);
                    window.AppStore.duplicateState.items = window.AppStore.duplicateState.items.filter(i => !ids.includes(i.entryId));
                    renderDuplicateList(window.AppStore.duplicateState.items);
                }
            }
        );
    };
})();
