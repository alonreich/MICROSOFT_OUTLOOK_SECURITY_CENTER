(function() {
    const api = window.securityApi;
    const modal = document.getElementById('duplicate-modal');
    const list = document.getElementById('duplicate-list');
    const scanBtn = document.getElementById('scan-duplicates-btn');
    const pauseBtn = document.getElementById('pause-duplicates-btn');
    const deleteBtn = document.getElementById('delete-duplicates-btn');
    const progressCard = document.getElementById('duplicate-progress-card');
    const progressFolder = document.getElementById('dup-progress-folder');
    const progressItem = document.getElementById('dup-progress-item');
    const statFound = document.getElementById('dup-stat-found');
    const statScanned = document.getElementById('dup-stat-scanned');
    const progressBar = document.getElementById('dup-progress-bar');
    const logArea = document.getElementById('duplicate-log');
    const checkAll = document.getElementById('check-all-duplicates');
    
    let isScanning = false;
    let isPaused = false;
    let foundDuplicates = [];
    let selectedIds = new Set();
    let scannedCount = 0;

    function addDupLog(msg, type = 'info') {
        const div = document.createElement('div');
        const ts = new Date().toLocaleTimeString();
        div.innerHTML = `<span style="color:var(--muted)">[${ts}]</span> <span style="color:${type === 'error' ? 'var(--danger)' : 'var(--accent)'}">${msg}</span>`;
        logArea.appendChild(div);
        logArea.scrollTop = logArea.scrollHeight;
    }

    document.getElementById('duplicate-detection-btn').onclick = () => {
        const settingsModal = document.getElementById('settings-modal');
        if (settingsModal) settingsModal.style.display = 'none';
        modal.style.display = 'flex';
        logArea.innerHTML = '';
        logArea.style.display = 'none';
        progressCard.style.display = 'none';
        pauseBtn.style.display = 'none';
        list.innerHTML = `
            <div class="empty-state">
                <svg class="empty-icon" viewBox="0 0 24 24"><path d="M15,14H13V11H11V14H9L12,17L15,14M21,16.5C21,16.88 20.79,17.21 20.47,17.38L12.57,21.82C12.41,21.94 12.21,22 12,22C11.79,22 11.59,21.94 11.43,21.82L3.53,17.38C3.21,17.21 3,16.88 3,16.5V7.5C3,7.12 3.21,6.79 3.53,6.62L11.43,2.18C11.59,2.06 11.79,2 12,2C12.21,2 12.41,2.06 12.57,2.18L20.47,6.62C20.79,6.79 21,7.12 21,7.5V16.5M12,4.15L5,8.09V15.91L12,19.85L19,15.91V8.09L12,4.15Z" /></svg>
                <div class="empty-text">START SCAN TO FIND REDUNDANT COPIES</div>
            </div>`;
        foundDuplicates = [];
        selectedIds.clear();
        scannedCount = 0;
        isPaused = false;
        checkAll.checked = false;
        updateButtons();
    };

    document.getElementById('close-duplicate-modal').onclick = () => {
        if (isScanning) isScanning = false;
        modal.style.display = 'none';
    };

    scanBtn.onclick = async () => {
        if (isScanning || isPaused) {
            addDupLog('User requested scan termination. Cleaning up engine state...');
            isScanning = false;
            isPaused = false;
            await api.resumeDuplicateScan();
            scanBtn.textContent = 'SCAN MAILBOX';
            scanBtn.className = 'btn-ui success';
            pauseBtn.style.display = 'none';
            progressCard.style.display = 'none';
            addDupLog('Scan sequence aborted. Engine standing by.');
            return;
        }

        addDupLog('Phase 1: Validating IPC bridge and engine status...');
        isScanning = true;
        isPaused = false;
        logArea.style.display = 'block';
        logArea.innerHTML = '';
        progressCard.style.display = 'block';
        progressFolder.textContent = 'CONNECTING TO ENGINE...';
        progressItem.textContent = 'Establishing secure handshake with Outlook...';
        statFound.textContent = '0';
        statScanned.textContent = '0';
        progressBar.style.width = '0%';
        
        addDupLog('Phase 2: Requesting master mailbox crawl from security worker...');
        scanBtn.textContent = 'STOP SCAN';
        scanBtn.className = 'btn-ui danger';
        pauseBtn.style.display = 'block';
        pauseBtn.textContent = 'PAUSE';
        pauseBtn.className = 'btn-ui';
        
        list.innerHTML = '';
        foundDuplicates = [];
        selectedIds.clear();
        scannedCount = 0;
        updateButtons();

        // Safety Timeout (Issue Detection)
        const scanTimeout = setTimeout(() => {
            if (isScanning && scannedCount === 0 && foundDuplicates.length === 0) {
                addDupLog('CRITICAL: Engine response delay detected. This may be due to Outlook being busy or a large mailbox index.', 'error');
                addDupLog('HINT: Check if any Outlook modal windows are open or if Outlook is performing an update.', 'info');
            }
        }, 15000);

        try {
            const res = await api.scanDuplicates();
            if (!res.ok) {
                clearTimeout(scanTimeout);
                addDupLog(`IPC FAILURE: ${res.error || 'Unknown communication error'}`, 'error');
                isScanning = false;
                scanBtn.textContent = 'SCAN MAILBOX';
                scanBtn.className = 'btn-ui success';
                progressCard.style.display = 'none';
                pauseBtn.style.display = 'none';
            } else {
                addDupLog('Phase 3: Handshake accepted. Waiting for discovery heartbeat...');
            }
        } catch (err) {
            clearTimeout(scanTimeout);
            addDupLog('SYSTEM CRASH: IPC pipe was severed during request.', 'error');
            isScanning = false;
        }
    };

    pauseBtn.onclick = async () => {
        if (!isPaused) {
            // Request Pause
            const res = await api.pauseDuplicateScan();
            if (res.ok) {
                isPaused = true;
                pauseBtn.textContent = 'RESUME';
                pauseBtn.className = 'btn-ui success';
                addDupLog('Pause signal sent. Waiting for engine to stabilize...');
            }
        } else {
            // Request Resume
            const res = await api.resumeDuplicateScan();
            if (res.ok) {
                isPaused = false;
                isScanning = true;
                pauseBtn.textContent = 'PAUSE';
                pauseBtn.className = 'btn-ui';
                addDupLog('Resuming discovery from last known position...');
                api.scanDuplicates(); // Trigger worker again
            }
        }
    };

    api.onDuplicateUpdate(d => {
        if (d.status === 'Paused') {
            isScanning = false;
            addDupLog('SCAN PAUSED: State preserved in memory.', 'warn');
            progressFolder.textContent = 'PAUSED';
            progressItem.textContent = 'Engine is standing by...';
            return;
        }

        if (!isScanning) return;
        
        if (d.status === 'Found') {
            scannedCount++;
            statScanned.textContent = scannedCount;
            statFound.textContent = d.count;
            progressItem.textContent = `Matching: "${d.current}"`;
            addDupLog(`Identified Redundant Copy: "${d.current}"`);
            
            const currentPct = Math.min(95, (scannedCount / (scannedCount + 10)) * 100);
            progressBar.style.width = currentPct + '%';
            
        } else if (d.status === 'Finished') {
            isScanning = false;
            isPaused = false;
            scanBtn.textContent = 'SCAN MAILBOX';
            scanBtn.className = 'btn-ui success';
            pauseBtn.style.display = 'none';
            progressBar.style.width = '100%';
            progressFolder.textContent = 'SCAN COMPLETE';
            progressItem.textContent = `Identification complete. ${d.items?.length || 0} duplicates mapped.`;
            
            foundDuplicates = d.items || [];
            renderDuplicateList();
            addDupLog(`Crawl Complete. Total Redundant Items: ${foundDuplicates.length}`);
            window.showNotification(`Scan complete. ${foundDuplicates.length} duplicates found.`);
        } else if (d.status === 'Progress') {
            progressFolder.textContent = d.details;
            addDupLog(d.details);
        }
    });

    function renderDuplicateList() {
        if (foundDuplicates.length === 0) {
            list.innerHTML = `
                <div class="empty-state">
                    <svg class="empty-icon" viewBox="0 0 24 24"><path d="M12,2L4.5,20.29L5.21,21L12,18L18.79,21L19.5,20.29L12,2Z" /></svg>
                    <div class="empty-text">SYSTEM CLEAN: NO DUPLICATES DETECTED</div>
                </div>`;
            return;
        }

        const frag = document.createDocumentFragment();
        foundDuplicates.forEach(i => {
            const row = document.createElement('div');
            row.className = `list-item ${selectedIds.has(i.entryId) ? 'selected' : ''}`;
            row.style.gridTemplateColumns = '40px 350px 250px 150px 100px 150px 1fr';
            
            const sizeKb = Math.round(i.size / 1024) + ' KB';
            const location = `${i.store} \\ ${i.folder}`;

            row.innerHTML = `
                <div class="col-check"><input type="checkbox" ${selectedIds.has(i.entryId) ? 'checked' : ''}></div>
                <div title="${window.escapeHTML(i.subject)}" style="text-align:left; font-weight:600;">${window.escapeHTML(i.subject)}</div>
                <div title="${window.escapeHTML(i.sender)}">${window.escapeHTML(i.sender)}</div>
                <div>${window.escapeHTML(i.timestamp)}</div>
                <div style="font-family:monospace;">${sizeKb}</div>
                <div style="font-weight:700; color:var(--muted); font-size:0.6rem;">${window.escapeHTML(location)}</div>
                <div style="color:var(--danger); font-style:italic; font-size:0.7rem;">SHA256 Content Collision</div>
            `;

            row.onclick = (e) => {
                const cb = row.querySelector('input[type="checkbox"]');
                if (e.target !== cb) cb.checked = !cb.checked;
                if (cb.checked) { selectedIds.add(i.entryId); row.classList.add('selected'); }
                else { selectedIds.delete(i.entryId); row.classList.remove('selected'); }
                updateButtons();
            };
            frag.appendChild(row);
        });
        list.innerHTML = '';
        list.appendChild(frag);
        updateButtons();
    }

    checkAll.onclick = () => {
        const check = checkAll.checked;
        selectedIds.clear();
        document.querySelectorAll('#duplicate-list .list-item').forEach((row, idx) => {
            const cb = row.querySelector('input[type="checkbox"]');
            cb.checked = check;
            if (check) { selectedIds.add(foundDuplicates[idx].entryId); row.classList.add('selected'); }
            else { row.classList.remove('selected'); }
        });
        updateButtons();
    };

    function updateButtons() {
        const hasSelection = selectedIds.size > 0;
        deleteBtn.disabled = !hasSelection;
        deleteBtn.style.opacity = hasSelection ? '1' : '0.5';
    }

    deleteBtn.onclick = () => {
        const ids = Array.from(selectedIds);
        if (ids.length === 0) return;

        window.showConfirmModal(
            'PERMANENT DELETION',
            `Are you sure you want to delete ${ids.length} redundant copies? This will free up Outlook storage space.`,
            async () => {
                addDupLog(`Initiating deletion of ${ids.length} items...`);
                deleteBtn.disabled = true;
                deleteBtn.textContent = 'PROCESSING...';
                
                const res = await api.deleteDuplicates({ entryIds: ids });
                if (res.ok) {
                    api.onFromMain(m => {
                        if (m.type === 'delete-summary') {
                            addDupLog(`DELETION SUCCESS: ${m.count} items removed from Outlook.`, 'info');
                            window.showNotification(`Cleanup Successful: ${m.count} duplicates deleted.`);
                            foundDuplicates = foundDuplicates.filter(i => !selectedIds.has(i.entryId));
                            selectedIds.clear();
                            checkAll.checked = false;
                            renderDuplicateList();
                            deleteBtn.textContent = 'DELETE SELECTED';
                            updateButtons();
                        }
                    });
                } else {
                    addDupLog('Deletion failed.', 'error');
                    deleteBtn.textContent = 'DELETE SELECTED';
                    updateButtons();
                }
            }
        );
    };
})();
