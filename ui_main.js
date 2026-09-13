// Centralized Sanitization & HTML Entity Encoder
window.escapeHtml = function(str) {
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

const api = window.securityApi;

// CENTRALIZED STATE MANAGER
window.AppState = {
    currentTab: 'dashboard',
    currentCategory: 'malicious',
    stats: { malicious: [], suspicious: [], spam: [], safe: [] },
    config: null,
    
    updateStats(data) {
        if (data.full) this.stats = data.stats;
        this.renderDashboard(false);
    },
    
    renderDashboard(resetScroll = false) {
        const totalAll = (this.stats.malicious ? this.stats.malicious.length : 0) +
                         (this.stats.suspicious ? this.stats.suspicious.length : 0) +
                         (this.stats.spam ? this.stats.spam.length : 0) +
                         (this.stats.safe ? this.stats.safe.length : 0);

        const elAll = document.getElementById('stat-all');
        if (elAll) elAll.textContent = totalAll;
        document.getElementById('stat-malicious').textContent = this.stats.malicious ? this.stats.malicious.length : 0;
        document.getElementById('stat-suspicious').textContent = this.stats.suspicious ? this.stats.suspicious.length : 0;
        document.getElementById('stat-spam').textContent = this.stats.spam ? this.stats.spam.length : 0;
        document.getElementById('stat-safe').textContent = this.stats.safe ? this.stats.safe.length : 0;
        
        if (window.vaultSearchActive && typeof window.executeVaultSearch === 'function') {
            window.executeVaultSearch();
            return;
        }

        let list;
        let title;
        if (this.currentCategory === 'all') {
            list = [
                ...(this.stats.malicious || []),
                ...(this.stats.suspicious || []),
                ...(this.stats.spam || []),
                ...(this.stats.safe || [])
            ];
            title = "All Scanned Emails";
        } else {
            list = this.stats[this.currentCategory] || [];
            title = this.currentCategory.charAt(0).toUpperCase() + this.currentCategory.slice(1) + " Incidents";
        }

        window.renderList(list, this.currentCategory, resetScroll);
        const titleEl = document.getElementById('current-list-title');
        if (titleEl) titleEl.textContent = title;
    }
};

function switchTab(tabId) {
    if (window.AppState.currentTab === tabId) return;
    
    const oldTab = document.getElementById('tab-' + window.AppState.currentTab);
    const newTab = document.getElementById('tab-' + tabId);
    if (!newTab) return;
    
    if (oldTab) oldTab.classList.remove('active');
    setTimeout(() => {
        window.AppState.currentTab = tabId;
        document.querySelectorAll('.nav-item').forEach(p => p.classList.remove('active'));
        const navItem = document.querySelector(`.nav-item[data-tab="${tabId}"]`);
        if (navItem) navItem.classList.add('active');
        
        newTab.classList.add('active');

        if (tabId === 'duplicates') {
            const dupView = document.getElementById('duplicate-view');
            if (dupView && (!dupView.children.length || dupView.innerHTML.trim() === '')) {
                if (typeof window.renderDuplicateInitialView === 'function') {
                    window.renderDuplicateInitialView();
                }
            }
        }
    }, 50);
}

function switchCategory(cat) {
    const changed = (window.AppState.currentCategory !== cat);
    window.AppState.currentCategory = cat;
    window.AppState.renderDashboard(changed);
}

function openModal(id) {
    const m = document.getElementById(id);
    m.style.display = 'flex';
    setTimeout(() => m.classList.add('active'), 10);
}

function closeModal(id) {
    const m = document.getElementById(id);
    m.classList.remove('active');
    setTimeout(() => m.style.display = 'none', 300);
}

function handleOverlayClick(e, id) {
    if (e.target === document.getElementById(id)) {
        closeModal(id);
    }
}

function openSettings() {
    openModal('settings-modal');
    window.syncSettingsUI();
}

function closeSettings() {
    closeModal('settings-modal');
}

function closeForensics() {
    closeModal('forensics-modal');
}

function closeVirusTotalModal() {
    closeModal('virustotal-modal');
}
window.closeVirusTotalModal = closeVirusTotalModal;

let confirmCallback = null;
function showConfirm(title, msg, keyword, onExecute) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = msg;
    const inputCont = document.getElementById('confirm-input-container');
    const input = document.getElementById('confirm-input');
    const kwLabel = document.getElementById('confirm-keyword');
    const executeBtn = document.getElementById('confirm-execute-btn');
    
    input.value = '';
    if (keyword) {
        inputCont.style.display = 'block';
        kwLabel.textContent = keyword;
        executeBtn.textContent = "Confirm Action";
    } else {
        inputCont.style.display = 'none';
        executeBtn.textContent = "Yes, Proceed";
    }
    
    openModal('confirm-modal');
    confirmCallback = () => {
        if (keyword && input.value.toUpperCase() !== keyword.toUpperCase()) {
            input.style.borderColor = 'var(--danger)';
            return;
        }
        onExecute();
        closeConfirm();
    };
    executeBtn.onclick = confirmCallback;
}

function closeConfirm() {
    closeModal('confirm-modal');
    confirmCallback = null;
}

async function requestFullScan() {
    window.showConfirm("Full Mailbox Audit", "This will perform a deep forensic scan of ALL items in your mailbox. This may take several minutes. Proceed?", "AUDIT", async () => {
        const btn = document.getElementById('full-scan-btn');
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Initializing...`;
        btn.disabled = true;
        // FORCE FULL RESCAN BY CLEARING PROCESSED CACHE
        await api.setEnabled(true);
        await api.setHistoryEnabled(true);
    });
}

async function factoryReset() {
    showConfirm(
        "DANGER: Factory Reset",
        "This will permanently wipe all security databases, settings, and logs. This action cannot be undone. Are you sure?",
        null,
        async () => {
            try { localStorage.removeItem('deskguard_first_run_completed'); } catch (e) {}
            await api.resetApp();
        }
    );
}

window.addActivity = (msg, type = 'info') => {
    const stream = document.getElementById('activity-stream');
    if (!stream) return;
    
    const line = document.createElement('div');
    line.className = 'activity-line';
    
    const ts = new Date().toLocaleTimeString([], { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });
    let msgClass = '';
    if (msg.includes('CRITICAL') || msg.includes('ERROR') || msg.includes('BLOCKED')) msgClass = 'danger';
    else if (msg.includes('Success') || msg.includes('Verified') || msg.includes('CLEAN')) msgClass = 'success';
    else if (msg.includes('Audit') || msg.includes('Engine') || msg.includes('Analyzing')) msgClass = 'highlight';

    const tsSpan = document.createElement('span');
    tsSpan.className = 'activity-ts';
    tsSpan.textContent = ts;

    const msgSpan = document.createElement('span');
    msgSpan.className = `activity-msg ${msgClass}`.trim();
    msgSpan.textContent = msg;

    line.appendChild(tsSpan);
    line.appendChild(msgSpan);
    
    stream.appendChild(line);
    stream.scrollTop = stream.scrollHeight;
    if (stream.children.length > 100) stream.removeChild(stream.firstChild);
};

// Override addLog for backward compatibility and to use the new stream
const originalAddLog = window.addLog;
window.addLog = (msg) => {
    window.addActivity(msg);
    if (originalAddLog) {
        const container = document.getElementById('live-logs');
        if (container) {
            const line = document.createElement('div');
            line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
            container.appendChild(line);
            container.scrollTop = container.scrollHeight;
        }
    }
};

// IPC Listeners
api.onScanUpdate(data => {
    const btn = document.getElementById('full-scan-btn');
    const progMini = document.getElementById('scan-progress-mini');
    const progBar = document.getElementById('scan-progress-bar');
    const progText = document.getElementById('scan-progress-text');

    if (data.status === 'SCANNING' || data.status === 'History') {
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg> Auditing...`;
        btn.disabled = true;
        progMini.style.display = 'flex';
        if (data.total > 0) {
            const pct = Math.round((data.count / data.total) * 100);
            progBar.style.width = pct + '%';
            progText.textContent = pct + '%';
        }
    } else if (data.status === 'Finished') {
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><path d="M12 8v4l3 3"></path></svg> Full Audit`;
        btn.disabled = false;
        progMini.style.display = 'none';
        api.setHistoryEnabled(false);
    }

    if (data.status === 'THREAT BLOCKED' || data.status === 'SPAM FILTERED' || data.status === 'Finished' || data.status === 'SCANNING' || data.status === 'INFO' || data.status === 'ERROR') {
        const prefix = data.status === 'THREAT BLOCKED' ? 'BLOCKED: ' : (data.status === 'SPAM FILTERED' ? 'SPAM: ' : (data.status === 'ERROR' ? 'CRITICAL: ' : ''));
        if (data.details) window.addActivity(prefix + data.details);
    }
});

api.onStatusSync(enabled => {
    const status = document.getElementById('engine-status');
    if (enabled) {
        status.className = 'badge badge-safe';
        status.textContent = 'Monitoring Active';
        window.addActivity("System: Security Protection has been ACTIVATED.");
    } else {
        status.className = 'badge badge-malicious';
        status.textContent = 'Protection Off';
        window.addActivity("System: Security Protection is currently DISABLED.", "danger");
    }
});

api.onStatsUpdate(data => window.AppState.updateStats(data));

api.onLiveLog(msg => window.addLog(msg));

api.onOutlookStatus(statusData => {
    const status = document.getElementById('engine-status');
    if (!status) return;
    const running = typeof statusData === 'object' ? !!statusData.running : !!statusData;
    const isStandby = typeof statusData === 'object' ? !!statusData.standby : !running;
    if (running) {
        status.className = 'badge badge-safe';
        status.textContent = 'Monitoring Active';
    } else if (isStandby) {
        status.className = 'badge badge-spam';
        status.textContent = 'Standby (Outlook Idle)';
    } else {
        status.className = 'badge badge-malicious';
        status.textContent = 'Protection Stopped';
    }
});

// Window Controls
window.winMinimize = () => api.minimizeWindow();
window.winMaximize = () => api.maximizeWindow ? api.maximizeWindow() : null;
window.winHide = () => api.hideWindow();
window.winClose = () => api.closeWindow ? api.closeWindow() : api.hideWindow();

async function init() {
    const stats = await api.getStats();
    window.AppState.updateStats({ full: true, stats: stats });
    try {
        const cfg = await api.getConfig();
        window.AppState.config = cfg;
        let localCompleted = false;
        try {
            localCompleted = localStorage.getItem('deskguard_first_run_completed') === 'true';
        } catch (e) {}

        if (!localCompleted && cfg && cfg.firstRun === true) {
            setTimeout(() => {
                openModal('first-run-wizard-modal');
            }, 300);
        }
    } catch (e) {}
}

window.useRecommendedSettings = async function() {
    try {
        try { localStorage.setItem('deskguard_first_run_completed', 'true'); } catch (e) {}
        const recommendedRubrics = {
            weights: { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 },
            toggles: { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true },
            spamThresholdPercent: 50
        };
        await api.setRubrics(recommendedRubrics);
        await api.setThreatIntelLevel(1);
        await api.setFirstRun(false);
        closeModal('first-run-wizard-modal');
        window.showNotification('Recommended protection activated: 8-engine security & zero-API threat intelligence enabled.');
        addLog('First-time setup completed: Recommended security policy applied.');
    } catch (err) {
        console.error('Wizard error:', err);
        try { localStorage.setItem('deskguard_first_run_completed', 'true'); } catch (e) {}
        closeModal('first-run-wizard-modal');
    }
};

window.customizeWizardSettings = async function() {
    try {
        try { localStorage.setItem('deskguard_first_run_completed', 'true'); } catch (e) {}
        await api.setFirstRun(false);
        closeModal('first-run-wizard-modal');
        setTimeout(() => {
            if (window.SecurityUI && window.SecurityUI.Engines && window.SecurityUI.Engines.open) {
                window.SecurityUI.Engines.open();
            } else {
                openSettings();
            }
        }, 350);
    } catch (err) {
        try { localStorage.setItem('deskguard_first_run_completed', 'true'); } catch (e) {}
        closeModal('first-run-wizard-modal');
    }
};

window.showNotification = function(msg, isError = false) {
    if (typeof window.addLog === 'function') {
        window.addLog(msg);
    }
    let toast = document.getElementById('app-notification-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'app-notification-toast';
        toast.style.cssText = 'position:fixed; bottom:20px; right:20px; z-index:99999; padding:12px 20px; border-radius:8px; font-size:0.85rem; font-weight:600; color:#fff; box-shadow:0 8px 24px rgba(0,0,0,0.5); transition:all 0.3s ease; pointer-events:none; opacity:0; transform:translateY(10px);';
        document.body.appendChild(toast);
    }
    toast.style.background = isError ? 'var(--danger, #d83b01)' : 'var(--accent, #0078d4)';
    toast.textContent = msg;
    toast.style.opacity = '1';
    toast.style.transform = 'translateY(0)';
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(10px)';
    }, 3500);
};

// Attach to globals for HTML usage
window.openModal = openModal;
window.closeModal = closeModal;
window.switchTab = switchTab;
window.switchCategory = switchCategory;
window.openSettings = openSettings;
window.closeSettings = closeSettings;
window.closeForensics = closeForensics;
window.requestFullScan = requestFullScan;
window.factoryReset = factoryReset;
window.handleOverlayClick = handleOverlayClick;
window.closeConfirm = closeConfirm;

init();

// Forensic modal wrapper
const originalOpenForensics = window.openForensics;
window.openForensics = (item) => {
    openModal('forensics-modal');
    if (window.renderForensics) window.renderForensics(item);
};

// Global Keyboard Navigation for Virtualized Lists and Modal Dismissal
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        const activeModal = document.querySelector('.modal-overlay.active');
        if (activeModal) {
            closeModal(activeModal.id);
            return;
        }
    }
    if (window.AppState.currentTab === 'dashboard' && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
        if (document.querySelector('.modal-overlay.active')) return;
        if (typeof window.handleListKeyNavigation === 'function') {
            window.handleListKeyNavigation(e);
        }
    }
});

// --- OFFLINE VAULT SEARCH & AUTOCOMPLETE TOOLBAR ---
window.vaultSearchActive = false;
let vaultSearchDebounceTimer = null;
let senderSuggestDebounceTimer = null;

window.executeVaultSearch = async function() {
    const searchInput = document.getElementById('vault-search-input');
    const fromInput = document.getElementById('vault-from-input');
    const scopeSelect = document.getElementById('vault-search-scope');
    const clearBtn = document.getElementById('vault-search-clear');

    const query = searchInput ? searchInput.value.trim() : '';
    const fromFilter = fromInput ? fromInput.value.trim() : '';
    const scope = scopeSelect ? scopeSelect.value : 'all';
    const category = window.AppState.currentCategory;

    const hasQuery = query.length > 0 || fromFilter.length > 0;
    if (clearBtn) clearBtn.style.display = hasQuery ? 'inline-block' : 'none';

    if (!hasQuery) {
        if (window.vaultSearchActive) {
            window.vaultSearchActive = false;
            window.AppState.renderDashboard(false);
        }
        return;
    }

    window.vaultSearchActive = true;

    try {
        const sortBy = window.sortOrder ? window.sortOrder.key : 'date';
        const sortOrder = (window.sortOrder && window.sortOrder.asc) ? 'asc' : 'desc';

        const res = await api.searchVault({
            query,
            scope,
            category,
            fromFilter,
            sortBy,
            sortOrder,
            limit: 500
        });

        if (res && res.ok && res.results) {
            const rows = res.results.rows || [];
            window.renderList(rows, category, false);
            const titleEl = document.getElementById('current-list-title');
            if (titleEl) {
                titleEl.textContent = `Vault Search (${res.results.total} matching)`;
            }
        }
    } catch (err) {
        console.error('[Vault Search Error]:', err);
    }
};

window.clearVaultSearch = function() {
    const searchInput = document.getElementById('vault-search-input');
    const fromInput = document.getElementById('vault-from-input');
    const clearBtn = document.getElementById('vault-search-clear');
    if (searchInput) searchInput.value = '';
    if (fromInput) fromInput.value = '';
    if (clearBtn) clearBtn.style.display = 'none';
    window.vaultSearchActive = false;
    window.AppState.renderDashboard(true);
};

function initVaultSearchToolbar() {
    const searchInput = document.getElementById('vault-search-input');
    const fromInput = document.getElementById('vault-from-input');
    const scopeSelect = document.getElementById('vault-search-scope');
    const datalist = document.getElementById('from-suggestions');

    if (searchInput) {
        searchInput.addEventListener('input', () => {
            clearTimeout(vaultSearchDebounceTimer);
            vaultSearchDebounceTimer = setTimeout(() => {
                window.executeVaultSearch();
            }, 250);
        });
    }

    if (scopeSelect) {
        scopeSelect.addEventListener('change', () => {
            if (window.vaultSearchActive || (searchInput && searchInput.value.trim())) {
                window.executeVaultSearch();
            }
        });
    }

    if (fromInput) {
        fromInput.addEventListener('input', () => {
            const val = fromInput.value.trim();
            // Autocomplete suggestions
            clearTimeout(senderSuggestDebounceTimer);
            senderSuggestDebounceTimer = setTimeout(async () => {
                if (val.length >= 1 && datalist) {
                    try {
                        const res = await api.getSenderSuggestions(val);
                        if (res && res.ok && Array.isArray(res.suggestions)) {
                            datalist.innerHTML = res.suggestions
                                .map(s => `<option value="${window.escapeHtml(s)}"></option>`)
                                .join('');
                        }
                    } catch (e) {}
                }
            }, 150);

            // Execute search filter
            clearTimeout(vaultSearchDebounceTimer);
            vaultSearchDebounceTimer = setTimeout(() => {
                window.executeVaultSearch();
            }, 250);
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initVaultSearchToolbar);
} else {
    initVaultSearchToolbar();
}
