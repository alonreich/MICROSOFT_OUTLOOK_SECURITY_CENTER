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
        document.getElementById('stat-malicious').textContent = this.stats.malicious.length;
        document.getElementById('stat-suspicious').textContent = this.stats.suspicious.length;
        document.getElementById('stat-spam').textContent = this.stats.spam.length;
        document.getElementById('stat-safe').textContent = this.stats.safe.length;
        
        const list = this.stats[this.currentCategory] || [];
        window.renderList(list, this.currentCategory, resetScroll);
        document.getElementById('current-list-title').textContent = 
            this.currentCategory.charAt(0).toUpperCase() + this.currentCategory.slice(1) + " Incidents";
    }
};

function switchTab(tabId) {
    if (window.AppState.currentTab === tabId) return;
    
    const oldTab = document.getElementById('tab-' + window.AppState.currentTab);
    const newTab = document.getElementById('tab-' + tabId);
    
    oldTab.classList.remove('active');
    setTimeout(() => {
        window.AppState.currentTab = tabId;
        document.querySelectorAll('.nav-item').forEach(p => p.classList.remove('active'));
        const navItem = document.querySelector(`.nav-item[data-tab="${tabId}"]`);
        if (navItem) navItem.classList.add('active');
        
        newTab.classList.add('active');
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
        async () => { await api.resetApp(); }
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
window.winHide = () => api.hideWindow();

async function init() {
    const stats = await api.getStats();
    window.AppState.updateStats({ full: true, stats: stats });
}

// Attach to globals for HTML usage
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

// Global Keyboard Navigation for Virtualized Lists
document.addEventListener('keydown', (e) => {
    if (window.AppState.currentTab === 'dashboard' && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
        if (document.querySelector('.modal-overlay.active')) return;
        if (typeof window.handleListKeyNavigation === 'function') {
            window.handleListKeyNavigation(e);
        }
    }
});
