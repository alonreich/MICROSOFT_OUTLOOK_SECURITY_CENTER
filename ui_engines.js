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
SecurityUI.Engines = (function() {
    const api = window.securityApi;
    let currentWeights = { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 };
    let currentToggles = { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true };
    let currentSpamThreshold = 50;
    let tooltipTimer = null;

    function stepWeight(key, delta) { 
        if (!currentToggles[key]) return; 
        let newVal = (currentWeights[key] || 0) + delta; 
        if (newVal < 1) newVal = 1; 
        if (newVal > 100) newVal = 100; 
        rebalanceWeights(key, newVal); 
        render(); 
    }

    function rebalanceWeights(key, newVal) {
        if (!currentToggles[key]) { currentWeights[key] = 0; } else {
            const oldVal = currentWeights[key] || 0; const diff = newVal - oldVal; currentWeights[key] = newVal;
            const activeKeys = Object.keys(currentWeights).filter(k => k !== key && currentToggles[k]);
            if (activeKeys.length > 0) { let rem = diff; activeKeys.forEach(k => { const share = Math.round(rem / activeKeys.length); currentWeights[k] -= share; rem -= share; }); currentWeights[activeKeys[0]] -= rem; }
        }
        Object.keys(currentWeights).forEach(k => { if (!currentToggles[k]) currentWeights[k] = 0; else if (currentWeights[k] < 1) currentWeights[k] = 1; });
        let total = 0; const activeKeys = Object.keys(currentWeights).filter(k => currentToggles[k]); activeKeys.forEach(k => total += currentWeights[k]);
        if (total !== 100 && activeKeys.length > 0) { const adjustKey = activeKeys.find(k => k !== key) || activeKeys[0]; currentWeights[adjustKey] += (100 - total); }
    }

    function render() {
        const container = document.getElementById('sensitivity-container');
        if (!container) return;
        const labels = { dmarc: 'DMARC AUTHENTICATION', alignment: 'SENDER ALIGNMENT', dkim: 'DKIM SIGNATURES', spf: 'SPF AUTHORIZATION', rdns: 'REVERSE DNS CHECK', body: 'ANTI-PHISHING SHIELD', heuristics: 'SMART HEURISTICS', rbl: 'GLOBAL BLACKLISTS' };
        const tips = {
            dmarc: "WHAT IT IS: A set of strict rules that prove the sender is really who they say they are. \n\nWHY USE IT: It's the strongest way to stop hackers from pretending to be your bank or boss. \n\nWHEN TO AVOID: If you send emails through a middle-man service that isn't set up perfectly, your own emails might get blocked.",
            alignment: "WHAT IT IS: Checks if the name on the outside of the envelope matches the name on the actual letter inside. \n\nWHY USE IT: It stops 'imposters' who use a real-looking name but a fake email address. \n\nWHEN TO AVOID: If you use mailing lists or newsletters, they sometimes change these names and might look like a mistake.",
            dkim: "WHAT IT IS: A digital 'wax seal' that proves the email was not changed by anyone after it was sent. \n\nWHY USE IT: It makes sure that links or files inside the email are exactly what the sender intended. \n\nWHEN TO AVOID: Only turn this off if you receive mail from very old systems that don't know how to 'seal' their emails.",
            spf: "WHAT IT IS: A list of 'Approved Post Offices' that a company uses to send their mail. \n\nWHY USE IT: It's a basic test that catches many simple scams and fake emails. \n\nWHEN TO AVOID: If a company forgets to update their list, their real emails might look fake. But usually, you should keep this ON.",
            rdns: "WHAT IT IS: Checks the 'Internet ID Card' of the computer that sent the mail to see if it's a real business or a hidden hacker bot. \n\nWHY USE IT: Most mass-spamming robots don't have a real ID card, so this blocks them instantly. \n\nWHEN TO AVOID: Some small or home-based businesses might not have a perfect ID card yet.",
            body: "WHAT IT IS: Scans the email for 'hidden tricks' like invisible text or links that try to fool your eyes. \n\nWHY USE IT: It catches hackers who try to hide their bad links inside what looks like a normal message. \n\nWHEN TO AVOID: If you get a lot of very fancy, colorful shopping ads, they might sometimes look 'tricky' to this engine.",
            heuristics: "WHAT IT IS: Looks for 'Red Flag' words that scammers love to use, like 'WINNER', 'BITCOIN', or 'URGENT'. \n\nWHY USE IT: It's great at catching classic 'Get Rich Quick' or 'You've Been Hacked' scams. \n\nWHEN TO AVOID: If your job involves money, law, or medicine, you might use these words normally, and this might flag them by mistake.",        
            rbl: "WHAT IT IS: A giant global 'Watch List' of known bad guys and spam servers. \n\nWHY USE IT: It's the fastest way to block millions of known spammers before they even reach you. \n\nWHEN TO AVOID: Almost never. This is one of the best tools for a clean inbox."
        };

        const tooltipEl = document.getElementById('tooltip');

        container.innerHTML = Object.keys(currentWeights).map(key => {
            const val = currentWeights[key]; const en = currentToggles[key];
            return `<div class="sensitivity-row" data-tip="${window.escapeHTML(tips[key])}" style="display:flex; align-items:center; gap:15px; opacity:${en?1:0.4}; cursor:help;"><input type="checkbox" class="w-tog" data-key="${key}" ${en?'checked':''}> <div style="flex:1.2; font-size:0.75rem; font-weight:bold;">${labels[key]}</div> <input type="range" class="w-slider" data-key="${key}" min="1" max="100" value="${val}" ${en?'':'disabled'} style="flex:2; cursor:pointer;"> <div style="display:flex; align-items:center; gap:8px;"> <button class="btn-step btn-minus" data-key="${key}">-</button> <div style="width:40px; text-align:center; font-family:monospace; color:var(--accent); font-weight:900;">${val}%</div> <button class="btn-step btn-plus" data-key="${key}">+</button> </div> </div>`;
        }).join('');

        document.querySelectorAll('.sensitivity-row').forEach(row => {
            row.onmouseenter = () => {
                if (!tooltipEl) return;
                const text = row.dataset.tip;
                clearTimeout(tooltipTimer);
                tooltipTimer = setTimeout(() => {
                    tooltipEl.textContent = text; tooltipEl.style.display = 'block';
                    const rect = row.getBoundingClientRect();
                    tooltipEl.style.left = Math.max(10, rect.left - 370) + 'px'; 
                    tooltipEl.style.top = Math.max(10, rect.top) + 'px';
                }, 400);
            };
            row.onmouseleave = () => { 
                clearTimeout(tooltipTimer); 
                if (tooltipEl) tooltipEl.style.display = 'none'; 
            };
        });

        document.querySelectorAll('.w-tog').forEach(c => c.onchange = () => { currentToggles[c.dataset.key] = c.checked; rebalanceWeights(c.dataset.key, c.checked ? 10 : 0); render(); });
        document.querySelectorAll('.w-slider').forEach(s => s.oninput = (e) => { rebalanceWeights(e.target.dataset.key, parseInt(e.target.value)); render(); });
        document.querySelectorAll('.btn-minus').forEach(b => b.onclick = () => stepWeight(b.dataset.key, -1));
        document.querySelectorAll('.btn-plus').forEach(b => b.onclick = () => stepWeight(b.dataset.key, 1));
    }

    async function open() { 
        const cfg = await api.getConfig(); 
        if (cfg && cfg.rubrics) { 
            currentWeights = { ...cfg.rubrics.weights }; 
            currentToggles = { ...cfg.rubrics.toggles }; 
            currentSpamThreshold = cfg.rubrics.spamThresholdPercent || 50; 
        } else {
            currentWeights = { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }; 
            currentToggles = { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }; 
            currentSpamThreshold = 50; 
        }
        const slider = document.getElementById('verdict-threshold-slider');
        if (slider) slider.value = currentSpamThreshold;
        const valEl = document.getElementById('verdict-threshold-val');
        if (valEl) valEl.textContent = currentSpamThreshold + '%';
        const descEl = document.getElementById('verdict-description');
        if (descEl) descEl.textContent = `Emails scoring ${currentSpamThreshold}% or lower will be marked as SPAM.`;
        render(); 
        if (typeof window.openModal === 'function') {
            window.openModal('sensitivity-modal');
        } else {
            const m = document.getElementById('sensitivity-modal');
            if (m) m.style.display = 'flex'; 
        }
    }

    function initListeners() {
        const btn = document.getElementById('sensitivity-btn');
        if (btn) btn.onclick = open;

        const closeBtn = document.getElementById('close-sensitivity');
        if (closeBtn) closeBtn.onclick = () => {
            if (typeof window.closeModal === 'function') window.closeModal('sensitivity-modal');
            else document.getElementById('sensitivity-modal').style.display = 'none';
        };

        const cancelBtn = document.getElementById('cancel-sensitivity');
        if (cancelBtn) cancelBtn.onclick = () => {
            if (typeof window.closeModal === 'function') window.closeModal('sensitivity-modal');
            else document.getElementById('sensitivity-modal').style.display = 'none';
        };

        const saveBtn = document.getElementById('save-sensitivity');
        if (saveBtn) saveBtn.onclick = async () => { 
            const cfg = await api.getConfig(); 
            await api.setRubrics({ ...(cfg ? cfg.rubrics : {}), weights: currentWeights, toggles: currentToggles, spamThresholdPercent: currentSpamThreshold }); 
            if (typeof window.showNotification === 'function') {
                window.showNotification('Anti-Spam Engine weights and thresholds updated.');
            }
            if (typeof window.closeModal === 'function') window.closeModal('sensitivity-modal');
            else document.getElementById('sensitivity-modal').style.display = 'none'; 
        };

        const slider = document.getElementById('verdict-threshold-slider');
        if (slider) slider.oninput = (e) => { 
            currentSpamThreshold = parseInt(e.target.value); 
            const valEl = document.getElementById('verdict-threshold-val');
            if (valEl) valEl.textContent = currentSpamThreshold + '%'; 
            const descEl = document.getElementById('verdict-description');
            if (descEl) descEl.textContent = `Emails scoring ${currentSpamThreshold}% or lower will be marked as SPAM.`; 
        };

        const resetBtn = document.getElementById('reset-sensitivity-defaults');
        if (resetBtn) resetBtn.onclick = () => { 
            currentWeights = { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }; 
            currentToggles = { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }; 
            currentSpamThreshold = 50; 
            const sl = document.getElementById('verdict-threshold-slider');
            if (sl) sl.value = 50;
            const vl = document.getElementById('verdict-threshold-val');
            if (vl) vl.textContent = '50%';
            const dc = document.getElementById('verdict-description');
            if (dc) dc.textContent = 'Emails scoring 50% or lower will be marked as SPAM.';
            render(); 
            if (typeof window.showNotification === 'function') {
                window.showNotification('Anti-spam engine weights reset to factory defaults.');
            }
        };
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initListeners);
    } else {
        setTimeout(initListeners, 50);
    }

    return { render: render, open: open };
})();
