document.body.insertAdjacentHTML('beforeend', `
<div id="sensitivity-modal" class="modal">
    <div class="modal-content" style="width: 650px;">
        <div class="modal-header">ANTI-SPAM ENGINES</div>
        <div id="sensitivity-container" style="display: flex; flex-direction: column; gap: 12px; padding: 10px; overflow-y: auto;"></div>
        <div style="margin-top: 20px; padding: 15px; border-top: 2px solid var(--border); background: rgba(142, 117, 255, 0.05); border-radius: 10px;">
            <div style="text-align: center; margin-bottom: 10px;"><div style="color: var(--accent); font-weight: 800; font-size: 0.85rem;">SPAM VERDICT AGGRESSIVENESS</div></div>
            <div style="display: flex; align-items: center; gap: 15px;"><div style="font-size: 0.7rem; font-weight: 800; color: var(--ok);">RELAXED</div><input type="range" id="verdict-threshold-slider" min="1" max="99" value="50" style="flex: 1; cursor: pointer;"><div style="font-size: 0.7rem; font-weight: 800; color: var(--danger);">STRICT</div></div>
            <div style="text-align: center; margin-top: 10px;"><span id="verdict-threshold-val" style="font-family: monospace; font-size: 1.4rem; font-weight: 900; color: var(--accent);">50%</span><div id="verdict-description" style="font-size: 0.65rem; color: var(--muted);">Emails scoring 50% or lower will be marked as SPAM.</div></div>
        </div>
        <div style="margin-top: 20px; display: flex; justify-content: center; gap: 15px;"><button id="save-sensitivity" class="btn-ui success">APPLY WEIGHTS</button><button id="reset-sensitivity-defaults" class="btn-ui">RESET DEFAULTS</button><button id="close-sensitivity" class="btn-ui danger">CANCEL</button></div>
    </div>
</div>
`);

(function() {
    const api = window.securityApi;
    const tooltipEl = document.getElementById('tooltip');
    let currentWeights = {}, currentToggles = {}, currentSpamThreshold = 50;
    let tooltipTimer = null;

    window.stepWeight = (key, delta) => { 
        if (!currentToggles[key]) return; 
        let newVal = currentWeights[key] + delta; 
        if (newVal < 1) newVal = 1; 
        if (newVal > 100) newVal = 100; 
        rebalanceWeights(key, newVal); 
        renderSensitivityUI(); 
    };

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

    function renderSensitivityUI() {
        const container = document.getElementById('sensitivity-container');
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

        container.innerHTML = Object.keys(currentWeights).map(key => {
            const val = currentWeights[key]; const en = currentToggles[key];
            return `<div class="sensitivity-row" data-tip="${window.escapeHTML(tips[key])}" style="display:flex; align-items:center; gap:15px; opacity:${en?1:0.4}; cursor:help;"><input type="checkbox" class="w-tog" data-key="${key}" ${en?'checked':''}> <div style="flex:1.2; font-size:0.75rem; font-weight:bold;">${labels[key]}</div> <input type="range" class="w-slider" data-key="${key}" min="1" max="100" value="${val}" ${en?'':'disabled'} style="flex:2; cursor:pointer;"> <div style="display:flex; align-items:center; gap:8px;"> <button class="btn-step" onclick="stepWeight('${key}',-1)" ${en?'':'disabled'}>-</button> <div style="width:40px; text-align:center; font-family:monospace; color:var(--accent); font-weight:900;">${val}%</div> <button class="btn-step" onclick="stepWeight('${key}',1)" ${en?'':'disabled'}>+</button> </div> </div>`;
        }).join('');

        document.querySelectorAll('.sensitivity-row').forEach(row => {
            row.onmouseenter = () => {
                const text = row.dataset.tip;
                clearTimeout(tooltipTimer);
                tooltipTimer = setTimeout(() => {
                    tooltipEl.textContent = text; tooltipEl.style.display = 'block';
                    const rect = row.getBoundingClientRect();
                    tooltipEl.style.left = (rect.left - 370) + 'px'; tooltipEl.style.top = rect.top + 'px';
                }, 800);
            };
            row.onmouseleave = () => { clearTimeout(tooltipTimer); tooltipEl.style.display = 'none'; };
        });

        document.querySelectorAll('.w-tog').forEach(c => c.onchange = () => { currentToggles[c.dataset.key] = c.checked; rebalanceWeights(c.dataset.key, c.checked ? 10 : 0); renderSensitivityUI(); });
        document.querySelectorAll('.w-slider').forEach(s => s.oninput = (e) => { rebalanceWeights(e.target.dataset.key, parseInt(e.target.value)); renderSensitivityUI(); });
    }

    document.getElementById('sensitivity-btn').onclick = async () => { 
        const cfg = await api.getConfig(); 
        if(cfg.rubrics){ 
            currentWeights = { ...cfg.rubrics.weights }; 
            currentToggles = { ...cfg.rubrics.toggles }; 
            currentSpamThreshold = cfg.rubrics.spamThresholdPercent || 50; 
        } 
        document.getElementById('verdict-threshold-slider').value = currentSpamThreshold;
        document.getElementById('verdict-threshold-val').textContent = currentSpamThreshold + '%';
        document.getElementById('verdict-description').textContent = `Emails scoring ${currentSpamThreshold}% or lower will be marked as SPAM.`;
        renderSensitivityUI(); 
        document.getElementById('sensitivity-modal').style.display = 'flex'; 
    };

    document.getElementById('close-sensitivity').onclick = () => document.getElementById('sensitivity-modal').style.display = 'none';

    document.getElementById('save-sensitivity').onclick = async () => { 
        const cfg = await api.getConfig(); 
        api.setRubrics({ ...cfg.rubrics, weights: currentWeights, toggles: currentToggles, spamThresholdPercent: currentSpamThreshold }); 
        document.getElementById('sensitivity-modal').style.display = 'none'; 
    };

    document.getElementById('verdict-threshold-slider').oninput = (e) => { 
        currentSpamThreshold = parseInt(e.target.value); 
        document.getElementById('verdict-threshold-val').textContent = currentSpamThreshold + '%'; 
        document.getElementById('verdict-description').textContent = `Emails scoring ${currentSpamThreshold}% or lower will be marked as SPAM.`; 
    };

    document.getElementById('reset-sensitivity-defaults').onclick = () => { 
        currentWeights = { dmarc: 13, alignment: 10, dkim: 7, spf: 25, rdns: 15, body: 10, heuristics: 10, rbl: 10 }; 
        currentToggles = { dmarc: true, alignment: true, dkim: true, spf: true, rdns: true, body: true, heuristics: true, rbl: true }; 
        currentSpamThreshold = 50; 
        document.getElementById('verdict-threshold-slider').value = 50;
        document.getElementById('verdict-threshold-val').textContent = '50%';
        document.getElementById('verdict-description').textContent = 'Emails scoring 50% or lower will be marked as SPAM.';
        renderSensitivityUI(); 
    };
})();

