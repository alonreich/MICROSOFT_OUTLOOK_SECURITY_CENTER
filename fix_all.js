const fs = require('fs');

function fixMainJs() {
    let code = fs.readFileSync('main.js', 'utf8');

    
    code = code.replace(
        /if \(msg\.type !== 'cmd'\) continue;/g,
        `if (msg.type === 'store-clear') { try { store.clear(); } catch(e){} continue; }\n                if (msg.type === 'store-set') { try { store.set(msg.key, msg.value); } catch(e) { logLine('STORE_ERROR', e.message, true); } continue; }\n                if (msg.type !== 'cmd') continue;`
    );

    code = code.replace(
        /async function safeStoreSet\(key, value\) \{ try \{ store\.set\(key, value\); return true; \} catch \(e\) \{ await logLine\('STORE_ERROR', e\.message, true\); return false; \} \}/g,
        `async function safeStoreSet(key, value) { if (!isServiceMode && uiPipeClient && !uiPipeClient.destroyed) { uiPipeClient.write(JSON.stringify({ type: 'store-set', key, value }) + '\\n'); return true; } try { store.set(key, value); return true; } catch (e) { await logLine('STORE_ERROR', e.message, true); return false; } }`
    );

    code = code.replace(
        /async function safeStoreMerge\(values\) \{ try \{ for \(const \[k, v\] of Object\.entries\(values\)\) \{ store\.set\(k, v\); \} return true; \} catch \(e\) \{ await logLine\('STORE_ERROR', e\.message, true\); return false; \} \}/g,
        `async function safeStoreMerge(values) { for (const [k, v] of Object.entries(values)) { await safeStoreSet(k, v); } return true; }`
    );

    const replacements = [
        ["store.set('enabled', v)", "await safeStoreSet('enabled', v)"],
        ["store.set('historyScanEnabled', v)", "await safeStoreSet('historyScanEnabled', v)"],
        ["store.set('vtApiKey', encryptKey(k))", "await safeStoreSet('vtApiKey', encryptKey(k))"],
        ["store.set('spamKeywords', k)", "await safeStoreSet('spamKeywords', k)"],
        ["store.set('rubrics', r)", "await safeStoreSet('rubrics', r)"],
        ["store.set('whitelist', w)", "await safeStoreSet('whitelist', w)"],
        ["store.set('blacklist', b)", "await safeStoreSet('blacklist', b)"],
        ["store.set('schedule', s)", "await safeStoreSet('schedule', s)"],
        ["store.set('columnWidths', w)", "await safeStoreSet('columnWidths', w)"],
        ["store.set('processedIds', [])", "await safeStoreSet('processedIds', [])"],
        ["store.set('stats', { spam: [], safe: [], malicious: [], suspicious: [] })", "await safeStoreSet('stats', { spam: [], safe: [], malicious: [], suspicious: [] })"],
        ["store.set('totals', { spam: 0, safe: 0, malicious: 0, suspicious: 0 })", "await safeStoreSet('totals', { spam: 0, safe: 0, malicious: 0, suspicious: 0 })"],
        ["store.set('stats', allStats)", "await safeStoreSet('stats', allStats)"],
        ["store.set(`totals.${src}`, Math.max(0, (store.get(`totals.${src}`) || 0) - 1))", "await safeStoreSet(`totals.${src}`, Math.max(0, (store.get(`totals.${src}`) || 0) - 1))"],
        ["store.set('stats.safe', safeStats.length > 1000 ? safeStats.slice(-1000) : safeStats)", "await safeStoreSet('stats.safe', safeStats.length > 1000 ? safeStats.slice(-1000) : safeStats)"],
        ["store.set('totals.safe', (store.get('totals.safe') || 0) + 1)", "await safeStoreSet('totals.safe', (store.get('totals.safe') || 0) + 1)"],
        ["store.set('stats.spam', spamStats.length > 1000 ? spamStats.slice(-1000) : spamStats)", "await safeStoreSet('stats.spam', spamStats.length > 1000 ? spamStats.slice(-1000) : spamStats)"],
        ["store.set('totals.spam', (store.get('totals.spam') || 0) + 1)", "await safeStoreSet('totals.spam', (store.get('totals.spam') || 0) + 1)"],
        ["store.set('whitelist', wl); store.set('blacklist', bl);", "await safeStoreSet('whitelist', wl); await safeStoreSet('blacklist', bl);"],
        ["store.set('windowBounds'", "safeStoreSet('windowBounds'"],
        ["store.clear()", "if(isServiceMode) { store.clear(); } else { if(uiPipeClient && !uiPipeClient.destroyed) { uiPipeClient.write(JSON.stringify({type:'store-clear'})+'\\n'); } }"]
    ];

    replacements.forEach(([from, to]) => {
        code = code.split(from).join(to);
    });

    fs.writeFileSync('main.js', code);
    console.log("main.js fixed");
}

function fixIndexHtml() {
    let code = fs.readFileSync('index.html', 'utf8');

    
    const oldLogic = `    function rebalanceWeights(key, newVal) {
        if (!currentToggles[key]) {
            currentWeights[key] = 0;
        } else {
            const oldVal = currentWeights[key] || 0;
            const diff = newVal - oldVal;
            currentWeights[key] = newVal;
            const activeKeys = Object.keys(currentWeights).filter(k => k !== key && currentToggles[k]);
            if (activeKeys.length > 0) {
                let rem = diff;
                activeKeys.forEach(k => {
                    const share = Math.round(rem / activeKeys.length);
                    currentWeights[k] -= share;
                    rem -= share;
                });
                currentWeights[activeKeys[0]] -= rem;
            }
        }
        
        Object.keys(currentWeights).forEach(k => {
            if (!currentToggles[k]) currentWeights[k] = 0;
            else if (currentWeights[k] < 1) currentWeights[k] = 1;
        });

        let total = 0;
        const activeKeys = Object.keys(currentWeights).filter(k => currentToggles[k]);
        activeKeys.forEach(k => total += currentWeights[k]);
        
        if (total !== 100 && activeKeys.length > 0) {
            const adjustKey = activeKeys.find(k => k !== key) || activeKeys[0];
            currentWeights[adjustKey] += (100 - total);
        }
    }`;

    const newLogic = `    function rebalanceWeights(key, newVal) {
        if (!currentToggles[key]) {
            currentWeights[key] = 0;
            newVal = 0;
        } else {
            if (newVal < 1) newVal = 1;
            if (newVal > 99) newVal = 99;
            currentWeights[key] = newVal;
        }

        const activeKeys = Object.keys(currentWeights).filter(k => k !== key && currentToggles[k]);
        if (activeKeys.length === 0) {
            if (currentToggles[key]) currentWeights[key] = 100;
            return;
        }

        let remainingTarget = 100 - newVal;
        let currentOthersTotal = activeKeys.reduce((sum, k) => sum + currentWeights[k], 0);

        if (currentOthersTotal === 0) {
            const share = Math.floor(remainingTarget / activeKeys.length);
            activeKeys.forEach(k => currentWeights[k] = share);
        } else {
            activeKeys.forEach(k => {
                currentWeights[k] = Math.max(1, Math.round((currentWeights[k] / currentOthersTotal) * remainingTarget));
            });
        }

        let newTotal = activeKeys.reduce((sum, k) => sum + currentWeights[k], 0) + newVal;
        if (newTotal !== 100) {
            let diff = 100 - newTotal;
            let maxKey = activeKeys.reduce((max, k) => currentWeights[k] > currentWeights[max] ? k : max, activeKeys[0]);
            currentWeights[maxKey] += diff;
            if (currentWeights[maxKey] < 1) currentWeights[maxKey] = 1;
        }
    }`;

    code = code.replace(oldLogic, newLogic);
    fs.writeFileSync('index.html', code);
    console.log("index.html fixed");
}

fixMainJs();
fixIndexHtml();
