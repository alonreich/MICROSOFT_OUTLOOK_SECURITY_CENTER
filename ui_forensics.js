(function() {
    const api = window.securityApi;
    const store = window.AppStore;
    window.currentForensicId = null;

    window.showForensics = function(id) {
        window.currentForensicId = id; 
        const activeCard = document.querySelector('.stat-card.active'); 
        const cat = activeCard ? activeCard.id.replace('card-','') : 'malicious'; 
        const items = (store.stats[cat] || []); 
        const index = items.findIndex(i => i.entryId === id || i.fingerprint === id); 
        const item = items[index]; 
        if (!item) return;
        
        const ts = item.timestamp || ""; 
        const [d, t] = ts.includes(' ') ? ts.split(' ') : [ts, ""];
        const sub = item.subject || item.details || "No Subject";
        
        const fields = [
            { l: 'DATE', v: d }, { l: 'TIME', v: t }, { l: 'FROM', v: item.sender }, { l: 'TO', v: item.to || 'N/A' },
            { l: 'CC', v: item.cc || 'N/A' }, { l: 'SUBJECT', v: sub }, { l: 'IP ADDRESS', v: item.ip || 'N/A' },
            { l: 'VERDICT', v: cat.toUpperCase() }, { l: 'SCORE', v: Math.round(item.score || 0) + '%' },
            { l: 'ACTION', v: item.action || 'None' }, { l: 'REASONING', v: item.tier || 'None' },
            { l: 'SECURITY FINGERPRINT', v: item.fingerprint || item.entryId || 'N/A' }
        ];        
        document.getElementById('forensic-fields').innerHTML = fields.map(f => `<div style="display:flex; align-items:center; gap:10px; padding-bottom:15px; margin-bottom:15px; border-bottom:1px solid rgba(255,255,255,0.05); user-select: text !important; -webkit-user-select: text !important;"><span style="font-weight:bold; min-width:180px; color:var(--muted); font-size:0.75rem; text-transform:uppercase; text-align:right; padding-right:20px; user-select: text !important; -webkit-user-select: text !important; cursor: text;">${f.l}:</span><span style="flex:1; color:#fff; font-size:0.85rem; word-break:break-all; user-select: text !important; -webkit-user-select: text !important; cursor: text;">${window.escapeHTML(f.v)}</span><div style="width:16px; height:16px; fill:var(--accent); cursor:pointer; display:flex; align-items:center; justify-content:center; user-select: none;" onclick="window.copyToClipboard('${window.escapeHTML(f.v)}')"><svg viewBox="0 0 24 24"><path d="M19,21H8V7H19M19,5H8A2,2 0 0,0 6,7V21A2,2 0 0,0 8,23H19A2,2 0 0,0 21,21V7A2,2 0 0,0 19,5M16,1H4A2,2 0 0,0 2,3V17H4V3H16V1Z" /></svg></div></div>`).join('');
        
        document.getElementById('forensic-modal').style.display = 'flex';
        document.getElementById('forensic-prev').disabled = index === 0;
        document.getElementById('forensic-next').disabled = index === items.length - 1;
    }

    document.getElementById('forensic-close').onclick = () => document.getElementById('forensic-modal').style.display = 'none';

    document.getElementById('forensic-rfc').onclick = async () => { 
        const data = await api.getForensics(window.currentForensicId); 
        const win = window.open('', '_blank', 'width=950,height=900'); 
        win.document.write(`<html><body style="background:#0a0e1c;color:#88c0d0;font-family:monospace;padding:20px;white-space:pre-wrap;word-break:break-all;">${window.escapeHTML(data.fullHeaders)}</body></html>`); 
    };

    document.getElementById('forensic-read').onclick = async () => { 
        const data = await api.getForensics(window.currentForensicId); 
        const win = window.open('', '_blank', 'width=950,height=900'); 
        win.document.write(`<html><body style="background:#0a0e1c;color:#e1e4e8;font-family:sans-serif;padding:20px;white-space:pre-wrap;">${window.escapeHTML(data.body)}</body></html>`); 
    };

    document.getElementById('forensic-next').onclick = () => { 
        const activeCard = document.querySelector('.stat-card.active'); 
        const cat = activeCard ? activeCard.id.replace('card-','') : 'malicious'; 
        const items = (store.stats[cat] || []); 
        const index = items.findIndex(i => (i.fingerprint || i.entryId) === window.currentForensicId); 
        if (index < items.length - 1) window.showForensics(items[index + 1].fingerprint || items[index + 1].entryId); 
    };

    document.getElementById('forensic-prev').onclick = () => { 
        const activeCard = document.querySelector('.stat-card.active'); 
        const cat = activeCard ? activeCard.id.replace('card-','') : 'malicious'; 
        const items = (store.stats[cat] || []); 
        const index = items.findIndex(i => (i.fingerprint || i.entryId) === window.currentForensicId); 
        if (index > 0) window.showForensics(items[index - 1].fingerprint || items[index - 1].entryId); 
    };
})();

