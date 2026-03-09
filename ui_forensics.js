document.body.insertAdjacentHTML('beforeend', `
<div id="forensic-modal" class="modal">
    <div class="modal-content" style="width: 1200px; font-family: Arial, sans-serif; font-size: 11pt;">
        <div class="modal-header">EMAIL FORENSIC DETAILS</div>
        <div id="forensic-fields" style="display: flex; flex-direction: column; gap: 0; padding: 10px;"></div>
        <div style="margin-top: 30px; display: flex; justify-content: center; flex-wrap: wrap; gap: 15px; padding-top: 20px; border-top: 1px solid var(--border);">
            <button id="forensic-rfc" class="btn-ui">SHOW EMAIL HEADERS (RFC822)</button>
            <button id="forensic-read" class="btn-ui"><svg style="width:16px;height:16px;fill:currentColor" viewBox="0 0 24 24"><path d="M20,8L12,13L4,8V6L12,11L20,6M20,4H4C2.89,4 2,4.89 2,6V18A2,2 0 0,0 4,20H20A2,2 0 0,0 22,18V6C22,4.89 21.1,4 20,4Z" /></svg>&nbsp;&nbsp;READ EMAIL CONTENT&nbsp;&nbsp;<svg style="width:16px;height:16px;fill:currentColor" viewBox="0 0 24 24"><path d="M20,8L12,13L4,8V6L12,11L20,6M20,4H4C2.89,4 2,4.89 2,6V18A2,2 0 0,0 4,20H20A2,2 0 0,0 22,18V6C22,4.89 21.1,4 20,4Z" /></svg></button>
            <button id="forensic-prev" class="btn-ui">PREVIOUS <<</button><button id="forensic-next" class="btn-ui">NEXT >></button><button id="forensic-close" class="btn-ui danger">CLOSE</button>
        </div>
    </div>
</div>
`);

(function() {
    const api = window.securityApi;
    window.currentForensicId = null;

    window.showForensics = function(id) {
        window.currentForensicId = id; 
        const activeCard = document.querySelector('.stat-card.active'); 
        const cat = activeCard ? activeCard.id.replace('card-','') : 'malicious'; 
        const items = (window.stats[cat] || []).slice().reverse(); 
        const index = items.findIndex(i => i.entryId === id); 
        const item = items[index]; 
        if (!item) return;
        
        const ts = item.timestamp || ""; 
        const [d, t] = ts.includes(' ') ? ts.split(' ') : [ts, ""];
        const sub = item.subject || item.details || "No Subject";
        
        const fields = [ 
            { l: 'DATE', v: d }, { l: 'TIME', v: t }, { l: 'FROM', v: item.sender }, { l: 'TO', v: item.to || 'N/A' }, 
            { l: 'CC', v: item.cc || 'N/A' }, { l: 'SUBJECT', v: sub }, { l: 'IP ADDRESS', v: item.ip || 'N/A' }, 
            { l: 'VERDICT', v: cat.toUpperCase() }, { l: 'SCORE', v: Math.round(item.score || 0) + '%' }, 
            { l: 'ACTION', v: item.action || 'None' }, { l: 'REASONING', v: item.tier || 'None' } 
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
        const items = (window.stats[cat] || []).slice().reverse(); 
        const index = items.findIndex(i => i.entryId === window.currentForensicId); 
        if (index < items.length - 1) window.showForensics(items[index + 1].entryId); 
    };

    document.getElementById('forensic-prev').onclick = () => { 
        const activeCard = document.querySelector('.stat-card.active'); 
        const cat = activeCard ? activeCard.id.replace('card-','') : 'malicious'; 
        const items = (window.stats[cat] || []).slice().reverse(); 
        const index = items.findIndex(i => i.entryId === window.currentForensicId); 
        if (index > 0) window.showForensics(items[index - 1].entryId); 
    };
})();

