// DeskGuard Offline High-Performance Email Vault
// Powered by native node:sqlite with SQLite FTS5 Inverted Index

const { DatabaseSync } = require('node:sqlite');
const path = require('path');
const fs = require('fs');

class VaultDB {
    constructor() {
        this.db = null;
        this.insertStmt = null;
        this.insertFtsStmt = null;
        this.deleteFtsStmt = null;
    }

    init(dbDir) {
        if (this.db) return;
        if (!fs.existsSync(dbDir)) {
            fs.mkdirSync(dbDir, { recursive: true });
        }
        const dbPath = path.join(dbDir, 'deskguard_vault.db');
        this.db = new DatabaseSync(dbPath);

        // Performance & Reliability pragmas: WAL mode for non-blocking concurrent writes
        this.db.exec(`
            PRAGMA journal_mode = WAL;
            PRAGMA synchronous = NORMAL;
            PRAGMA cache_size = -64000; -- 64MB cache
            PRAGMA temp_store = MEMORY;

            CREATE TABLE IF NOT EXISTS emails (
                id TEXT PRIMARY KEY,
                entry_id TEXT NOT NULL,
                store_id TEXT,
                fingerprint TEXT UNIQUE,
                subject TEXT,
                sender TEXT,
                recipient TEXT,
                cc TEXT,
                date TEXT,
                time TEXT,
                timestamp TEXT,
                body TEXT,
                headers TEXT,
                verdict TEXT,
                score REAL DEFAULT 100,
                tier TEXT,
                attachments_json TEXT,
                unread INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_emails_date ON emails(date DESC, time DESC);
            CREATE INDEX IF NOT EXISTS idx_emails_sender ON emails(sender);
            CREATE INDEX IF NOT EXISTS idx_emails_recipient ON emails(recipient);
            CREATE INDEX IF NOT EXISTS idx_emails_verdict ON emails(verdict);
            CREATE INDEX IF NOT EXISTS idx_emails_fingerprint ON emails(fingerprint);

            CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
                subject,
                body,
                sender,
                recipient,
                attachments,
                content='emails',
                content_rowid='rowid'
            );

            CREATE TRIGGER IF NOT EXISTS emails_ai AFTER INSERT ON emails BEGIN
                INSERT INTO emails_fts(rowid, subject, body, sender, recipient, attachments)
                VALUES (new.rowid, new.subject, new.body, new.sender, new.recipient, new.attachments_json);
            END;

            CREATE TRIGGER IF NOT EXISTS emails_ad AFTER DELETE ON emails BEGIN
                INSERT INTO emails_fts(emails_fts, rowid, subject, body, sender, recipient, attachments)
                VALUES ('delete', old.rowid, old.subject, old.body, old.sender, old.recipient, old.attachments_json);
            END;

            CREATE TRIGGER IF NOT EXISTS emails_au AFTER UPDATE ON emails BEGIN
                INSERT INTO emails_fts(emails_fts, rowid, subject, body, sender, recipient, attachments)
                VALUES ('delete', old.rowid, old.subject, old.body, old.sender, old.recipient, old.attachments_json);
                INSERT INTO emails_fts(rowid, subject, body, sender, recipient, attachments)
                VALUES (new.rowid, new.subject, new.body, new.sender, new.recipient, new.attachments_json);
            END;
        `);

        this.insertStmt = this.db.prepare(`
            INSERT INTO emails (
                id, entry_id, store_id, fingerprint, subject, sender, recipient, cc,
                date, time, timestamp, body, headers, verdict, score, tier,
                attachments_json, unread
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?
            )
            ON CONFLICT(fingerprint) DO UPDATE SET
                entry_id = excluded.entry_id,
                verdict = excluded.verdict,
                score = excluded.score,
                tier = excluded.tier,
                unread = excluded.unread;
        `);
    }

    insertEmail(item) {
        if (!this.db || !item) return;
        try {
            const id = item.entryId || item.fingerprint || `msg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
            const entryId = item.entryId || id;
            const storeId = item.storeId || '';
            const fingerprint = item.fingerprint || entryId;
            const subject = item.subject || item.details || '(No Subject)';
            const sender = item.sender || item.from || 'Unknown';
            const recipient = item.to || item.recipient || '';
            const cc = item.cc || '';
            
            // Extract & normalize date and time
            let dateStr = item.date || '';
            let timeStr = item.time || '';
            let tsStr = item.timestamp || '';
            
            if (!dateStr || !timeStr) {
                const parsedDate = item.timestamp ? new Date(item.timestamp) : new Date();
                if (!isNaN(parsedDate.getTime())) {
                    if (!dateStr) dateStr = parsedDate.toISOString().slice(0, 10);
                    if (!timeStr) timeStr = parsedDate.toTimeString().slice(0, 8);
                    if (!tsStr) tsStr = `${dateStr} ${timeStr}`;
                } else {
                    const now = new Date();
                    dateStr = dateStr || now.toISOString().slice(0, 10);
                    timeStr = timeStr || now.toTimeString().slice(0, 8);
                    tsStr = tsStr || `${dateStr} ${timeStr}`;
                }
            }

            // Decode body if base64 encoded
            let bodyText = item.body || '';
            if (bodyText && typeof bodyText === 'string') {
                const clean = bodyText.trim();
                if (/^[A-Za-z0-9+/]+={0,2}$/.test(clean) && clean.length % 4 === 0 && clean.length > 20 && !clean.includes(' ')) {
                    try {
                        const decoded = Buffer.from(clean, 'base64').toString('utf8');
                        if (decoded && !decoded.includes('\ufffd')) bodyText = decoded;
                    } catch {}
                }
            }

            // Decode headers if base64 encoded
            let headersText = item.fullHeaders || item.headers || '';
            if (headersText && typeof headersText === 'string') {
                const clean = headersText.trim();
                if (/^[A-Za-z0-9+/]+={0,2}$/.test(clean) && clean.length % 4 === 0 && clean.length > 20 && !clean.includes(' ')) {
                    try {
                        const decodedH = Buffer.from(clean, 'base64').toString('utf8');
                        if (decodedH && !decodedH.includes('\ufffd')) headersText = decodedH;
                    } catch {}
                }
            }

            const verdict = item.verdict || 'Safe';
            const score = typeof item.score === 'number' ? item.score : 100;
            const tier = item.tier || '';
            
            let attsJson = '[]';
            if (item.attachments) {
                attsJson = typeof item.attachments === 'string' ? item.attachments : JSON.stringify(item.attachments);
            }

            const unread = item.unread ? 1 : 0;

            this.insertStmt.run(
                id, entryId, storeId, fingerprint, subject, sender, recipient, cc,
                dateStr, timeStr, tsStr, bodyText, headersText, verdict, score, tier,
                attsJson, unread
            );
        } catch (err) {
            console.error('[VaultDB] insertEmail error:', err.message);
        }
    }

    searchEmails(params = {}) {
        if (!this.db) return { total: 0, rows: [] };

        const {
            query = '',
            scope = 'all', // 'all', 'attachments', 'body', 'subject', 'from', 'to'
            category = 'all', // 'all', 'malicious', 'suspicious', 'spam', 'safe'
            fromFilter = '',
            toFilter = '',
            limit = 100,
            offset = 0,
            sortBy = 'date',
            sortOrder = 'desc'
        } = params;

        const conditions = [];
        const bindings = [];

        // Category filter
        if (category && category !== 'all') {
            conditions.push('LOWER(e.verdict) LIKE ?');
            bindings.push(`%${category.toLowerCase()}%`);
        }

        // Dedicated From filter
        if (fromFilter && fromFilter.trim().length > 0) {
            conditions.push('e.sender LIKE ?');
            bindings.push(`%${fromFilter.trim()}%`);
        }

        // Dedicated To filter
        if (toFilter && toFilter.trim().length > 0) {
            conditions.push('e.recipient LIKE ?');
            bindings.push(`%${toFilter.trim()}%`);
        }

        // Full-Text & Wildcard Query Logic
        const trimmedQ = query.trim();
        let useFTS = false;
        let ftsQuery = '';

        if (trimmedQ.length > 0) {
            if (scope === 'attachments') {
                // Wildcard substring search inside attachments JSON (filenames, types, hashes)
                conditions.push('e.attachments_json LIKE ?');
                bindings.push(`%${trimmedQ.replace(/\*/g, '%')}%`);
            } else if (scope === 'body') {
                useFTS = true;
                const cleanTerm = trimmedQ.replace(/['"*]/g, '');
                ftsQuery = `body: ${cleanTerm}*`;
            } else if (scope === 'subject') {
                useFTS = true;
                const cleanTerm = trimmedQ.replace(/['"*]/g, '');
                ftsQuery = `subject: ${cleanTerm}*`;
            } else {
                // Omni-search ('all'): check wildcard in FTS5 across all indexed columns
                useFTS = true;
                const cleanTerm = trimmedQ.replace(/['"*]/g, '');
                ftsQuery = cleanTerm ? `${cleanTerm}*` : '';
            }
        }

        let sql = '';
        let countSql = '';

        const orderDirection = sortOrder.toLowerCase() === 'asc' ? 'ASC' : 'DESC';
        const allowedSortCols = {
            date: 'e.date ' + orderDirection + ', e.time ' + orderDirection,
            time: 'e.time ' + orderDirection,
            subject: 'e.subject ' + orderDirection,
            sender: 'e.sender ' + orderDirection,
            recipient: 'e.recipient ' + orderDirection,
            verdict: 'e.verdict ' + orderDirection,
            score: 'e.score ' + orderDirection
        };
        const orderClause = allowedSortCols[sortBy] || 'e.date DESC, e.time DESC';

        if (useFTS && ftsQuery) {
            const ftsWhere = ['emails_fts MATCH ?', ...conditions].join(' AND ');
            const ftsBindings = [ftsQuery, ...bindings];

            sql = `
                SELECT e.* 
                FROM emails_fts f
                JOIN emails e ON f.rowid = e.rowid
                WHERE ${ftsWhere}
                ORDER BY ${orderClause}
                LIMIT ? OFFSET ?;
            `;
            countSql = `
                SELECT COUNT(*) as cnt
                FROM emails_fts f
                JOIN emails e ON f.rowid = e.rowid
                WHERE ${ftsWhere};
            `;

            try {
                const countRow = this.db.prepare(countSql).get(...ftsBindings);
                const total = countRow ? countRow.cnt : 0;
                const rows = this.db.prepare(sql).all(...ftsBindings, limit, offset);
                return { total, rows: rows.map(r => this.formatRow(r)) };
            } catch (err) {
                console.warn('[VaultDB] FTS fallback to LIKE search:', err.message);
                // Fallback to standard LIKE matching if FTS syntax error
                return this.fallbackLikeSearch(trimmedQ, scope, conditions, bindings, orderClause, limit, offset);
            }
        } else {
            const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
            sql = `
                SELECT e.* 
                FROM emails e
                ${whereClause}
                ORDER BY ${orderClause}
                LIMIT ? OFFSET ?;
            `;
            countSql = `
                SELECT COUNT(*) as cnt 
                FROM emails e
                ${whereClause};
            `;

            const countRow = this.db.prepare(countSql).get(...bindings);
            const total = countRow ? countRow.cnt : 0;
            const rows = this.db.prepare(sql).all(...bindings, limit, offset);
            return { total, rows: rows.map(r => this.formatRow(r)) };
        }
    }

    fallbackLikeSearch(query, scope, baseConditions, baseBindings, orderClause, limit, offset) {
        const conds = [...baseConditions];
        const binds = [...baseBindings];
        const likePattern = `%${query.replace(/\*/g, '%')}%`;

        if (scope === 'subject') {
            conds.push('e.subject LIKE ?');
            binds.push(likePattern);
        } else if (scope === 'body') {
            conds.push('e.body LIKE ?');
            binds.push(likePattern);
        } else if (scope === 'attachments') {
            conds.push('e.attachments_json LIKE ?');
            binds.push(likePattern);
        } else {
            conds.push('(e.subject LIKE ? OR e.body LIKE ? OR e.sender LIKE ? OR e.recipient LIKE ? OR e.attachments_json LIKE ?)');
            binds.push(likePattern, likePattern, likePattern, likePattern, likePattern);
        }

        const whereClause = conds.length > 0 ? `WHERE ${conds.join(' AND ')}` : '';
        const countSql = `SELECT COUNT(*) as cnt FROM emails e ${whereClause};`;
        const countRow = this.db.prepare(countSql).get(...binds);
        const total = countRow ? countRow.cnt : 0;

        const sql = `SELECT e.* FROM emails e ${whereClause} ORDER BY ${orderClause} LIMIT ? OFFSET ?;`;
        const rows = this.db.prepare(sql).all(...binds, limit, offset);
        return { total, rows: rows.map(r => this.formatRow(r)) };
    }

    getSenderSuggestions(prefix, limit = 15) {
        if (!this.db) return [];
        try {
            const clean = (prefix || '').trim();
            const pattern = `%${clean}%`;
            const rows = this.db.prepare(`
                SELECT DISTINCT sender 
                FROM emails 
                WHERE sender LIKE ? AND sender != '' AND sender != 'Unknown'
                ORDER BY sender ASC 
                LIMIT ?;
            `).all(pattern, limit);
            return rows.map(r => r.sender);
        } catch (err) {
            console.error('[VaultDB] getSenderSuggestions error:', err.message);
            return [];
        }
    }

    getStats() {
        if (!this.db) return { total: 0, malicious: 0, suspicious: 0, spam: 0, safe: 0 };
        try {
            const rows = this.db.prepare(`
                SELECT 
                    COUNT(*) as total,
                    SUM(CASE WHEN LOWER(verdict) LIKE '%malicious%' THEN 1 ELSE 0 END) as malicious,
                    SUM(CASE WHEN LOWER(verdict) LIKE '%suspicious%' THEN 1 ELSE 0 END) as suspicious,
                    SUM(CASE WHEN LOWER(verdict) LIKE '%spam%' THEN 1 ELSE 0 END) as spam,
                    SUM(CASE WHEN LOWER(verdict) LIKE '%safe%' THEN 1 ELSE 0 END) as safe
                FROM emails;
            `).get();
            return {
                total: rows.total || 0,
                malicious: rows.malicious || 0,
                suspicious: rows.suspicious || 0,
                spam: rows.spam || 0,
                safe: rows.safe || 0
            };
        } catch (err) {
            console.error('[VaultDB] getStats error:', err.message);
            return { total: 0, malicious: 0, suspicious: 0, spam: 0, safe: 0 };
        }
    }

    getEmailById(id) {
        if (!this.db || !id) return null;
        try {
            const row = this.db.prepare(`SELECT * FROM emails WHERE id = ? OR entry_id = ? OR fingerprint = ? LIMIT 1`).get(id, id, id);
            return row ? this.formatRow(row) : null;
        } catch (err) {
            console.error('[VaultDB] getEmailById error:', err.message);
            return null;
        }
    }

    formatRow(row) {
        let atts = [];
        try {
            if (row.attachments_json) atts = JSON.parse(row.attachments_json);
        } catch {}

        return {
            id: row.id,
            entryId: row.entry_id,
            storeId: row.store_id,
            fingerprint: row.fingerprint,
            subject: row.subject,
            sender: row.sender,
            from: row.sender,
            recipient: row.recipient,
            to: row.recipient,
            cc: row.cc,
            date: row.date,
            time: row.time,
            timestamp: row.timestamp,
            body: row.body,
            headers: row.headers,
            fullHeaders: row.headers,
            verdict: row.verdict,
            score: row.score,
            tier: row.tier,
            attachments: atts,
            unread: !!row.unread
        };
    }

    close() {
        if (this.db) {
            try { this.db.close(); } catch {}
            this.db = null;
        }
    }
}

module.exports = new VaultDB();
