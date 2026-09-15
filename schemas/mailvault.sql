-- MailVault local store. SQLite 3, WAL.
-- Apply with: sqlite3 mailvault.db < schemas/mailvault.sql

PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS stores (
    store_id        TEXT PRIMARY KEY,
    display_name    TEXT NOT NULL,
    store_kind      TEXT NOT NULL CHECK (store_kind IN ('ost', 'pst', 'other')),
    profile_name    TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS messages (
    id                    INTEGER PRIMARY KEY,
    store_id              TEXT NOT NULL REFERENCES stores(store_id),
    entry_id              TEXT NOT NULL,
    internet_message_id   TEXT,
    conversation_id       TEXT,
    conversation_index    TEXT,
    folder                TEXT NOT NULL,
    is_sent               INTEGER NOT NULL DEFAULT 0 CHECK (is_sent IN (0, 1)),
    sent_at               TEXT,
    received_at           TEXT,
    from_addr             TEXT,
    to_addrs              TEXT,
    cc_addrs              TEXT,
    subject               TEXT,
    body_text             TEXT,
    body_hash             TEXT NOT NULL,
    has_attachments       INTEGER NOT NULL DEFAULT 0 CHECK (has_attachments IN (0, 1)),
    attachment_names      TEXT,
    status                TEXT NOT NULL DEFAULT 'ingested'
                          CHECK (status IN (
                              'ingested', 'embedded', 'classified',
                              'drafted', 'skipped', 'quarantined'
                          )),
    quarantine_reason     TEXT,
    ingested_at           TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (store_id, entry_id)
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_messages_imid
    ON messages(internet_message_id)
    WHERE internet_message_id IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_messages_conversation
    ON messages(conversation_id, sent_at);

CREATE INDEX IF NOT EXISTS idx_messages_status
    ON messages(status, received_at);

CREATE TABLE IF NOT EXISTS threads (
    conversation_id   TEXT PRIMARY KEY,
    store_id          TEXT NOT NULL REFERENCES stores(store_id),
    message_count     INTEGER NOT NULL DEFAULT 0,
    last_activity     TEXT,
    needs_reply       INTEGER NOT NULL DEFAULT 0 CHECK (needs_reply IN (0, 1)),
    subject_canonical TEXT
);

CREATE TABLE IF NOT EXISTS classifications (
    id              INTEGER PRIMARY KEY,
    message_id      INTEGER NOT NULL REFERENCES messages(id),
    priority        INTEGER NOT NULL CHECK (priority BETWEEN 0 AND 100),
    needs_reply     INTEGER NOT NULL CHECK (needs_reply IN (0, 1)),
    intent          TEXT NOT NULL,
    language        TEXT,
    suggested_folder TEXT,
    deadline_hint   TEXT,
    entities_json   TEXT,
    confidence      REAL NOT NULL,
    rule_hit        TEXT,
    model_id        TEXT,
    prompt_hash     TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_classifications_message
    ON classifications(message_id, created_at);

CREATE TABLE IF NOT EXISTS contacts (
    id              INTEGER PRIMARY KEY,
    email           TEXT NOT NULL UNIQUE,
    display_name    TEXT,
    language        TEXT,
    form_of_address TEXT,
    notes           TEXT,
    confirmed_by_user INTEGER NOT NULL DEFAULT 0 CHECK (confirmed_by_user IN (0, 1)),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS embeddings (
    id              INTEGER PRIMARY KEY,
    message_id      INTEGER NOT NULL REFERENCES messages(id),
    chunk_index     INTEGER NOT NULL,
    chunk_text      TEXT NOT NULL,
    embedding       BLOB NOT NULL,
    model_id        TEXT NOT NULL,
    UNIQUE (message_id, chunk_index)
);

CREATE TABLE IF NOT EXISTS drafts (
    id                  INTEGER PRIMARY KEY,
    message_id          INTEGER NOT NULL REFERENCES messages(id),
    variant             INTEGER NOT NULL CHECK (variant BETWEEN 1 AND 3),
    subject             TEXT,
    body_text           TEXT NOT NULL,
    citations_json      TEXT NOT NULL,
    risks_json          TEXT,
    model_id            TEXT NOT NULL,
    prompt_hash         TEXT,
    accepted            INTEGER NOT NULL DEFAULT 0 CHECK (accepted IN (0, 1)),
    edited_body         TEXT,
    copied_to_outlook_at TEXT,
    created_at          TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (message_id, variant)
);

CREATE TABLE IF NOT EXISTS jobs (
    id              INTEGER PRIMARY KEY,
    kind            TEXT NOT NULL CHECK (kind IN (
                        'ingest', 'embed', 'classify', 'draft', 'backfill'
                    )),
    message_id      INTEGER REFERENCES messages(id),
    priority        INTEGER NOT NULL DEFAULT 50,
    status          TEXT NOT NULL DEFAULT 'queued'
                    CHECK (status IN ('queued', 'leased', 'done', 'failed')),
    attempts        INTEGER NOT NULL DEFAULT 0,
    lease_until     TEXT,
    last_error      TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_jobs_queue
    ON jobs(status, priority, created_at);

CREATE TABLE IF NOT EXISTS audit_events (
    id              INTEGER PRIMARY KEY,
    at              TEXT NOT NULL DEFAULT (datetime('now')),
    actor           TEXT NOT NULL,
    action          TEXT NOT NULL,
    message_id      INTEGER,
    draft_id        INTEGER,
    body_hash       TEXT,
    detail          TEXT
);

CREATE TABLE IF NOT EXISTS rules (
    id              INTEGER PRIMARY KEY,
    name            TEXT NOT NULL,
    enabled         INTEGER NOT NULL DEFAULT 1 CHECK (enabled IN (0, 1)),
    sender_pattern  TEXT,
    domain_pattern  TEXT,
    subject_pattern TEXT,
    header_pattern  TEXT,
    action          TEXT NOT NULL CHECK (action IN (
                        'vip', 'newsletter', 'ndr', 'force_reply', 'ignore'
                    )),
    priority_set    INTEGER,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE VIRTUAL TABLE IF NOT EXISTS messages_fts USING fts5(
    subject,
    body_text,
    from_addr,
    content='messages',
    content_rowid='id'
);

CREATE TRIGGER IF NOT EXISTS messages_ai AFTER INSERT ON messages BEGIN
    INSERT INTO messages_fts(rowid, subject, body_text, from_addr)
    VALUES (new.id, new.subject, new.body_text, new.from_addr);
END;

CREATE TRIGGER IF NOT EXISTS messages_ad AFTER DELETE ON messages BEGIN
    INSERT INTO messages_fts(messages_fts, rowid, subject, body_text, from_addr)
    VALUES ('delete', old.id, old.subject, old.body_text, old.from_addr);
END;

CREATE TRIGGER IF NOT EXISTS messages_au AFTER UPDATE ON messages BEGIN
    INSERT INTO messages_fts(messages_fts, rowid, subject, body_text, from_addr)
    VALUES ('delete', old.id, old.subject, old.body_text, old.from_addr);
    INSERT INTO messages_fts(rowid, subject, body_text, from_addr)
    VALUES (new.id, new.subject, new.body_text, new.from_addr);
END;
