-- Amazon Ads Optimizer — reporting database
-- Rebuilt from scratch on every run of build_db.py (see that file's docstring).
-- Source of truth stays the client folders (config.json, logs/, context/); this
-- database is a derived, queryable view for dashboards. Never hand-edit it.

DROP TABLE IF EXISTS bid_decisions;
DROP TABLE IF EXISTS batches;
DROP TABLE IF EXISTS product_sales;
DROP TABLE IF EXISTS product_roi_targets;
DROP TABLE IF EXISTS clients;

CREATE TABLE clients (
    client_id   TEXT PRIMARY KEY,   -- folder name under clients/, e.g. 'jmn'
    target_acos REAL                -- percent, e.g. 15 = 15%
);

CREATE TABLE batches (
    batch_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id     TEXT NOT NULL REFERENCES clients(client_id),
    batch_date    TEXT NOT NULL,    -- ISO date, from logs/changes_<date>.csv filename
    account_aov   REAL,
    baseline_cvr  REAL,
    UNIQUE(client_id, batch_date)
);

CREATE TABLE bid_decisions (
    decision_id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id    INTEGER NOT NULL REFERENCES batches(batch_id),
    action      TEXT NOT NULL,      -- 'changed' | 'held'
    entity      TEXT,               -- 'Keyword' | 'Product targeting'
    target_id   TEXT,               -- Keyword ID / Product Targeting ID
    campaign    TEXT,
    ad_group    TEXT,
    target_text TEXT,
    old_bid     REAL,
    new_bid     REAL,
    clicks      REAL,
    spend       REAL,
    sales       REAL,
    orders      REAL,
    reason      TEXT,
    tier        TEXT,               -- 'normal' | 'extreme' | NULL, parsed from reason
    direction   TEXT                -- 'up' | 'down' | 'hold', derived from old/new bid
);

CREATE TABLE product_sales (
    client_id TEXT NOT NULL REFERENCES clients(client_id),
    sku       TEXT NOT NULL,
    market    TEXT NOT NULL,        -- 'CAD' | 'USD'
    year      INTEGER NOT NULL,
    month     INTEGER NOT NULL,     -- 1-12
    sales     REAL,
    PRIMARY KEY (client_id, sku, market, year, month)
);

CREATE TABLE product_roi_targets (
    client_id  TEXT NOT NULL REFERENCES clients(client_id),
    sku        TEXT NOT NULL,
    bucket     TEXT NOT NULL,       -- 'Q4' | 'Q1-2-3'
    roi_target REAL,                -- sales / spend
    PRIMARY KEY (client_id, sku, bucket)
);
