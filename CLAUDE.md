# Amazon Ads Optimization Engine

Multi-client Amazon Ads (Sponsored Products) bid optimization engine for Leonard Agence Web
(user: Joseph, joseph.papillon@leonardagenceweb.com). Built July 2026 with Claude Code.
This file is the complete project context — it lets any Claude session (in any organization)
continue the work without the original conversation.

## What this project does

Clients export Amazon Ads **Bulk Operations files** (multi-tab .xlsx from Seller Central) into a
per-client `input/` folder. The engine (`optimize.py`) applies rule-based bid optimizations and
writes a **ready-to-re-upload bulk file** to `output/` plus a plain-English decision log to `logs/`.
The human reviews the log, then manually uploads the output file in Seller Central.

**Deliberate design choices (validated with Joseph — do not silently change):**
- **Manual-trigger, human-uploads workflow.** No Amazon Ads API automation (access not confirmed).
  The engine never pushes changes itself.
- **Optimizations run on 14-day exports only.** Joseph explicitly rejected using 60-day data as
  the decision window. Longer exports are context/reference only.
- **Small bid swings by default** (5–10%), larger moves (up to 20–25%) reserved for extreme cases.
- **Statistical confidence before cutting zero-order keywords** — not flat spend thresholds.
- Scope v1: bid adjustments (done) + keyword harvesting + negative harvesting (pending Search
  Term Report — see Open items).

## Folder structure

```
amazon-ads-optimizer/
  CLAUDE.md                 <- this file
  optimize.py               <- the engine (Python 3, requires openpyxl)
  config.template.json      <- defaults used to seed new clients
  db/
    schema.sql              <- SQLite schema for the reporting database
    build_db.py             <- rebuilds db/optimizer.db + db/export.json from clients/ (read-only
                                on clients/; run `python3 db/build_db.py` after any new batch)
    optimizer.db, export.json  <- generated, gitignored — never hand-edit, just rerun build_db.py
  clients/
    jmn/                    <- first live client (only active one so far)
      config.json           <- target_acos: 15 (percent) — verified by Joseph
      input/                <- drop 14-day bulk export here (engine picks newest .xlsx)
      output/               <- bulk_upload_ready_<date>.xlsx (upload this to Seller Central)
      logs/                 <- changes_<date>.csv — decision log AND the engine's memory
      context/              <- reference files, NOT read by optimize.py (see below)
    client-2 .. client-6/   <- placeholders: same structure, target_acos still null, no data yet
```

Run with: `python3 optimize.py <client-folder-name>` (e.g. `python3 optimize.py jmn`)
from the project root. Only .xlsx input supported; newest file in `input/` wins.

## How the engine decides (optimize.py)

Operates on rows Entity ∈ {Keyword, Product targeting}, State=enabled, in the
"Sponsored Products Campaigns" tab. Sets `Operation=Update` + new `Bid` only on changed rows;
everything else left untouched so the upload only affects intended rows.

Computed from the file itself each run:
- **account AOV** = total Sales / total Orders (across eligible rows with orders)
- **baseline CVR** = total Orders / total Clicks — powers the zero-order confidence math
- **target CPA** = target_acos × AOV

Decision ladder per keyword/target (all thresholds in config.json):
1. **Data gate:** < 10 clicks → no change.
2. **Has orders:**
   - ACOS > target → cut toward `bid × target/actual`, capped -10% (normal) or -25% (extreme:
     ACOS ≥ 2× target).
   - ACOS ≤ 80% of target → raise +5–10% (normal, scaled) or +20% (extreme: ACOS ≤ 50% of target).
   - In between (mid-band) → no change (prevents churn).
3. **Zero orders:** compute p_zero = (1 − baseline_CVR)^clicks = probability an average keyword
   would show 0 orders by chance. Only cut if p_zero ≤ 15% (cap -10%), extreme if p_zero ≤ 5%
   AND spend ≥ 2× target CPA (cap -25%). Bid floor $0.02. Rationale: e.g. 23 clicks/0 orders at
   4% CVR is ~40% likely to be pure variance — not evidence.

### History / memory system (important)

The engine reads all past `logs/changes_*.csv` (excluding same-day, so re-runs replace today's
batch) and applies `history_rules`:
- **Cooldown 7d:** keyword changed in the last batch isn't touched again (14d report window
  still overlaps the change → would double-count evidence). Extreme tier bypasses.
- **Reversal protection 14d:** direction flip vs last batch requires 14d fresh data or extreme
  signal. This is Joseph's "don't cancel last week's work" requirement.
- **Manual-override detection:** if exported bid ≠ what engine set last time, log notes it.
- Held decisions are logged with `action=held` + reason; `action=changed` rows are the memory.

**Caveat:** history assumes every output batch was actually uploaded. If a batch is skipped,
its log should be deleted (or a not-applied marker added — feature not built yet).

## Client: jmn (the only active client)

Wood cutting boards / butcher blocks brand ("WFC" SKU prefix, competitor benchmark vs John Boos).
Markets: Canada (CAD) + USA (USD). Also runs Facebook ads (outside this engine's scope).
- **target_acos: 15** (Joseph first said 20%, then verified: 15% average).
- Account stats from 14d file: AOV ≈ $118, baseline CVR ≈ 3.95%, ~300 eligible targets,
  ~23K rows (mostly negative keywords).
- Campaign naming: "NS - SP - Phrase - …", "SP - Auto - Low Bid - CA", etc. Ad groups mix
  multiple SKUs (e.g. one ad group advertises 6 maple-line SKUs) — this blocks clean per-SKU
  target mapping (see Open items).
- **Batch of 2026-07-16** (23 changes: 16 cuts incl. maple-line keywords at 88–324% ACOS,
  7 raises on proven winners) sits in output/ — upload status unknown; ask Joseph before
  treating it as applied.

### context/ folder contents (reference only, never auto-read by the engine)

- `jmn_14d.xlsx` in input/ = live decision file (14-day bulk export, taken ~2026-07-16).
- `context/jmn_60d.xlsx` = 60-day bulk export. Analysis showed 45/300 targets get different
  calls on 60d vs 14d data, some direction reversals — evidence 14d windows are noisy, but
  Joseph still mandates 14d for decisions.
- `context/JMN - Données rapports mensuels - 2022-2026.xlsx` = monthly performance workbook
  (French), 2022→2026: monthly CA/US sales, Amazon+Facebook spend, TACOS, per-product sales,
  market share tabs, competitor tab (WFC vs Boos).
  - **`ROI target` tab:** per-SKU minimum ROI targets in two buckets — col B "Q4", col C
    "Q1-2-3" (ROI = sales/spend = 1/ACOS). E.g. CLA-201-MA: ROI 6 (Q4) / 7 (rest) → ACOS
    16.7%/14.3%. Eight CLA-line SKUs listed; MO250 = "n/a"; BPR/ALT/MZU/MO473 lines absent.
    These targets are NOT wired into the engine (Joseph dismissed the mapping questions —
    unresolved, see Open items).
  - A note in that tab from patricia.rochette@leonardagenceweb.com about Q4 aggressiveness is
    STALE — Joseph said to ignore it.

## Open items / next steps

1. **Harvesting not yet implemented** (keyword + negative): needs a **Search Term Report**
   export dropped into input/ alongside the bulk file. Planned rules (in config): promote
   search terms with ≥3 orders to exact match + negate in source; negate terms with ≥15 clicks
   / 0 orders. decide_bid untouched by this.
2. **Per-SKU ROI targets unmapped:** open questions Joseph hasn't answered — how to map
   per-SKU targets onto ad groups that mix SKUs (proposed: sales-weighted blend), what target
   for unlisted product lines, and whether "Q4" bucket = calendar Q4. Flat 15% governs until
   he decides.
3. **Other 5 clients:** rename folders, set target_acos, get exports.
4. **Batch-not-applied marker** (see history caveat).
5. Sponsored Brands tab present in exports but not optimized (SP only so far).

## Collaboration (git)

This project is shared with a colleague, so it's a git repo (initialized 2026-07-17) instead of a
synced folder (Drive/Dropbox). Reason: `logs/*.csv` is the engine's memory (cooldown/reversal
rules read it to decide what's safe to touch) — a plain file-sync tool has no locking, so two
people running `optimize.py` around the same time can silently diverge or overwrite each other's
history. Git surfaces that as a merge conflict instead of silent data loss.

- **Tracked:** everything under `clients/` (config.json, input/*.xlsx, output/*.xlsx,
  logs/*.csv, context/*.xlsx), `optimize.py`, `db/build_db.py`, `db/schema.sql`, this file.
- **Gitignored:** `db/optimizer.db`, `db/export.json` (derived — rerun `build_db.py` instead of
  committing), `.DS_Store`, `__pycache__/`.
- **Workflow:** only one person runs `optimize.py` for a given client at a time. `git pull`
  before running it, commit + `git push` right after (the new `logs/changes_<date>.csv` and
  `output/bulk_upload_ready_<date>.xlsx` it produces) so the other person's next run sees it.
- No remote is configured yet — this is a local repo only until Joseph decides where to host it
  (e.g. a private GitHub repo) for the colleague to clone.

## Conventions

- All money in file currency (mixed CAD/USD across campaigns — not currency-adjusted).
- `target_acos` in config is a percent (15 = 15%).
- Logs are append-only history — never delete/rewrite past `changes_*.csv` for an uploaded batch.
- Communication: Joseph reviews every batch before upload; keep logs human-readable.
