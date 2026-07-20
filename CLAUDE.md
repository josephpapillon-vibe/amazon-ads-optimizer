# Amazon Ads Optimization Engine

Multi-client Amazon Ads (Sponsored Products) bid optimization engine for Leonard Agence Web
(user: Joseph, joseph.papillon@leonardagenceweb.com). Built July 2026 with Claude Code.
This file is the complete project context — it lets any Claude session (in any organization)
continue the work without the original conversation, and lets a teammate get oriented without
having sat through the original build.

Repo: **github.com/josephpapillon-vibe/amazon-ads-optimizer** (private). See "Collaboration
(git)" below for how to get access and the day-to-day workflow.

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
  .gitignore                <- see "Collaboration (git)" for what's excluded and why
  Synchroniser avec Git.command  <- double-click launcher for tools/git_sync.py (macOS)
  optimize.py               <- the engine (Python 3, requires openpyxl)
  config.template.json      <- defaults used to seed new clients
  db/
    schema.sql              <- SQLite schema for the reporting database
    build_db.py             <- rebuilds db/optimizer.db + db/export.json from clients/ (read-only
                                on clients/; run `python3 db/build_db.py` after any new batch)
    optimizer.db, export.json  <- generated, gitignored — never hand-edit, just rerun build_db.py
  dashboard/
    template.html           <- the dashboard's HTML/CSS/JS, with a __DATA_JSON__ placeholder
    build_dashboard.py       <- reads db/export.json + clients/jmn/config.json, writes the file below
    jmn_dashboard.html       <- generated, gitignored — see "Dashboard" section below
  tools/
    git_sync.py             <- tiny localhost app: two buttons, Récupérer (pull) / Envoyer (push)
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

### ROAS baseline + escalation (added 2026-07-20)

The ladder above (steps 2-3) only decides **whether** a change is justified and its **tier**
(`normal` or `extreme`) — it no longer directly sets the applied bid on its own. Once a tier is
assigned, `roas_baseline_bid()` computes a second, more conservative number from a flat manual
rule of thumb (thresholds in `config.json`'s `roas_baseline` block, ROAS = Sales/Spend):
- spend < $10 → +2%
- ROAS > 4.5 → +5%
- ROAS < 3.5 & spend > $100 → -7%
- ROAS < 3.5 & spend $30-100 → -5%
- ROAS < 3.5 & spend $10-30 → -3%
- ROAS 3.5-4.5 (and spend ≥ $10) → neutral, no change

**Which bid actually gets applied:**
- **`normal` tier → the ROAS baseline governs.** This is deliberately more conservative than the
  ACOS ladder's own normal-tier move; the ACOS ladder's number is kept only as a comparison note
  in the log. If baseline is neutral, nothing is applied at all (row doesn't appear in the log,
  same as any other no-change case) — even if the ACOS ladder wanted a small move. If baseline
  disagrees on direction with the ACOS ladder (happens near the target ACOS boundary), the
  baseline's direction wins and the log flags it as "diverges" for manual review.
- **`extreme` tier → always escalates past the baseline to the full ACOS-ladder move.** This is
  the "highlight when the change should be more extreme" behavior Anthony asked for: `extreme`
  already means the ACOS ladder found strong statistical/ratio evidence (≥2× target ACOS, ≤50%
  of target, or a high-confidence zero-order cut) — that evidence is exactly when a human's flat
  ROAS rule under-reacts, so the engine overrides it and says so in the log.

This baseline mirrors a manual rule the team already used (not from Joseph — described by
Anthony on 2026-07-20; see [[jmn-roas-baseline-rule]]). `logs/changes_*.csv` gained two columns:
`baseline_bid` (what the ROAS rule alone would have set, blank when tier is `None`) and
`escalated` (`yes`/`no`, blank when tier is `None`) — use these instead of parsing the reason
text if analyzing a batch programmatically.

**Not yet done:** the 5 placeholder clients' `config.json` files are on an older, different
schema (no `history_rules`, no `roas_baseline` — see Open items #3) and were left alone.

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
its log should be deleted (or a not-applied marker added — feature not built yet). **This
already happened for real:** the 2026-07-16 batch's engine output was never uploaded (manual
bid changes were made instead, independently — see "Client: jmn" below), so
`logs/changes_2026-07-16.csv` fed false "applied history" into the 2026-07-20 run's
cooldown/reversal logic.

## Client: jmn (the only active client)

Wood cutting boards / butcher blocks brand ("WFC" SKU prefix, competitor benchmark vs John Boos).
Markets: Canada (CAD) + USA (USD). Also runs Facebook ads (outside this engine's scope).
- **target_acos: 15** (Joseph first said 20%, then verified: 15% average).
- Account stats from 14d file: AOV ≈ $118, baseline CVR ≈ 3.95%, ~300 eligible targets,
  ~23K rows (mostly negative keywords). Batch of 2026-07-20 re-run on a fresh 14d export
  gave AOV ≈ $115.34, baseline CVR ≈ 3.67% — normal week-to-week drift, not a data error.
- Campaign naming: "NS - SP - Phrase - …", "SP - Auto - Low Bid - CA", etc. Ad groups mix
  multiple SKUs (e.g. one ad group advertises 6 maple-line SKUs) — this blocks clean per-SKU
  target mapping (see Open items).
- **Bulk export language gotcha:** Seller Central can export the Bulk Operations file with
  French headers (e.g. "Entité", "État", "Enchère") depending on account/browser language —
  `optimize.py` only recognizes English headers and exits with a clear "missing columns" error
  if given a French export. Fix is to re-export in English (done successfully 2026-07-20); no
  code change was made to support French headers (Joseph's colleague chose re-export over
  patching optimize.py when asked).
- **Batch of 2026-07-16** (engine output: 23 changes — 16 cuts incl. maple-line keywords at
  88–324% ACOS, 7 raises on proven winners) — **the engine's output file was NOT what got
  uploaded.** Bids in the account were changed on/around 2026-07-16 via separate manual
  analysis, independent of this engine's recommendations. `logs/changes_2026-07-16.csv` records
  what the engine *proposed*, not what was actually applied to the account — treat it as
  unreliable memory for cooldown/reversal purposes (see caveat below and Open items #4).
- **Batch of 2026-07-20** — re-run twice the same day (same-day re-runs replace the batch, by
  design): first on the ACOS ladder alone (14 changes: 4 raises, 10 cuts; 10 held), then again
  after adding the ROAS baseline + escalation layer (see above). **Current output/log reflects
  the second run:** 14 changes (5 raises, 9 cuts, 11 of the 14 escalated to the full ACOS-ladder
  move), 8 held. 2 of the original 10 holds (butcher block, wood cutting board /556706285412304)
  no longer appear at all — their ROAS lands in the 3.5-4.5 neutral zone, so the baseline itself
  says no change, making the cooldown question moot for those two. Upload status unknown as of
  this writing; ask before treating it as applied.

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

## Dashboard (visual explainer + product performance)

`dashboard/build_dashboard.py` turns `db/export.json` (run `db/build_db.py` first) into one
self-contained HTML file, `dashboard/jmn_dashboard.html`. It's a **static snapshot** — re-run
both scripts after every new batch or monthly-report update, it does not live-refresh. Publish
the generated file as a Claude Artifact (or just open it locally) to view it.

Two tabs:
1. **"Comment ça décide"** — walks through optimize.py's decision ladder (data gate → ACOS
   high/mid/low → zero-order statistical test → history rules), with a live calculator that
   reproduces `decide_bid()` in JS so anyone can test hypothetical numbers, plus a real
   breakdown of the 2026-07-16 batch (23 decisions) and the memory/cooldown rules explained.
2. **"Performance produit"** — monthly sales per SKU 2022–2026 (CAD/USD), sparklines, MoM/YoY,
   sourced from the `context/` monthly workbook, not from the bid engine. Clearly caveats that
   no per-SKU *actual* ROI is shown (ad groups mix SKUs, no per-SKU spend breakdown exists yet —
   see Open items #2) and that the latest month is partial when applicable.

**Known gap:** `ACCOUNT_AOV` / `BASELINE_CVR_PCT` at the top of `build_dashboard.py` are
hand-entered constants (from this file's "Client: jmn" numbers below), because `optimize.py`
computes account AOV / baseline CVR at run time but doesn't persist them into
`logs/changes_*.csv` — same root cause as the history caveat above. Update those two constants
by hand after each `optimize.py` run, or wire real persistence into the engine.

A version was published once already at `https://claude.ai/code/artifact/5f9df2fd-616e-46d1-ba12-1736c0308e8f`
— Artifacts are **private by default**, so Joseph needs to use that page's own share menu for
a teammate to see it directly. Anyone with this repo can also just rebuild and publish their own
copy from `dashboard/jmn_dashboard.html`.

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
6. **Persist account_aov / baseline_cvr per batch** into the log (currently computed then
   discarded — see the Dashboard section's "known gap").

## Collaboration (git)

This project is shared with a colleague, so it's a git repo (initialized 2026-07-17, pushed to
GitHub the same day) instead of a synced folder (Drive/Dropbox). Reason: `logs/*.csv` is the
engine's memory (cooldown/reversal rules read it to decide what's safe to touch) — a plain
file-sync tool has no locking, so two people running `optimize.py` around the same time can
silently diverge or overwrite each other's history. Git surfaces that as a merge conflict
instead of silent data loss.

- **Remote:** `git@github.com:josephpapillon-vibe/amazon-ads-optimizer.git` — **private** repo.
  Auth is via SSH key (GitHub stopped accepting account passwords over HTTPS in 2021, and a
  Personal Access Token is easy to mistype/lose); HTTPS + token was tried first and abandoned
  for that reason.
- **Tracked:** everything under `clients/` (config.json, input/*.xlsx, output/*.xlsx,
  logs/*.csv, context/*.xlsx), `optimize.py`, `db/build_db.py`, `db/schema.sql`,
  `dashboard/template.html`, `dashboard/build_dashboard.py`, `tools/git_sync.py`, this file.
- **Gitignored** (all regenerable — rerun the relevant script instead of committing):
  `db/optimizer.db`, `db/export.json`, `dashboard/jmn_dashboard.html`, `.DS_Store`, `__pycache__/`.
- **Workflow:** only one person runs `optimize.py` (or anything that writes into `clients/`)
  for a given client at a time. Pull before running it, commit + push right after (the new
  `logs/changes_<date>.csv` and `output/bulk_upload_ready_<date>.xlsx` it produces) so the other
  person's next run sees it.

### tools/git_sync.py — no-terminal sync

A colleague uncomfortable in a terminal can double-click **`Synchroniser avec Git.command`** at
the project root: it starts a tiny localhost web page (`tools/git_sync.py`, port 8765, bound to
127.0.0.1 only) with two buttons — **Récupérer** (`git pull`) and **Envoyer** (`git add -A` +
commit with an auto-generated message + `git push`). Output/errors show directly on the page.

**Caution:** "Envoyer" runs `git add -A` — it stages *everything* changed in the folder, no
per-file review. Glance at what changed before clicking it; an accidental rename or delete would
get pushed too. (Confirmed in testing: a client folder rename was picked up and pushed as-is,
then reverted in a follow-up commit — the tool did exactly what it was asked, for better or worse.)

### Onboarding a new collaborator

1. Joseph adds them on GitHub: repo → **Settings** → **Collaborators** → **Add people**.
2. They generate their own SSH key (`ssh-keygen -t ed25519 -C "their@email"`) and add the
   public key at **github.com/settings/ssh/new** — never share a private key or a token.
3. `git clone git@github.com:josephpapillon-vibe/amazon-ads-optimizer.git`
4. `python3 db/build_db.py` (needs `openpyxl`: `pip3 install openpyxl` if missing) to get a local
   `db/optimizer.db` — it's gitignored, everyone builds their own from the tracked source files.

## Conventions

- All money in file currency (mixed CAD/USD across campaigns — not currency-adjusted).
- `target_acos` in config is a percent (15 = 15%).
- Logs are append-only history — never delete/rewrite past `changes_*.csv` for an uploaded batch.
- Communication: Joseph reviews every batch before upload; keep logs human-readable.
