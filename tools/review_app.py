#!/usr/bin/env python3
"""Local checkbox UI to review a pending optimize.py batch before it's applied.

Double-click "Réviser un lot.command" at the project root: finds the most recent
review_<date>.csv/.json pair under clients/*/review/ that hasn't been applied yet
(no matching output/bulk_upload_ready_<date>.xlsx), shows every proposed change with
a checkbox (checked = approved), and an Appliquer button that rewrites the review
CSV's approve column and runs `python3 optimize.py <client> --apply <date>` for you.

Runs on 127.0.0.1 only. Single-user tool, no auth needed.
"""
import csv
import glob
import html
import http.server
import json
import os
import subprocess
import webbrowser

PORT = 8767
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLIENTS_DIR = os.path.join(BASE, "clients")


def find_pending_review():
    candidates = []
    for review_csv in glob.glob(os.path.join(CLIENTS_DIR, "*", "review", "review_*.csv")):
        parts = review_csv.split(os.sep)
        client = parts[-3]
        date_str = os.path.basename(review_csv)[len("review_"):-len(".csv")]
        applied = os.path.exists(os.path.join(CLIENTS_DIR, client, "output", f"bulk_upload_ready_{date_str}.xlsx"))
        candidates.append({"client": client, "date": date_str, "path": review_csv, "applied": applied})
    candidates.sort(key=lambda c: c["date"], reverse=True)
    pending = [c for c in candidates if not c["applied"]]
    return pending[0] if pending else None


def load_rows(review_csv):
    with open(review_csv, newline="") as f:
        return list(csv.DictReader(f))


def fmt_bid(v):
    return f"{float(v):.2f} $"


def row_html(i, row):
    direction = "up" if float(row["new_bid"]) > float(row["old_bid"]) else "down"
    tone = "good" if direction == "up" else "cut"
    checked = "checked" if row["approve"].strip().upper() in ("TRUE", "1", "YES") else ""
    return f"""
    <label class="row">
      <input type="checkbox" data-id="{html.escape(row['id'])}" {checked}>
      <div class="row-main">
        <div class="row-top">
          <span class="target">{html.escape(row['target'] or '')}</span>
          <span class="bid {tone}">{fmt_bid(row['old_bid'])} → {fmt_bid(row['new_bid'])}</span>
        </div>
        <div class="row-sub">{html.escape(row['campaign'] or '')} · {html.escape(row['ad_group'] or '')} · {row['clicks']} clics</div>
        <div class="row-reason">{html.escape(row['reason'] or '')}</div>
      </div>
    </label>"""


def render_page():
    pending = find_pending_review()
    if not pending:
        return """<!doctype html><meta charset="utf-8"><title>Révision de lot</title>
        <body style="font-family:-apple-system,sans-serif;max-width:640px;margin:80px auto;text-align:center;color:#555">
        <h1 style="color:#222">Aucune revue en attente</h1>
        <p>Lance <code>python3 optimize.py &lt;client&gt;</code> pour proposer un nouveau lot.</p>
        </body>"""

    rows = load_rows(pending["path"])
    rows_html = "\n".join(row_html(i, r) for i, r in enumerate(rows))
    n_approved = sum(1 for r in rows if r["approve"].strip().upper() in ("TRUE", "1", "YES"))

    return f"""<!doctype html>
<meta charset="utf-8">
<title>Révision de lot — {pending['client']}</title>
<style>
  body{{font-family:-apple-system,sans-serif;max-width:760px;margin:40px auto;padding:0 20px;color:#1B2226;background:#F4F4F2;}}
  h1{{font-size:19px;margin-bottom:2px;}}
  .sub{{color:#666;font-size:13px;margin-bottom:20px;}}
  .toolbar{{display:flex;gap:10px;margin-bottom:14px;align-items:center;}}
  .toolbar button.link{{background:none;border:none;color:#2F6FA8;font:inherit;font-size:13px;cursor:pointer;padding:0;text-decoration:underline;}}
  .count{{font-size:13px;color:#666;margin-left:auto;}}
  .list{{background:white;border:1px solid #ddd;border-radius:10px;overflow:hidden;}}
  .row{{display:flex;gap:12px;padding:12px 16px;border-bottom:1px solid #eee;cursor:pointer;align-items:flex-start;}}
  .row:last-child{{border-bottom:none;}}
  .row:hover{{background:#fafafa;}}
  .row input{{margin-top:4px;flex-shrink:0;}}
  .row-main{{flex:1;min-width:0;}}
  .row-top{{display:flex;justify-content:space-between;gap:12px;}}
  .target{{font-weight:600;font-size:14px;}}
  .bid{{font-family:ui-monospace,monospace;font-size:13px;white-space:nowrap;}}
  .bid.up{{color:#2F7A55;}} .bid.down{{color:#A8402F;}}
  .row-sub{{font-size:12px;color:#888;margin-top:2px;}}
  .row-reason{{font-size:12.5px;color:#555;margin-top:4px;}}
  .apply-bar{{position:sticky;bottom:0;background:#F4F4F2;padding:16px 0;margin-top:16px;}}
  #apply{{width:100%;font:inherit;font-size:15px;font-weight:600;padding:16px;border-radius:10px;border:1px solid #2F7A55;background:#2F7A55;color:white;cursor:pointer;}}
  #apply:disabled{{opacity:.5;cursor:default;}}
  pre{{background:#111;color:#ddd;padding:14px;border-radius:8px;font-size:12.5px;white-space:pre-wrap;min-height:20px;max-height:300px;overflow:auto;margin-top:12px;}}
</style>
<h1>Lot du {pending['date']} — {pending['client']}</h1>
<p class="sub">Décoche ce que tu ne veux pas appliquer, puis clique Appliquer. Tout ce qui est retenu par les règles d'historique n'apparaît pas ici — c'est déjà bloqué automatiquement.</p>
<div class="toolbar">
  <button class="link" id="all">Tout cocher</button>
  <button class="link" id="none">Tout décocher</button>
  <span class="count" id="count">{n_approved} / {len(rows)} approuvé(s)</span>
</div>
<div class="list">{rows_html}</div>
<div class="apply-bar">
  <button id="apply">Appliquer ({n_approved} approuvé(s))</button>
  <pre id="out" style="display:none"></pre>
</div>
<script>
const client = {json.dumps(pending['client'])};
const date = {json.dumps(pending['date'])};

function updateCount() {{
  const boxes = document.querySelectorAll('.row input');
  const checked = document.querySelectorAll('.row input:checked').length;
  document.getElementById('count').textContent = checked + ' / ' + boxes.length + ' approuvé(s)';
  document.getElementById('apply').textContent = 'Appliquer (' + checked + ' approuvé(s))';
}}
document.querySelectorAll('.row input').forEach(cb => cb.addEventListener('change', updateCount));
document.getElementById('all').addEventListener('click', () => {{
  document.querySelectorAll('.row input').forEach(cb => cb.checked = true);
  updateCount();
}});
document.getElementById('none').addEventListener('click', () => {{
  document.querySelectorAll('.row input').forEach(cb => cb.checked = false);
  updateCount();
}});
document.getElementById('apply').addEventListener('click', async () => {{
  const approvedIds = Array.from(document.querySelectorAll('.row input'))
    .filter(cb => cb.checked).map(cb => cb.dataset.id);
  const btn = document.getElementById('apply');
  const out = document.getElementById('out');
  btn.disabled = true;
  out.style.display = 'block';
  out.textContent = 'Application en cours...';
  try {{
    const res = await fetch('/apply', {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify({{client, date, approved_ids: approvedIds}})
    }});
    const data = await res.json();
    out.textContent = data.output;
    if (data.ok) {{ btn.textContent = 'Appliqué ✓'; }}
    else {{ btn.disabled = false; }}
  }} catch (e) {{
    out.textContent = 'Erreur : ' + e;
    btn.disabled = false;
  }}
}});
</script>
"""


def apply_review(client, date_str, approved_ids):
    review_csv = os.path.join(CLIENTS_DIR, client, "review", f"review_{date_str}.csv")
    rows = load_rows(review_csv)
    approved_set = set(approved_ids)
    for row in rows:
        row["approve"] = "TRUE" if row["id"] in approved_set else "FALSE"
    with open(review_csv, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)

    result = subprocess.run(
        ["python3", "optimize.py", client, "--apply", date_str],
        cwd=BASE, capture_output=True, text=True, timeout=120,
    )
    output = (result.stdout + result.stderr).strip()
    return result.returncode == 0, output


class Handler(http.server.BaseHTTPRequestHandler):
    def _send(self, code, body, content_type="text/html; charset=utf-8"):
        body = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/":
            self._send(200, render_page())
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        if self.path == "/apply":
            length = int(self.headers["Content-Length"])
            payload = json.loads(self.rfile.read(length))
            ok, output = apply_review(payload["client"], payload["date"], payload["approved_ids"])
            self._send(200, json.dumps({"ok": ok, "output": output}), "application/json")
        else:
            self._send(404, json.dumps({"ok": False, "output": "route inconnue"}), "application/json")

    def log_message(self, fmt, *args):
        pass


if __name__ == "__main__":
    server = http.server.HTTPServer(("127.0.0.1", PORT), Handler)
    webbrowser.open(f"http://127.0.0.1:{PORT}/")
    print(f"Révision de lot en cours sur http://127.0.0.1:{PORT}/ — Ctrl+C pour arrêter.")
    server.serve_forever()
