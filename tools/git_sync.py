#!/usr/bin/env python3
"""Local one-page app with two buttons — Récupérer (pull) and Envoyer (add+commit+push) —
for the amazon-ads-optimizer repo. No terminal typing needed day to day: double-click
"Synchroniser avec Git.command" at the project root, a browser tab opens on localhost.

Runs on 127.0.0.1 only (not exposed on the network). Single-user tool, no auth needed.
"""
import http.server
import json
import socket
import subprocess
import webbrowser
from datetime import datetime

PORT = 8765
PROJECT_DIR = "/Users/josephpapillon/claudecode/amazon-ads-optimizer"

PAGE = """<!doctype html>
<meta charset="utf-8">
<title>Sync Git — amazon-ads-optimizer</title>
<style>
  body{font-family:-apple-system,sans-serif;max-width:640px;margin:60px auto;padding:0 20px;color:#1B2226;background:#F4F4F2;}
  h1{font-size:20px;}
  .row{display:flex;gap:14px;margin:28px 0;}
  button{flex:1;font:inherit;font-size:15px;font-weight:600;padding:16px;border-radius:10px;border:1px solid #d8d8d4;cursor:pointer;background:white;}
  button:hover{background:#eee;}
  button:disabled{opacity:.5;cursor:default;}
  #pull{color:#2F6FA8;} #push{color:#2F7A55;}
  pre{background:#111;color:#ddd;padding:14px;border-radius:8px;font-size:12.5px;white-space:pre-wrap;min-height:60px;max-height:360px;overflow:auto;}
  .status{font-size:13px;color:#666;margin-top:6px;}
</style>
<h1>amazon-ads-optimizer — synchronisation Git</h1>
<p class="status">Dossier : <code>{project_dir}</code></p>
<div class="row">
  <button id="pull">⬇️ Récupérer (pull)</button>
  <button id="push">⬆️ Envoyer (push)</button>
</div>
<pre id="out">Prêt.</pre>
<script>
async function run(action, btn){{
  const out = document.getElementById('out');
  document.querySelectorAll('button').forEach(b=>b.disabled=true);
  out.textContent = 'En cours...';
  try {{
    const res = await fetch('/'+action, {{method:'POST'}});
    const data = await res.json();
    out.textContent = data.output || '(rien)';
  }} catch(e) {{
    out.textContent = 'Erreur : ' + e;
  }}
  document.querySelectorAll('button').forEach(b=>b.disabled=false);
}}
document.getElementById('pull').addEventListener('click', ()=>run('pull'));
document.getElementById('push').addEventListener('click', ()=>run('push'));
</script>
""".replace("{project_dir}", PROJECT_DIR)


def run_git(*args):
    result = subprocess.run(
        ["git", *args], cwd=PROJECT_DIR, capture_output=True, text=True, timeout=60
    )
    text = (result.stdout + result.stderr).strip()
    return result.returncode == 0, text


def do_pull():
    ok, out = run_git("pull")
    return ok, out or "Déjà à jour."


def do_push():
    run_git("add", "-A")
    ok, status_out = run_git("status", "--porcelain")
    if ok and not status_out:
        return True, "Rien à envoyer — tout est déjà à jour."
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    ok, commit_out = run_git("commit", "-m", f"Sync depuis {socket.gethostname()} — {stamp}")
    if not ok:
        return False, commit_out
    ok, push_out = run_git("push")
    if not ok:
        return False, commit_out + "\n" + push_out + "\n\n(Astuce : essaie 'Récupérer' d'abord si le message parle de commits distants manquants.)"
    return True, commit_out + "\n" + push_out


class Handler(http.server.BaseHTTPRequestHandler):
    def _send(self, code, payload, content_type="application/json"):
        body = payload if isinstance(payload, bytes) else payload.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/":
            self._send(200, PAGE, "text/html; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        if self.path == "/pull":
            ok, output = do_pull()
        elif self.path == "/push":
            ok, output = do_push()
        else:
            self._send(404, json.dumps({"ok": False, "output": "route inconnue"}))
            return
        self._send(200, json.dumps({"ok": ok, "output": output}))

    def log_message(self, fmt, *args):
        pass


if __name__ == "__main__":
    server = http.server.HTTPServer(("127.0.0.1", PORT), Handler)
    webbrowser.open(f"http://127.0.0.1:{PORT}/")
    print(f"Sync Git en cours sur http://127.0.0.1:{PORT}/ — Ctrl+C pour arrêter.")
    server.serve_forever()
