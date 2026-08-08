"""Netatmo OAuth2 – Verbindungstest mit interaktiver Credentials-Eingabe"""

import http.server
import os
import urllib.parse
import secrets
import webbrowser
import requests
import json
import getpass

print("=" * 60)
print("NETATMO AUTH TEST")
print("=" * 60)
print()
print("Zugangsdaten eingeben:")
print("(Enter = Wert aus NETATMO_CLIENT_ID / NETATMO_CLIENT_SECRET uebernehmen)")
print()

DEF_ID = os.environ.get("NETATMO_CLIENT_ID", "")
DEF_SEC = os.environ.get("NETATMO_CLIENT_SECRET", "")

id_hint = DEF_ID if DEF_ID else "keine Vorgabe"
entered_id = input(f"Client ID [{id_hint}]: ").strip()
client_id = entered_id if entered_id else DEF_ID
if not client_id:
    raise SystemExit("Keine Client ID angegeben.")

print("Client Secret (wird verborgen eingegeben):")
entered_sec = getpass.getpass(f"Client Secret: ")
client_secret = entered_sec if entered_sec else DEF_SEC
if not client_secret:
    raise SystemExit("Kein Client Secret angegeben.")

REDIRECT_URI = "http://localhost:9876/callback"
SCOPE = "read_station"
AUTH_URL = "https://api.netatmo.com/oauth2/authorize"
TOKEN_URL = "https://api.netatmo.com/oauth2/token"

print(f"\nVerwende:")
print(f"  client_id     = {client_id}")
print(f"  client_secret = {client_secret[:6]}...{client_secret[-4:]}")

# ── Schritt 1: Auth URL ───────────────────────────────────────────────────────

state = secrets.token_urlsafe(16)
auth_url = (
    AUTH_URL
    + "?response_type=code"
    + "&client_id=" + urllib.parse.quote(client_id, safe="")
    + "&redirect_uri=" + urllib.parse.quote(REDIRECT_URI, safe="")
    + "&scope=" + urllib.parse.quote(SCOPE, safe="")
    + "&state=" + state
)

print(f"\nAuth URL:\n{auth_url}\n")

# ── Schritt 2: Callback abfangen ──────────────────────────────────────────────

result = {"code": None, "state": None, "error": None}


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        p = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        result["code"] = p.get("code",  [None])[0]
        result["state"] = p.get("state", [None])[0]
        result["error"] = p.get("error", [None])[0]
        self.send_response(200)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"<h2>Callback empfangen - Fenster schliessen</h2>")

    def log_message(self, *_): pass


srv = http.server.HTTPServer(("localhost", 9876), Handler)
srv.timeout = 120

print("Oeffne Browser...")
webbrowser.open(auth_url)
print("Warte auf Callback (max 120s)...")
srv.handle_request()
srv.server_close()

print(f"code  = {result['code']}")
print(f"state match = {result['state'] == state}")
print(f"error = {result['error']}")

if result["error"] or not result["code"]:
    print("ABBRUCH: kein Code")
    raise SystemExit(1)

# ── Schritt 3: Token Exchange – alle Varianten ───────────────────────────────

CT = {"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"}
code = result["code"]

print("\n--- Token Exchange ---")
for label, data, extra in [
    ("Standard (mit redirect_uri + scope)", {
        "grant_type": "authorization_code",
        "client_id": client_id, "client_secret": client_secret,
        "code": code, "redirect_uri": REDIRECT_URI, "scope": SCOPE}, {}),
    ("Ohne redirect_uri", {
        "grant_type": "authorization_code",
        "client_id": client_id, "client_secret": client_secret,
        "code": code, "scope": SCOPE}, {}),
    ("Ohne scope", {
        "grant_type": "authorization_code",
        "client_id": client_id, "client_secret": client_secret,
        "code": code, "redirect_uri": REDIRECT_URI}, {}),
    ("Nur code + credentials", {
        "grant_type": "authorization_code",
        "client_id": client_id, "client_secret": client_secret,
        "code": code}, {}),
    ("Basic Auth", {
        "grant_type": "authorization_code",
        "code": code, "scope": SCOPE,
        "redirect_uri": REDIRECT_URI},
     {"auth": (client_id, client_secret)}),
]:
    resp = requests.post(TOKEN_URL, headers=CT, data=data, timeout=15, **extra)
    print(f"\n[{label}]")
    print(f"  HTTP {resp.status_code}: ", end="")
    try:
        b = resp.json()
        print(json.dumps(b))
    except Exception:
        print(resp.text[:300])
    if resp.status_code == 200:
        tok = resp.json()["access_token"]
        print(f"\n>>> ERFOLG! Token = {tok[:50]}...")
        break

print("\nTest abgeschlossen.")
