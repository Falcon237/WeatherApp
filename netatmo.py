#!/usr/bin/env python3
"""
Netatmo OAuth2-Client für den Netatmo Daten-Viewer.

Netatmo hat den direkten Access-Token (Password Grant) abgeschaltet.
Dieser Modul implementiert den Authorization Code Flow mit lokalem
Redirect-Server (einmalig) und automatischer Token-Erneuerung via
Refresh-Token (danach dauerhaft ohne Benutzerinteraktion).

Exports:
  NetatmoAuth             – Token-Verwaltung (OAuth2 + Refresh)
  NetatmoDataDownloader   – Messdaten-Abruf via API
  LoginWindow             – Tkinter-Dialog (öffnet Browser für OAuth)
"""

import datetime
import json
import os
import secrets
import threading
import time
import tkinter as tk
from http.server import BaseHTTPRequestHandler, HTTPServer
from tkinter import messagebox, ttk
from urllib import parse as urllib_parse
from urllib import request as urllib_request
from urllib import error as urllib_error

# ── Endpunkte ──────────────────────────────────────────────────────────────────
_AUTH_URL = 'https://api.netatmo.com/oauth2/authorize'
_TOKEN_URL = 'https://api.netatmo.com/oauth2/token'
_STATIONS_URL = 'https://api.netatmo.com/api/getstationsdata'
_MEASURE_URL = 'https://api.netatmo.com/api/getmeasure'

_API_MAX_ATTEMPTS = 6
_API_MIN_INTERVAL = 0.25
_API_RETRY_STATUS = {429, 500, 502, 503, 504}

_REDIRECT_PORT = 9731          # lokaler Port für OAuth-Callback
_REDIRECT_URI = f'http://localhost:{_REDIRECT_PORT}/callback'
_SCOPE = 'read_station'

_APPDATA_DIR = os.path.join(
    os.environ.get('APPDATA') or os.path.expanduser('~'), 'NetatmoViewer')
_TOKEN_FILE = os.path.join(_APPDATA_DIR, 'tokens.json')
_CREDS_FILE = os.path.join(_APPDATA_DIR, 'creds.json')

_UNIT_MAP: dict[str, str] = {
    'Temperature': '°C', 'Humidity': '%', 'CO2': 'ppm',
    'Pressure': 'hPa', 'Noise': 'dB', 'Rain': 'mm',
    'WindStrength': 'km/h', 'WindAngle': '°', 'GustStrength': 'km/h',
    'GustAngle': '°', 'sum_rain_1': 'mm', 'sum_rain_24': 'mm',
}


# ── Token-Persistenz ───────────────────────────────────────────────────────────

def _save_tokens(data: dict) -> None:
    os.makedirs(os.path.dirname(_TOKEN_FILE), exist_ok=True)
    with open(_TOKEN_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f)


def _load_tokens() -> dict | None:
    if not os.path.isfile(_TOKEN_FILE):
        return None
    try:
        with open(_TOKEN_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def _save_creds(client_id: str, client_secret: str) -> None:
    os.makedirs(_APPDATA_DIR, exist_ok=True)
    with open(_CREDS_FILE, 'w', encoding='utf-8') as f:
        json.dump({'client_id': client_id, 'client_secret': client_secret}, f)


def _load_creds() -> dict:
    if not os.path.isfile(_CREDS_FILE):
        return {}
    try:
        with open(_CREDS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


# ── Token-Austausch (urllib, keine externe Abhängigkeit) ──────────────────────

def _post_form(url: str, fields: dict) -> dict:
    data = urllib_parse.urlencode(fields).encode()
    req = urllib_request.Request(url, data=data, method='POST')
    req.add_header('Content-Type', 'application/x-www-form-urlencoded')
    with urllib_request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


# ── NetatmoAuth ────────────────────────────────────────────────────────────────

class NetatmoAuth:
    """Verwaltet Access- und Refresh-Token; erneuert automatisch bei Ablauf."""

    def __init__(self, client_id: str, client_secret: str,
                 access_token: str | None = None,
                 refresh_token: str | None = None,
                 expires_at: float = 0):
        self.client_id = client_id
        self.client_secret = client_secret
        self._access_token = access_token
        self._refresh_token = refresh_token
        self._expires_at = expires_at

    # -- Fabrikmethode ---------------------------------------------------------

    @classmethod
    def from_stored(cls, client_id: str, client_secret: str) -> 'NetatmoAuth | None':
        """Lädt gespeicherte Tokens; gibt None zurück wenn keine vorhanden."""
        data = _load_tokens()
        if not data or not data.get('refresh_token'):
            return None
        return cls(client_id, client_secret,
                   access_token=data.get('access_token'),
                   refresh_token=data['refresh_token'],
                   expires_at=data.get('expires_at', 0))

    # -- Token setzen ----------------------------------------------------------

    def set_tokens(self, access_token: str, refresh_token: str,
                   expires_in: int) -> None:
        self._access_token = access_token
        self._refresh_token = refresh_token
        self._expires_at = time.time() + expires_in - 60  # 60 s Puffer
        _save_tokens({
            'access_token': access_token,
            'refresh_token': refresh_token,
            'expires_at': self._expires_at,
        })

    # -- Erneuern --------------------------------------------------------------

    def refresh(self) -> None:
        if not self._refresh_token:
            raise RuntimeError('Kein Refresh-Token. Bitte neu anmelden.')
        resp = _post_form(_TOKEN_URL, {
            'grant_type': 'refresh_token',
            'refresh_token': self._refresh_token,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        })
        self.set_tokens(resp['access_token'], resp['refresh_token'],
                        resp['expires_in'])

    # -- Zugriff ---------------------------------------------------------------

    def get_access_token(self) -> str:
        if not self._access_token or time.time() >= self._expires_at:
            self.refresh()
        return self._access_token

    @property
    def refresh_token(self) -> str | None:
        return self._refresh_token


# ── Lokaler OAuth-Callback-Server ─────────────────────────────────────────────

class _CallbackHandler(BaseHTTPRequestHandler):
    """Fängt den OAuth-Redirect ab und speichert den Code im Server-Objekt."""

    def do_GET(self) -> None:
        params = urllib_parse.parse_qs(urllib_parse.urlparse(self.path).query)
        self.server.auth_code = params.get('code', [None])[0]
        self.server.auth_state = params.get('state', [None])[0]
        if self.server.auth_code:
            body = ('<html><body><h2 style="font-family:sans-serif">'
                    'Anmeldung erfolgreich — dieses Fenster kann geschlossen werden.'
                    '</h2></body></html>').encode()
        else:
            body = ('<html><body><h2 style="font-family:sans-serif">'
                    'Anmeldung fehlgeschlagen oder abgebrochen.'
                    '</h2></body></html>').encode()
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_args) -> None:
        pass  # Konsolenausgabe unterdrücken


def _do_oauth_code_flow(client_id: str, client_secret: str,
                        status_callback=None) -> tuple[str, str, int]:
    """
    Führt den OAuth2 Authorization Code Flow durch:
    1. Lokalen HTTP-Server starten
    2. Browser zum Netatmo-Auth-Endpoint öffnen
    3. Auf den Callback warten (max. 5 Minuten)
    4. Code gegen Tokens tauschen

    Gibt (access_token, refresh_token, expires_in) zurück.
    Wirft RuntimeError bei Misserfolg.
    """
    import webbrowser

    state = secrets.token_hex(16)
    params = {
        'client_id': client_id,
        'redirect_uri': _REDIRECT_URI,
        'scope': _SCOPE,
        'state': state,
    }
    auth_url = _AUTH_URL + '?' + urllib_parse.urlencode(params)

    # Lokalen Server starten
    server = HTTPServer(('localhost', _REDIRECT_PORT), _CallbackHandler)
    server.auth_code = None
    server.auth_state = None
    server.timeout = 300  # 5 Minuten

    if status_callback:
        status_callback('Browser öffnet sich für Netatmo-Anmeldung…')

    webbrowser.open(auth_url)
    server.handle_request()  # blockiert bis Redirect oder Timeout

    if not server.auth_code:
        raise RuntimeError(
            'Kein Autorisierungs-Code erhalten.\n'
            'Bitte Anmeldung im Browser abschließen (Timeout: 5 Minuten).')

    if server.auth_state != state:
        raise RuntimeError('Sicherheitsfehler: State stimmt nicht überein.')

    if status_callback:
        status_callback('Code erhalten, tausche gegen Token…')

    resp = _post_form(_TOKEN_URL, {
        'grant_type': 'authorization_code',
        'client_id': client_id,
        'client_secret': client_secret,
        'code': server.auth_code,
        'redirect_uri': _REDIRECT_URI,
        'scope': _SCOPE,
    })
    return resp['access_token'], resp['refresh_token'], resp['expires_in']


# ── NetatmoDataDownloader ──────────────────────────────────────────────────────

class NetatmoDataDownloader:
    """Lädt Messdaten von der Netatmo API."""

    def __init__(self, client_id: str, client_secret: str,
                 access_token: str | None = None):
        # Refresh-Token aus gespeicherten Daten laden falls vorhanden
        stored = _load_tokens() or {}
        self._auth = NetatmoAuth(
            client_id, client_secret,
            access_token=access_token or stored.get('access_token'),
            refresh_token=stored.get('refresh_token'),
            expires_at=stored.get('expires_at', 0),
        )
        self._last_api_request = 0.0

    def _wait_before_request(self) -> None:
        """Verhindert eine zu schnelle Folge von Netatmo-API-Aufrufen."""
        remaining = (_API_MIN_INTERVAL
                     - (time.monotonic() - self._last_api_request))
        if remaining > 0:
            time.sleep(remaining)

    def _api_get(self, url: str, params: dict) -> dict:
        token = self._auth.get_access_token()
        qs = urllib_parse.urlencode(params)
        req = urllib_request.Request(f'{url}?{qs}')
        req.add_header('Authorization', f'Bearer {token}')
        last_error = None
        for attempt in range(_API_MAX_ATTEMPTS):
            self._wait_before_request()
            try:
                self._last_api_request = time.monotonic()
                with urllib_request.urlopen(req, timeout=30) as resp:
                    return json.loads(resp.read())
            except urllib_error.HTTPError as exc:
                last_error = exc
                if exc.code not in _API_RETRY_STATUS:
                    raise
                retry_after = exc.headers.get('Retry-After')
                try:
                    delay = float(retry_after) if retry_after else 0.0
                except (TypeError, ValueError):
                    delay = 0.0
            except urllib_error.URLError as exc:
                last_error = exc
                delay = 0.0

            if attempt + 1 < _API_MAX_ATTEMPTS:
                # Exponentiell warten, aber einen Retry-After-Header beachten.
                time.sleep(max(delay, min(2 ** attempt, 30)))

        raise RuntimeError(
            f'Netatmo API nach {_API_MAX_ATTEMPTS} Versuchen nicht erreichbar: '
            f'{last_error}') from last_error

    def get_stations_data(self) -> dict:
        return self._api_get(_STATIONS_URL, {'get_favorites': 'false'})

    def extract_modules(self, devices: list) -> dict:
        """Gibt {module_id: {name, data_types}} für alle Module zurück."""
        modules: dict = {}
        for device in devices:
            modules[device['_id']] = {
                'name': device.get('module_name', 'Indoor'),
                'data_types': device.get('data_type', []),
            }
            for mod in device.get('modules', []):
                modules[mod['_id']] = {
                    'name': mod.get('module_name', mod['_id']),
                    'data_types': mod.get('data_type', []),
                }
        return modules

    def get_sensor_data(self, device_id: str, module_id: str,
                        module_name: str, sensor_type: str,
                        start: datetime.datetime,
                        end: datetime.datetime) -> list:
        """
        Lädt alle Messpunkte für einen Sensor in 30-Minuten-Auflösung.
        Paginiert automatisch über den gesamten Zeitraum.
        Gibt Zeilen im Format zurück, das _convert_netatmo_rows erwartet.
        """
        rows = []
        date_begin = int(start.timestamp())
        date_end = int(end.timestamp())

        while date_begin < date_end:
            params = {
                'device_id': device_id,
                'module_id': module_id,
                'type': sensor_type,
                'scale': '30min',
                'date_begin': date_begin,
                'date_end': date_end,
                'limit': 1024,
                'optimize': 'false',
                'real_time': 'false',
            }
            data = self._api_get(_MEASURE_URL, params)
            body = data.get('body', {})
            if not body:
                break

            # optimize=false returns {"ts_str": [val, ...], ...}
            last_ts = 0
            for ts_str, val_list in body.items():
                t = int(ts_str)
                if t <= date_begin:
                    continue
                v = val_list[0] if val_list else None
                if v is None:
                    continue
                rows.append({
                    'Sortierung': datetime.datetime.fromtimestamp(t),
                    'Module': module_name,
                    'Messwert': sensor_type,
                    'Wert': v,
                    'Einheit': _UNIT_MAP.get(sensor_type, ''),
                })
                last_ts = max(last_ts, t)

            if last_ts == 0 or last_ts <= date_begin:
                break
            date_begin = last_ts

        return rows


# ── Tkinter LoginWindow ────────────────────────────────────────────────────────

class LoginWindow:
    """
    Tkinter-Dialog für Netatmo OAuth2-Anmeldung.

    Ruft on_success(creds) mit folgenden Schlüsseln:
      client_id, client_secret, access_token

    Beim ersten Aufruf öffnet sich der Browser (einmalig).
    Danach wird der Refresh-Token automatisch verwendet.

    Konfiguration in Netatmo Connect (dev.netatmo.com/apps):
      Redirect URI muss eingetragen sein: http://localhost:9731/callback
    """

    def __init__(self, parent: tk.Misc, *, on_success):
        self._parent = parent
        self._on_success = on_success
        self._build_ui()

    # -- UI -------------------------------------------------------------------

    def _build_ui(self) -> None:
        p = self._parent
        p.title('Netatmo Anmelden')
        p.geometry('460x310')
        p.resizable(False, False)

        frame = ttk.Frame(p, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text='Netatmo Connect – Zugangsdaten',
                  font=('', 11, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0, 12))

        # Pre-fill from saved creds.json if available
        _saved = _load_creds()

        ttk.Label(frame, text='Client ID:').grid(
            row=1, column=0, sticky='w', pady=4, padx=(0, 10))
        self._id_var = tk.StringVar(value=_saved.get('client_id', ''))
        ttk.Entry(frame, textvariable=self._id_var, width=36).grid(
            row=1, column=1, sticky='ew')

        ttk.Label(frame, text='Client Secret:').grid(
            row=2, column=0, sticky='w', pady=4, padx=(0, 10))
        self._secret_var = tk.StringVar(value=_saved.get('client_secret', ''))
        ttk.Entry(frame, textvariable=self._secret_var,
                  width=36, show='*').grid(row=2, column=1, sticky='ew')

        self._status_var = tk.StringVar()
        stored = _load_tokens()
        has_refresh = bool(stored and stored.get('refresh_token'))
        if has_refresh:
            self._status_var.set(
                'Gespeicherter Token gefunden – kein Browser nötig.\n'
                'Nur Client ID + Secret eingeben, dann „Anmelden" drücken.')
            fg = '#2a7a2a'
        else:
            self._status_var.set(
                'Erstanmeldung: Browser öffnet sich einmalig.\n'
                'Redirect URI muss in dev.netatmo.com/apps eingetragen sein:\n'
                'http://localhost:9731/callback')
            fg = '#555555'

        ttk.Label(frame, textvariable=self._status_var,
                  foreground=fg, wraplength=400,
                  justify='left').grid(
            row=3, column=0, columnspan=2, pady=12, sticky='w')

        btns = ttk.Frame(frame)
        btns.grid(row=4, column=0, columnspan=2, sticky='e', pady=(8, 0))
        ttk.Button(btns, text='Abbrechen',
                   command=p.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        self._login_btn = ttk.Button(btns, text='Anmelden',
                                     command=self._start_login)
        self._login_btn.pack(side=tk.RIGHT)

    # -- Anmeldung -------------------------------------------------------------

    def _start_login(self) -> None:
        client_id = self._id_var.get().strip()
        client_secret = self._secret_var.get().strip()
        if not client_id or not client_secret:
            messagebox.showerror('Fehler',
                                 'Bitte Client ID und Client Secret eingeben.',
                                 parent=self._parent)
            return

        _save_creds(client_id, client_secret)
        self._login_btn.config(state=tk.DISABLED)

        stored = _load_tokens()
        if stored and stored.get('refresh_token'):
            # Refresh-Token vorhanden → direkt erneuern, kein Browser
            threading.Thread(
                target=self._refresh_thread,
                args=(client_id, client_secret, stored['refresh_token']),
                daemon=True).start()
        else:
            # Erstanmeldung → Browser öffnen
            threading.Thread(
                target=self._oauth_thread,
                args=(client_id, client_secret),
                daemon=True).start()

    def _refresh_thread(self, client_id: str, client_secret: str,
                        refresh_token: str) -> None:
        try:
            self._parent.after(
                0, lambda: self._status_var.set('Token wird erneuert…'))
            resp = _post_form(_TOKEN_URL, {
                'grant_type': 'refresh_token',
                'refresh_token': refresh_token,
                'client_id': client_id,
                'client_secret': client_secret,
            })
            auth = NetatmoAuth(client_id, client_secret)
            auth.set_tokens(resp['access_token'], resp['refresh_token'],
                            resp['expires_in'])
            token = resp['access_token']
            self._parent.after(0, lambda: self._on_success({
                'client_id': client_id,
                'client_secret': client_secret,
                'access_token': token,
            }))
        except Exception as exc:
            err = str(exc)
            self._parent.after(0, lambda e=err: self._on_refresh_error(
                client_id, client_secret, e))

    def _on_refresh_error(self, client_id: str, client_secret: str,
                          error: str) -> None:
        """Gespeicherter Token abgelaufen → Browser-Anmeldung starten."""
        self._status_var.set(
            f'Token abgelaufen ({error}).\nBrowser-Anmeldung wird gestartet…')
        # Alten Token löschen und neu anmelden
        _save_tokens({})
        self._login_btn.config(state=tk.NORMAL)

    def _oauth_thread(self, client_id: str, client_secret: str) -> None:
        try:
            def on_status(msg: str) -> None:
                self._parent.after(0, lambda m=msg: self._status_var.set(m))

            access_token, refresh_token, expires_in = _do_oauth_code_flow(
                client_id, client_secret, status_callback=on_status)

            auth = NetatmoAuth(client_id, client_secret)
            auth.set_tokens(access_token, refresh_token, expires_in)

            self._parent.after(0, lambda: self._on_success({
                'client_id': client_id,
                'client_secret': client_secret,
                'access_token': access_token,
            }))
        except Exception as exc:
            err = str(exc)
            self._parent.after(0, lambda e=err: (
                messagebox.showerror('Anmelde-Fehler', e,
                                     parent=self._parent),
                self._login_btn.config(state=tk.NORMAL),
            ))
