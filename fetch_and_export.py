#!/usr/bin/env python3
"""
Headless-Skript für GitHub Actions:
  - Liest Netatmo-Daten via Refresh-Token (kein Browser nötig)
  - Fügt neue Punkte zur bestehenden docs/data.json hinzu (Delta-Update)
  - Hält maximal MAX_DAYS Tage an Daten vor
  - Regeneriert docs/index.html

Umgebungsvariablen (GitHub Secrets):
  NETATMO_CLIENT_ID
  NETATMO_CLIENT_SECRET
  NETATMO_REFRESH_TOKEN   ← wird nach Ausführung aktualisiert (Rotation!)

Gibt am Ende die neue Zeile "NEW_REFRESH_TOKEN=<wert>" aus,
damit der GitHub-Actions-Step den Secret via gh CLI aktualisieren kann.

Abhängigkeiten: nur Python-Stdlib (kein requests, kein tkinter).
"""

import datetime
import json
import os
import sys
from pathlib import Path
from urllib import parse as urllib_parse
from urllib import request as urllib_request

# ── Konfiguration ──────────────────────────────────────────────────────────────
MAX_DAYS = 365 * 2          # maximaler Daten-Horizont (2 Jahre)
DELTA_DAYS = 7              # Tage zurück beim Delta-Abruf (Überlappung)
SCALE = '30min'

_TOKEN_URL = 'https://api.netatmo.com/oauth2/token'
_STATIONS_URL = 'https://api.netatmo.com/api/getstationsdata'
_MEASURE_URL = 'https://api.netatmo.com/api/getmeasure'

_UNIT_MAP: dict[str, str] = {
    'Temperature': '°C', 'Humidity': '%', 'CO2': 'ppm',
    'Pressure': 'hPa', 'Noise': 'dB', 'Rain': 'mm',
    'WindStrength': 'km/h', 'WindAngle': '°', 'GustStrength': 'km/h',
    'GustAngle': '°', 'sum_rain_1': 'mm', 'sum_rain_24': 'mm',
}

_SCRIPT_DIR = Path(__file__).resolve().parent
_DATA_FILE = _SCRIPT_DIR / 'docs' / 'data.json'
_HTML_FILE = _SCRIPT_DIR / 'docs' / 'index.html'


# ── Hilfsfunktionen ────────────────────────────────────────────────────────────

def _post_form(url: str, fields: dict) -> dict:
    data = urllib_parse.urlencode(fields).encode()
    req = urllib_request.Request(url, data=data, method='POST')
    req.add_header('Content-Type', 'application/x-www-form-urlencoded')
    with urllib_request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def _api_get(url: str, params: dict, access_token: str) -> dict:
    qs = urllib_parse.urlencode(params)
    req = urllib_request.Request(f'{url}?{qs}')
    req.add_header('Authorization', f'Bearer {access_token}')
    with urllib_request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


# ── Token-Verwaltung ───────────────────────────────────────────────────────────

def refresh_access_token(client_id: str, client_secret: str,
                         refresh_token: str) -> tuple[str, str]:
    """Gibt (access_token, new_refresh_token) zurück."""
    resp = _post_form(_TOKEN_URL, {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
    })
    return resp['access_token'], resp['refresh_token']


# ── Netatmo-Datenabruf ─────────────────────────────────────────────────────────

def get_stations(access_token: str) -> list:
    data = _api_get(_STATIONS_URL, {'get_favorites': 'false'}, access_token)
    return data.get('body', {}).get('devices', [])


def extract_modules(devices: list) -> dict:
    """Gibt {module_id: {name, data_types}} zurück."""
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


def fetch_sensor(device_id: str, module_id: str, module_name: str,
                 sensor_type: str, start: datetime.datetime,
                 end: datetime.datetime, access_token: str) -> list:
    """Paginiert über den Zeitraum und gibt normalisierte Datenpunkte zurück."""
    rows = []
    date_begin = int(start.timestamp())
    date_end = int(end.timestamp())

    while date_begin < date_end:
        params = {
            'device_id': device_id,
            'module_id': module_id,
            'type': sensor_type,
            'scale': SCALE,
            'date_begin': date_begin,
            'date_end': date_end,
            'limit': 1024,
            'optimize': 'false',
            'real_time': 'false',
        }
        data = _api_get(_MEASURE_URL, params, access_token)
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
            dt = datetime.datetime.fromtimestamp(t)
            rows.append({
                'ts': t * 1000,
                'date': dt.date().isoformat(),
                'module': module_name,
                'sensor': sensor_type,
                'value': v,
                'unit': _UNIT_MAP.get(sensor_type, ''),
            })
            last_ts = max(last_ts, t)

        if last_ts == 0 or last_ts <= date_begin:
            break
        date_begin = last_ts

    return rows


# ── Datei-Management ───────────────────────────────────────────────────────────

def load_existing_data() -> list:
    if not _DATA_FILE.exists():
        return []
    try:
        with _DATA_FILE.open('r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        print(f'WARNUNG: data.json konnte nicht gelesen werden: {exc}')
        return []


def save_data(data: list) -> None:
    _DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    # Auf MAX_DAYS begrenzen
    cutoff_ts = (
        datetime.datetime.now() - datetime.timedelta(days=MAX_DAYS)
    ).timestamp() * 1000
    data = [d for d in data if d['ts'] >= cutoff_ts]
    with _DATA_FILE.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'data.json: {len(data):,} Punkte gespeichert.')


def merge_data(existing: list, new_rows: list) -> tuple[list, int]:
    """Fügt neue Punkte ein, überspringt Duplikate. Gibt (merged, added) zurück."""
    seen: set = {(d['ts'], d['module'], d['sensor']) for d in existing}
    added = 0
    for row in new_rows:
        key = (row['ts'], row['module'], row['sensor'])
        if key not in seen:
            existing.append(row)
            seen.add(key)
            added += 1
    return existing, added


# ── HTML-Generierung ───────────────────────────────────────────────────────────

def prepare_payload(data: list) -> dict:
    """Wie prepare_chart_payload in netatmo_viewer.py (ohne tkinter)."""
    module_colors = [
        '#4dc9f6', '#f67019', '#f53794', '#537bc4',
        '#acc236', '#166a8f', '#00a950', '#8549ba',
        '#e8c83a', '#58595b',
    ]
    sensors = sorted({d['sensor'] for d in data})
    modules = sorted({d['module'] for d in data})
    units = {d['sensor']: d['unit'] for d in data}
    colors = {m: module_colors[i % len(module_colors)]
              for i, m in enumerate(modules)}

    chart_data: dict = {}
    for sensor in sensors:
        chart_data[sensor] = {}
        for module in modules:
            pts = sorted(
                [[d['ts'], d['value']] for d in data
                 if d['sensor'] == sensor and d['module'] == module],
                key=lambda x: x[0],
            )
            if pts:
                chart_data[sensor][module] = pts

    return {'sensors': sensors, 'modules': modules,
            'units': units, 'colors': colors, 'data': chart_data}


def generate_html(payload: dict) -> str:
    """Liest das HTML-Template aus netatmo_viewer.py und setzt Payload ein."""
    viewer = _SCRIPT_DIR / 'netatmo_viewer.py'
    source = viewer.read_text(encoding='utf-8')

    # Template zwischen HTML_TEMPLATE = r"""...""" extrahieren
    start_marker = 'HTML_TEMPLATE = r"""'
    end_marker = '"""\n\ndef generate_html'
    s = source.find(start_marker)
    e = source.find(end_marker)
    if s == -1 or e == -1:
        raise RuntimeError(
            'HTML_TEMPLATE in netatmo_viewer.py nicht gefunden.')

    template = source[s + len(start_marker):e]
    json_str = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    return template.replace('__PAYLOAD__', json_str)


def write_html(html: str) -> None:
    _HTML_FILE.parent.mkdir(parents=True, exist_ok=True)
    nojekyll = _HTML_FILE.parent / '.nojekyll'
    nojekyll.touch()
    _HTML_FILE.write_text(html, encoding='utf-8')
    size_kb = len(html.encode()) // 1024
    print(f'index.html geschrieben ({size_kb} KB).')


# ── Lokale Credentials-Datei (%APPDATA%\NetatmoViewer\creds.json) ─────────────

_APPDATA = Path(os.environ.get('APPDATA') or Path.home())
_CREDS_FILE = _APPDATA / 'NetatmoViewer' / 'creds.json'
_TOKENS_FILE = _APPDATA / 'NetatmoViewer' / 'tokens.json'


def _load_local_creds() -> dict:
    """Liest client_id/secret aus creds.json und refresh_token aus tokens.json."""
    creds: dict = {}
    if _CREDS_FILE.exists():
        try:
            creds.update(json.loads(_CREDS_FILE.read_text(encoding='utf-8')))
        except (OSError, json.JSONDecodeError):
            pass
    if _TOKENS_FILE.exists():
        try:
            tok = json.loads(_TOKENS_FILE.read_text(encoding='utf-8'))
            if tok.get('refresh_token'):
                creds['refresh_token'] = tok['refresh_token']
        except (OSError, json.JSONDecodeError):
            pass
    return creds


def _save_local_refresh_token(new_token: str) -> None:
    """Überschreibt den refresh_token in der lokalen tokens.json."""
    _TOKENS_FILE.parent.mkdir(parents=True, exist_ok=True)
    existing: dict = {}
    if _TOKENS_FILE.exists():
        try:
            existing = json.loads(_TOKENS_FILE.read_text(encoding='utf-8'))
        except (OSError, json.JSONDecodeError):
            pass
    existing['refresh_token'] = new_token
    _TOKENS_FILE.write_text(json.dumps(existing), encoding='utf-8')


def _save_local_creds(client_id: str, client_secret: str) -> None:
    _CREDS_FILE.parent.mkdir(parents=True, exist_ok=True)
    _CREDS_FILE.write_text(
        json.dumps({'client_id': client_id, 'client_secret': client_secret}),
        encoding='utf-8')
    print(f'Zugangsdaten gespeichert: {_CREDS_FILE}')


# ── Hauptprogramm ──────────────────────────────────────────────────────────────

def main() -> None:
    client_id = os.environ.get('NETATMO_CLIENT_ID', '').strip()
    client_secret = os.environ.get('NETATMO_CLIENT_SECRET', '').strip()
    refresh_token = os.environ.get('NETATMO_REFRESH_TOKEN', '').strip()

    # Lokale Fallback-Quelle (für direkte Ausführung ohne GitHub Actions)
    if not client_id or not client_secret or not refresh_token:
        local = _load_local_creds()
        client_id = client_id or local.get('client_id', '')
        client_secret = client_secret or local.get('client_secret', '')
        refresh_token = refresh_token or local.get('refresh_token', '')

    # Interaktiv nachfragen falls immer noch unvollständig
    if not client_id or not client_secret:
        print('Zugangsdaten nicht gefunden. Bitte eingeben (werden lokal gespeichert):')
        if not client_id:
            client_id = input('  Client ID:     ').strip()
        if not client_secret:
            client_secret = input('  Client Secret: ').strip()
        if client_id and client_secret:
            _save_local_creds(client_id, client_secret)

    if not refresh_token:
        print('FEHLER: Kein Refresh-Token gefunden.\n'
              'Bitte zuerst im Desktop-Viewer "Netatmo direkt laden" ausführen,\n'
              f'um einen Token zu erzeugen ({_TOKENS_FILE}).',
              file=sys.stderr)
        sys.exit(1)

    if not client_id or not client_secret:
        print('FEHLER: Client ID / Secret fehlen.', file=sys.stderr)
        sys.exit(1)

    # ── Token erneuern ────────────────────────────────────────────────────────
    print('Token wird erneuert…')
    access_token, new_refresh_token = refresh_access_token(
        client_id, client_secret, refresh_token)
    print('Token erfolgreich erneuert.')

    # ── Zeitraum bestimmen ───────────────────────────────────────────────────
    existing = load_existing_data()
    if existing:
        last_ts_ms = max(d['ts'] for d in existing)
        # DELTA_DAYS Überlappung für eventuelle Lücken
        start = (datetime.datetime.fromtimestamp(last_ts_ms / 1000)
                 - datetime.timedelta(days=DELTA_DAYS))
        print(f'Delta-Modus: ab {start.date()} (letzter Punkt: '
              f'{datetime.datetime.fromtimestamp(last_ts_ms / 1000).date()})')
    else:
        start = datetime.datetime.now() - datetime.timedelta(days=MAX_DAYS)
        print(f'Erstlauf: lade Daten ab {start.date()}')

    end = datetime.datetime.now()

    # ── Stationen laden ───────────────────────────────────────────────────────
    print('Lade Stationsdaten…')
    devices = get_stations(access_token)
    if not devices:
        print('FEHLER: Keine Geräte gefunden.', file=sys.stderr)
        sys.exit(1)

    device_id = devices[0]['_id']
    modules = extract_modules(devices)
    print(f'Gerät: {device_id}, Module: {len(modules)}')

    # ── Messdaten laden ───────────────────────────────────────────────────────
    new_rows: list = []
    total = sum(len(m['data_types']) for m in modules.values())
    done = 0
    for module_id, info in modules.items():
        for sensor_type in info['data_types']:
            done += 1
            print(f'  [{done}/{total}] {info["name"]} / {sensor_type}…',
                  end=' ', flush=True)
            rows = fetch_sensor(device_id, module_id, info['name'],
                                sensor_type, start, end, access_token)
            new_rows.extend(rows)
            print(f'{len(rows)} Punkte')

    # ── Merge + Speichern ─────────────────────────────────────────────────────
    merged, added = merge_data(existing, new_rows)
    print(f'Gesamt: {len(merged):,} Punkte ({added:,} neu).')
    save_data(merged)

    # ── HTML generieren ───────────────────────────────────────────────────────
    print('Generiere HTML…')
    payload = prepare_payload(merged)
    html = generate_html(payload)
    write_html(html)

    # ── Neuen Refresh-Token speichern / ausgeben ──────────────────────────────
    # Lokal aktualisieren (für direkte Ausführung); GitHub Actions liest stdout
    _save_local_refresh_token(new_refresh_token)
    print(f'NEW_REFRESH_TOKEN={new_refresh_token}')


if __name__ == '__main__':
    main()
