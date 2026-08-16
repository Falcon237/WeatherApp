#!/usr/bin/env python3
"""
Passwortschutz fuer Innenraum-Module im statischen GitHub-Pages-Export.

Da die Seite rein statisch ist (kein Server/Login), gibt es keinen echten
Zugriffsschutz -- wer die Seite oeffnet, kann grundsaetzlich alle Bytes der
Datei lesen. Um die Innenraum-Messwerte (Schlafzimmer/Wohnzimmer/Zimmer)
trotzdem nicht im Klartext auszuliefern, werden sie hier client-seitig
verschluesselt (Passwort -> PBKDF2 -> HMAC-Keystream, mit HMAC-MAC zur
Erkennung falscher Passwoerter). Das ist eine einfache, aber echte
Verschluesselung (keine reine UI-Blende) -- ihre Staerke haengt direkt von
der Qualitaet des gewaehlten Passworts ab. Fuer sensible Daten waere ein
echtes Server-seitiges Login noetig; fuer private Innenraum-Klimawerte ist
das ein angemessener Kompromiss.

Nur Python-Stdlib, keine Zusatzabhaengigkeit.
"""

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
from pathlib import Path

PBKDF2_ITERATIONS = 200_000

# Module, die OEFFENTLICH bleiben (alles andere gilt als Innenraum/geschuetzt)
_PUBLIC_MODULE_RE = re.compile(
    r'outdoor|outside|au[ßs]en|garten|balkon|terrasse|regen|rain|external|extern',
    re.IGNORECASE)


def is_public_module(name: str) -> bool:
    return bool(_PUBLIC_MODULE_RE.search(name or ''))


def _password_file() -> Path:
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    folder = Path(base) / 'NetatmoViewer'
    folder.mkdir(parents=True, exist_ok=True)
    return folder / 'viewer_password.txt'


def get_viewer_password() -> str:
    """Liest das Innenraum-Passwort aus Env-Var, lokaler Datei oder Konsole.

    Reihenfolge: NETATMO_VIEWER_PASSWORD (z.B. GitHub Secret) -> lokale Datei
    -> interaktive Eingabe (wird danach lokal gespeichert).
    """
    env_pw = os.environ.get('NETATMO_VIEWER_PASSWORD', '').strip()
    if env_pw:
        return env_pw

    pw_file = _password_file()
    if pw_file.is_file():
        saved = pw_file.read_text(encoding='utf-8').strip()
        if saved:
            return saved

    pw = input(
        'Passwort fuer Innenraum-Daten (wird lokal gespeichert, '
        'nicht im Repo): ').strip()
    if pw:
        pw_file.write_text(pw, encoding='utf-8')
    return pw


def _derive_key(password: str, salt: bytes,
                iterations: int = PBKDF2_ITERATIONS) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt,
                               iterations, dklen=32)


def _keystream(key: bytes, nonce: bytes, length: int) -> bytes:
    out = bytearray()
    counter = 0
    while len(out) < length:
        block = hmac.new(
            key, nonce + counter.to_bytes(4, 'big'), hashlib.sha256).digest()
        out += block
        counter += 1
    return bytes(out[:length])


def encrypt_json(obj: dict, password: str) -> dict:
    """Verschluesselt obj (JSON-serialisierbar) mit password.

    Rueckgabe enthaelt nur Base64-Strings, sicher zum Einbetten in HTML.
    """
    salt = secrets.token_bytes(16)
    nonce = secrets.token_bytes(16)
    key = _derive_key(password, salt)
    plaintext = json.dumps(
        obj, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    ks = _keystream(key, nonce, len(plaintext))
    ciphertext = bytes(a ^ b for a, b in zip(plaintext, ks))
    mac = hmac.new(key, nonce + ciphertext, hashlib.sha256).digest()
    return {
        'salt': base64.b64encode(salt).decode('ascii'),
        'nonce': base64.b64encode(nonce).decode('ascii'),
        'iterations': PBKDF2_ITERATIONS,
        'ciphertext': base64.b64encode(ciphertext).decode('ascii'),
        'mac': base64.b64encode(mac).decode('ascii'),
    }


def split_locked_modules(chart_data: dict, modules: list) -> tuple:
    """Trennt {sensor: {module: pts}} in (oeffentlicher Teil, gesperrter Teil).

    Gibt (public_chart_data, locked_plain_dict, locked_modules) zurueck.
    """
    locked_modules = sorted(m for m in modules if not is_public_module(m))
    if not locked_modules:
        return chart_data, None, []

    locked_set = set(locked_modules)
    public_data: dict = {}
    locked_data: dict = {}
    for sensor, per_module in chart_data.items():
        pub_modules = {m: pts for m, pts in per_module.items()
                       if m not in locked_set}
        loc_modules = {m: pts for m, pts in per_module.items()
                       if m in locked_set}
        if pub_modules:
            public_data[sensor] = pub_modules
        if loc_modules:
            locked_data[sensor] = loc_modules

    return public_data, (locked_data or None), locked_modules
