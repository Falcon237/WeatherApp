"""
Versioniertes Daten-Archiv im Repo (Backup + Datenbasis für GitHub Actions).

Layout:  archive/<JAHR>.json.gz
Inhalt:  {"v": 1, "series": [{"module", "sensor", "unit", "pts": [[ts_ms, wert], ...]}, ...]}

Pro Jahr eine Datei: abgeschlossene Jahre ändern sich nie (kein Repo-Wachstum),
nur das laufende Jahr wird neu geschrieben. Die Ausgabe ist deterministisch
(sortiert, gzip ohne Zeitstempel), unveränderte Jahre erzeugen also keinen Diff.
"""

import datetime
import gzip
import json
from collections import defaultdict
from pathlib import Path

ARCHIVE_DIR = Path(__file__).resolve().parent / 'archive'


def _year(ts_ms: int) -> int:
    return datetime.datetime.fromtimestamp(ts_ms / 1000).year


def load_archive(directory: Path = ARCHIVE_DIR) -> list:
    """Liest alle Jahresdateien und gibt flache Zeilen (wie data.json) zurück."""
    rows = []
    for path in sorted(directory.glob('*.json.gz')):
        try:
            with gzip.open(path, 'rt', encoding='utf-8') as f:
                obj = json.load(f)
        except (OSError, ValueError) as exc:
            print(f'WARNUNG: {path.name} nicht lesbar: {exc}')
            continue
        for s in obj.get('series', []):
            for ts, value in s['pts']:
                rows.append({
                    'ts': ts,
                    'date': datetime.datetime.fromtimestamp(ts / 1000).date().isoformat(),
                    'module': s['module'],
                    'sensor': s['sensor'],
                    'value': value,
                    'unit': s['unit'],
                })
    return rows


def save_archive(rows: list, directory: Path = ARCHIVE_DIR) -> list:
    """Schreibt geänderte Jahresdateien. Gibt die Namen der geschriebenen Dateien zurück."""
    by_year = defaultdict(lambda: defaultdict(dict))
    for d in rows:
        key = (d['module'], d['sensor'], d.get('unit', ''))
        by_year[_year(d['ts'])][key][d['ts']] = d['value']

    directory.mkdir(parents=True, exist_ok=True)
    written = []
    for year, series in sorted(by_year.items()):
        obj = {'v': 1, 'series': [
            {'module': m, 'sensor': s, 'unit': u,
             'pts': [[ts, vals[ts]] for ts in sorted(vals)]}
            for (m, s, u), vals in sorted(series.items())
        ]}
        raw = json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
        data = gzip.compress(raw, compresslevel=9, mtime=0)
        path = directory / f'{year}.json.gz'
        if path.exists() and path.read_bytes() == data:
            continue
        path.write_bytes(data)
        written.append(path.name)
    return written
