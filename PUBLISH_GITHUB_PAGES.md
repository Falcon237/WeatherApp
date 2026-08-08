# Netatmo Viewer auf GitHub Pages veroeffentlichen

Diese App ist lokal eine Python/Tkinter-Anwendung. Fuer das Internet wird nicht die Python-GUI veroeffentlicht, sondern die daraus erzeugte statische Web-App in `docs/index.html`.

## Ablauf

1. `netatmo_viewer.py` starten.
2. Daten laden oder den Netatmo-Direktimport ausfuehren.
3. Button `Fuer GitHub Pages exportieren` klicken.
4. Dadurch entstehen:
   - `docs/index.html` als fertige Web-App
   - `docs/.nojekyll` fuer GitHub Pages
5. Dateien zu GitHub pushen.
6. In GitHub im Repository zu `Settings` > `Pages` gehen.
7. Unter `Build and deployment` waehlen:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/docs`
8. Speichern.

Danach ist die App unter diesem Muster erreichbar:

```text
https://<github-user>.github.io/<repository-name>/
```

Diesen Link kannst du versenden. Die Empfaenger sehen direkt die geladene App, nicht die Repository-Ansicht.

## In dieses Repository veroeffentlichen

Ziel-Repository:

```text
https://github.com/Falcon237/WeatherApp.git
```

Wenn dieses lokale Verzeichnis noch kein Git-Repository ist:

```powershell
git init
git add .gitignore netatmo.py netatmo_viewer.py test_auth2.py PUBLISH_GITHUB_PAGES.md docs/index.html docs/.nojekyll
git commit -m "Publish Netatmo viewer"
git branch -M main
git remote add origin https://github.com/Falcon237/WeatherApp.git
git push -u origin main
```

Wenn lokal schon ein Git-Repository existiert, aber noch kein oder ein falscher Remote gesetzt ist:

```powershell
git remote -v
git remote add origin https://github.com/Falcon237/WeatherApp.git
```

Falls `origin` schon existiert und geaendert werden soll:

```powershell
git remote set-url origin https://github.com/Falcon237/WeatherApp.git
```

Danach nur exportieren, committen und pushen:

```powershell
git add .gitignore netatmo.py netatmo_viewer.py test_auth2.py PUBLISH_GITHUB_PAGES.md docs/index.html docs/.nojekyll
git commit -m "Update Netatmo viewer app"
git push
```

Die GitHub-Pages-URL fuer dieses Repository lautet danach:

```text
https://falcon237.github.io/WeatherApp/
```

## SQLite auf GitHub

Das lokale SQLite-Archiv liegt nicht im Projektordner, sondern hier:

```text
%APPDATA%\NetatmoViewer\netatmo_archive.sqlite
```

GitHub Pages ist statisches Hosting. Das bedeutet:

- Python/Tkinter laeuft dort nicht.
- SQLite kann dort nicht beschrieben oder automatisch aktualisiert werden.
- Eine SQLite-Datei im Repository waere oeffentlich und nur als statische Download-Datei nutzbar.
- Die aktuelle Web-App nutzt fuer GitHub Pages `docs/index.html`; die Messdaten sind beim Export direkt in dieser Datei enthalten.

Wenn du die SQLite-Datei trotzdem bewusst als oeffentliche/read-only Datei im Repository ablegen willst:

```powershell
Copy-Item "$env:APPDATA\NetatmoViewer\netatmo_archive.sqlite" "docs\netatmo_archive.sqlite"
git add -f docs/netatmo_archive.sqlite
git commit -m "Add Netatmo SQLite archive"
git push
```

Danach ist die Datei abrufbar unter:

```text
https://falcon237.github.io/WeatherApp/netatmo_archive.sqlite
```

Fuer eine echte Online-Datenbank mit Schreibzugriff brauchst du statt GitHub Pages ein Backend, z.B. einen kleinen Python/FastAPI-Server mit SQLite auf Render, Railway, Fly.io oder einem eigenen Server.

## Netatmo Zugangsdaten lokal setzen

Die Zugangsdaten gehoeren nicht ins Repository. Optional kannst du sie lokal als Umgebungsvariablen setzen, damit die GUI sie vorfuellt:

```powershell
setx NETATMO_CLIENT_ID "deine-client-id"
setx NETATMO_CLIENT_SECRET "dein-client-secret"
```

Danach ein neues Terminal bzw. VS Code neu starten.

## Wichtig zu Datenschutz

`docs/index.html` enthaelt die exportierten Messdaten direkt in der Datei. Wer den Pages-Link kennt, kann diese Daten sehen. Netatmo Tokens, Client Secret und lokale Cache-Dateien werden dabei nicht gebraucht und sollten nicht mitveroeffentlicht werden.
