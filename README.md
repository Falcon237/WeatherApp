# WeatherApp
Netatmo WeatherApp

Die GitHub-Pages-Daten werden automatisch alle fünf Minuten über GitHub Actions
mit Netatmo abgeglichen. Geöffnete Live-Karussellseiten prüfen minütlich eine
kleine Versionsdatei und laden sich bei neuen veröffentlichten Daten neu.
GitHub Actions kann geplante Workflows frühestens alle fünf Minuten ausführen;
der tatsächliche Start kann sich zusätzlich verzögern.
