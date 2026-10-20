# WeatherApp
Netatmo WeatherApp

Die GitHub-Pages-Daten werden automatisch alle fünf Minuten über GitHub Actions
mit Netatmo abgeglichen. Geöffnete Live-Karussellseiten prüfen minütlich eine
kleine Versionsdatei und laden sich bei neuen veröffentlichten Daten neu.
GitHub Actions kann geplante Workflows frühestens alle fünf Minuten ausführen;
der tatsächliche Start kann sich zusätzlich verzögern.

Für den Workflow müssen `NETATMO_CLIENT_ID`, `NETATMO_CLIENT_SECRET`,
`NETATMO_REFRESH_TOKEN` und `NETATMO_VIEWER_PASSWORD` als Repository-Secrets
gesetzt sein. `NETATMO_VIEWER_PASSWORD` ist dasselbe Passwort, das zum
Entsperren der verschlüsselten Innenraumdaten in der Website verwendet wird.
`GH_PAT` muss ein Fine-grained Personal Access Token sein, auf das Repository
`WeatherApp` beschränkt sein und die Repository-Berechtigung **Secrets: Read and
write** besitzen, damit der Workflow den von Netatmo erneuerten Refresh-Token
zurückschreiben kann.
