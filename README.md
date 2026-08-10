# Brick Archive

Private Lego-Sammlung mit automatischer Anreicherung über die Rebrickable-API.

## Funktionsweise

1. Sets werden in `data/Lego_Liste_bereinigt.xlsx` gepflegt (neue Zeile hinzufügen = neues Set)
2. Bei jedem Push auf `main` (der die Excel-Datei ändert) läuft automatisch:
   - `scripts/build.py` liest die Excel-Datei
   - fragt für neue/unbekannte Sets die Rebrickable-API ab (Bild, Jahr, Teileanzahl)
   - speichert die Antworten dauerhaft in `data/api_cache.json` (verhindert unnötige erneute Abfragen)
   - generiert `docs/index.html` – die eigentliche Homepage
3. GitHub Pages veröffentlicht `docs/` automatisch

## Setup

1. Repo auf GitHub erstellen und diese Dateien pushen
2. Unter **Settings → Secrets and variables → Actions** ein Secret `REBRICKABLE_API_KEY` mit deinem Rebrickable-API-Key anlegen
3. Unter **Settings → Pages** als Quelle "GitHub Actions" wählen
4. Bei Bedarf manuell auslösen: Tab **Actions → Build und Deploy Lego-Sammlung → Run workflow**

## Neues Set hinzufügen

1. Excel-Datei lokal öffnen, neue Zeile ausfüllen (Verkaufsspalten können leer bleiben)
2. Datei speichern, ins Repo hochladen/committen (z.B. direkt im Browser auf GitHub: Datei ersetzen)
3. Der Workflow läuft automatisch los und die Seite aktualisiert sich in ca. 1–2 Minuten
