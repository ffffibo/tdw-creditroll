---
description: Deployment Workflow für Render.com
---

Dieser Workflow beschreibt, wie die fertiggestellte Web-Anwendung (als Flask App) live und kostenlos auf Render.com bereitgestellt wird, sodass beliebige Personen darauf zugreifen können.

## 1. Voraussetzungen prüfen
- Eine Datei `requirements.txt` muss im Hauptverzeichnis existieren (sie enthält alle nötigen Python-Pakete wie `Flask`, `pandas`, `gunicorn`, etc.).
- Eine Datei `Procfile` muss im Hauptverzeichnis existieren (Inhalt: `web: gunicorn app:app`).
- Eine Datei `.python-version` sollte im Hauptverzeichnis liegen (Inhalt: `3.11.6`), um sicherzustellen, dass Render eine stabile Python-Version nutzt (standardmäßig wird sonst oft eine zu neue Test-Version verwendet, die Fehler verursacht).
- Eine Datei `.gitignore` sollte sicherstellen, dass keine temporären Dateien oder Caches (z.B. `.DS_Store`, `__pycache__/`) hochgeladen werden.
- Alle benötigten Schriften (Fonts) und das Standard-Bild (`tdw-gallpeters.png`) müssen im Ordnerstruktur vorhanden und relativ referenziert sein (keine absoluten `/Users/...` Pfade).

## 2. GitHub-Repository erstellen & Code hochladen
1. Gehe zu [GitHub](https://github.com/) und erstelle ein neues Repository (z.B. `tdw-creditroll`). Kopiere die Repository-URL (z.B. `https://github.com/nutzer/tdw-creditroll.git`).
2. Führe die Vorbereitung in diesem Ordner aus:

// turbo
3. Git initialisieren und Dateien vormerken:
   ```bash
   git init && git add . && git commit -m "Initial commit for Render deployment" && git branch -M main
   ```

4. Den Remote-Server hinzufügen und hochladen (Ersetze `URL` durch deine kopierte URL):
   ```bash
   git remote add origin URL
   git push -u origin main
   ```

## 3. Web Service auf Render erstellen
1. Registriere dich / Melde dich auf [Render.com](https://render.com/) an.
2. Klicke im Dashboard oben rechts auf **"New +"** und wähle **"Web Service"**.
3. Wähle **"Build and deploy from a Git repository"** und klicke auf "Next".
4. Verbinde deinen GitHub-Account (falls noch nicht geschehen) und wähle das frisch erstellte Repository `tdw-creditroll` aus. Klicke auf "Connect".

## 4. Render Web Service konfigurieren
Fülle die Felder wie folgt aus:
- **Name:** Wähle einen Namen für die App (z.B. `tdw-creditroll`). Dies wird Teil der URL (z.B. `tdw-creditroll.onrender.com`).
- **Region:** Wähle "Frankfurt (EU)" für die niedrigste Latenz in Europa.
- **Branch:** `main`
- **Environment:** `Python 3`
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app:app` (falls dieser Text nicht schon aus dem Procfile übernommen wurde).
- **Plan:** Wähle den `Free` (kostenlos) Plan aus, da dieser für dieses Tool ausreicht.

## 5. Deployment abwarten & Testen
1. Scrolle nach unten und klicke auf den Button **"Create Web Service"**.
2. Render beginnt nun automatisch damit, die App zu bauen. Du kannst im Terminal-Fenster von Render zusehen, wie die Pakete installiert werden.
3. Sobald oben links der grüne Status "Live" erscheint, kannst du auf die angezeigte URL klicken (z.B. `https://tdw-creditroll.onrender.com`).
4. **Fertig!** Jeder, der diesen Link hat, kann die App nun benutzen, lokale Excel-Dateien hochladen und direkt das generierte Zip-PDF herunterladen.

> [!TIP]
> Der Free-Plan von Render schickt Apps ("Web Services") nach 15 Minuten Inaktivität in den "Schlafmodus". Wenn also 15 Minuten lang niemand die App benutzt hat, dauert es beim nächsten Aufruf der Website etwa 30 bis 60 Sekunden, bis die Startseite lädt (da der Server erst wieder hochgefahren wird). Danach läuft wieder alles in Echtzeit.
