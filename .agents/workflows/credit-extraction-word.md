---
description: Credit-Daten Extraktion (Word zu Excel)
---

Dieser Workflow beschreibt, wie Credits (Namen und Rollen) aus Word-Dokumenten extrahiert und in eine strukturierte Excel-Tabelle für die PDF-Generierung überführt werden.

## 1. Verzeichnisse vorbereiten
- Lege alle Word-Dokumente (`.docx`) in den Ordner `in/`.
- Stelle sicher, dass die Dokumente Tabellen oder strukturierte "Creditblöcke" (Rolle: Name) enthalten.

## 2. Scan-Skript ausführen
Führe das Python-Skript aus, um die Dokumente zu analysieren:
```bash
python3 scan_docx.py
```
- Das Skript durchsucht alle Dateien im `in/` Ordner.
- Erzeugt eine CSV/XLS Datei mit den gefundenen Namen und Rollen.

## 3. Daten veredeln (Optional)
Falls die Namen bereinigt oder Typen (Mensch, Firma) klassifiziert werden sollen:
```bash
python3 reclassify_typ.py
```

## 4. Ergebnis prüfen
- Die finale Excel-Datei (z.B. `tdw-creditroll-[datum].xlsx`) befindet sich im Hauptverzeichnis.
- Diese Datei kann nun im **Credit PDF Generator** (Web-App) hochgeladen werden.

## Troubleshooting
- **Keine Daten gefunden:** Prüfe, ob die Dateiendung `.docx` kleingeschrieben ist (nicht `.DOCX`).
- **Formatierung:** Das Skript ist darauf optimiert, Tabellen zu lesen. Falls Daten im Fließtext stehen, müssen sie ggf. manuell in die Excel übertragen werden.
