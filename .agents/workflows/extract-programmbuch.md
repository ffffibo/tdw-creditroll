---
description: Programmbuch Projekt-Daten Extraktion (InDesign)
---

Dieser Workflow beschreibt, wie Daten aus den "Projekt"-Textrahmen eines InDesign-Programmbuchs extrahiert und als tabellarische TSV-Liste ausgegeben werden.

## 1. Voraussetzungen im InDesign-Dokument
- Die Textrahmen müssen mit den Objektformaten **"Projekt DE"** oder **"Projekt EN"** versehen sein.
- Innerhalb der Rahmen müssen spezifische Absatzformate (z.B. `01. Projekt Monat + Datum DE`, `02. Projekt Titel DE`, etc.) verwendet werden.
- Jedes Projekt benötigt eine ID im Format `**ID001` im ersten Absatz.

## 2. Skript ausführen
1. Öffne das InDesign-Dokument, aus dem die Daten extrahiert werden sollen.
2. Öffne das **Skripte-Bedienfeld** (Fenster > Hilfsprogramme > Skripte).
3. Navigiere zum Skript `c25-programmbuch-extractData-01.jsx`.
4. Doppelklicke auf das Skript, um es zu starten.

## 3. Ergebnis verarbeiten
- Das Skript erstellt automatisch einen **neuen Textrahmen** auf der letzten Seite des Dokuments.
- Dieser Rahmen enthält alle extrahierten Daten im **TSV-Format** (Tab-Separated Values).
- Du kannst diesen Text kopieren und direkt in **Excel** oder **Google Sheets** einfügen, um eine strukturierte Tabelle zu erhalten.

## 4. Troubleshooting
- **Fehlende ID:** Wenn ein Projekt keine `**ID...` Kennung hat, wird es ignoriert.
- **Merge-Logik:** Das Skript verbindet deutsche und englische Texte automatisch über die ID. Stelle sicher, dass die IDs in beiden Sprachen identisch sind.
- **Zeilenumbrüche:** Interne Zeilenumbrüche in den Texten werden vom Skript automatisch durch Leerzeichen ersetzt, um das Tabellenformat nicht zu zerstören.
