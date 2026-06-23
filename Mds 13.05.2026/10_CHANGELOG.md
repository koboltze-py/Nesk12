# Changelog – Alle Versionen (Stand: 02.06.2026)

> Vollständiger Changelog aus `CHANGELOG.md` – für Rekonstruktion

---

## v3.8.0 – 02.06.2026

### Workflow-Modul – SM↔Dienstplan Abgleich
- Neues GUI-Modul `gui/workflow.py`: Stärkemeldungen (Word) ↔ Tagesdienstpläne (Excel) automatisch abgleichen
- Personen, Dienste und Zeiten werden verglichen (identisch / Abweichung / nur Stärke / nur Dienstplan)
- **Monat-Selektor**: Monat anlegen, SM/DP-Dateipfade laden; Stand wird in DB gespeichert und beim nächsten Start automatisch wiederhergestellt
- **Session-Persistenz**: neue SQLite-Tabelle `workflow_session` in `database/workflow_db.py`
- **Carmen-Abgleich**: Pro Tag markierbar mit Zeitstempel (grünes ✓), persistiert in `workflow_tag`
- **Notizen** pro Tag, ebenfalls persistiert (📝 Icon)
- **Detail-Dialog** (1200px): alle Personen des Tages mit SM/DP-Gegenüberstellung
  - SL/Dispo-Gruppe oben, Betreuer-Gruppe unten mit Trennzeile
  - Sortierung Dispo/SL: SL+DT3 → SL+DN3 → DT → DN
  - Sortierung Betreuer: T → T10 → T8 → N → N10
- DP-Notizen aus Excel-Zellen rechts von ENDE-Spalte angezeigt
- Fuzzy-Matching mit compound-Namen (El Mojahid, Moeeni Mahvelati etc.)
- Rufbereitschaft (R) und Lars Peters automatisch vom Abgleich ausgeschlossen
- Warn-Banner bei fehlenden SM/DP-Paaren

---

## v3.7.0 – 06.05.2026

### Sonderaufgaben – Vorfeldmitarbeiter & Druckdialog
- Neuer Abschnitt "👷 Vorfeldmitarbeiter" mit 3 Gruppen (09-14, 14-19, 19-00 Uhr)
- Dropdowns aus aktuellem Dienstplan (Tag + Nacht, dedupliziert)
- Excel-Export Vorfeldmitarbeiterliste (Querformat A4, S/W)
- Druckdialog: individuelle Anzahl pro Dokument (Spinner)
- Drucken via PowerShell COM-Automation (Excel.Application)

---

## v3.6.0 – 26.03.2026

### Schulungen-Modul
- Neuer Tab "👥 Mitarbeiter-Liste": Freitextsuche, Status-Filter, Schulungs-Filter
- Matrix-Tabelle: EH, Refresher, ZÜP, Ärztl., FS-K. mit Farbkodierung
- `_MitarbeiterDetailDialog`: alle 14 Schulungstypen
- `_SchulungBearbeitenDialog`: Datum-Picker + auto Gültig-bis-Berechnung
- DB-Reset + Neu-Import aus Excel (176 Mitarbeiter)

---

## v3.5.1 – 21.03.2026

### Tab-Design-Harmonisierung
- Einheitliches Tab-Design über alle 11 GUI-Dateien
- Primärfarbe `#1565a8`, Segoe UI Schriftart
- Full-Page-Tabs und Nested-Tabs unterschiedlich gestylt

### Seitennavigation – Fade-Animation
- Sanfte Fade-In-Animation (180ms, OutCubic) bei Seitenwechsel
- `QGraphicsOpacityEffect` + `QPropertyAnimation`

### Mitarbeiter – Reorganisation
- Tab "Dokumente" → "🗂️ Verwaltung"
- Nur noch 2 Top-Level-Tabs: Verwaltung | Übersicht

### Sonderaufgaben
- Treeview: "Gespeicherte Aufgaben" → "📁 Dienstpläne"
- Button "📂 Ordner öffnen"
- Button "↩️ Wiederherstellen" mit Dropdown

---

## v3.5.0 – 21.03.2026

### Passagieranfragen (neues Modul)
- Neues Widget `gui/passagieranfragen.py` (Sidebar-Index 16)
- Outlook-Posteingang direkt in der App (letzte 75 E-Mails)
- Automatische Datenextraktion: Name, E-Mail, Flugnummer, Datum
- 4 Antwort-Szenarien (Bestätigung, Fehlende Infos, Parkplatz, PRM-Info)
- Personalisierte Begrüßung ("Sehr geehrter Herr Müller,")
- Outlook-Entwurf via win32com mit DRK-Logo als CID-Inline-Bild
- Betreff: "PRM-Service – Flughafen Köln/Bonn | Name | Flug | Datum"

---

## v3.4.5 – 20.03.2026

### Sidebar – Animiertes Logo
- `_NeskLogoWidget` ersetzt statisches Logo
- Teal-Ring + Gold-Ring + Shimmer-Effekt, 33 FPS
- Sidebar scrollbar via QScrollArea

### Übergabe – HTML-E-Mail
- Vollständiges Redesign: DRK-roter Header, farbige Boxen, HTML-Tabellen
- Neue Sektion "Patienten DRK Station"

---

## v3.4.4 – 14.03.2026

### Dienstplan
- Doppeltes Speichern nach Export entfernt
- Speicherort-Button entfernt; direkter Datei-Dialog

---

## v3.4.3 / v1.1 – 12.03.2026

### Übergabe – Verspätungen
- Nachtdienst: Vortag-Verspätungen nicht mehr automatisch
- Manuelle Verspätungen: Datum-Feld auswählbar
- Dedup-Logik für Vortag-Einträge
- E-Mail: Datum pro Verspätung

### Backup-System
- Neue Tages-Ordner-Struktur: `db_backups/YYYY-MM-DD/`
- Max. 5 Backups/Tag je DB, max. 7 Tages-Ordner

---

## v3.4.2 – 12.03.2026

### Übergabe Bugfixes
- Auto-Einträge nach Speichern sichtbar
- Doppelungen beim Reload behoben
- E-Mail-Dialog: Datum-Filter
- "Aus Verspätungen wählen": letzte 7 Tage

### Einsätze
- Sortierung: aufsteigend (chronologisch)
- Alle Felder optional (keine Pflichtfelder)

### verspaetung_db.py
- `lade_verspaetungen_letzter_zeitraum(tage: int = 7)` neu

---

## v3.4.1 – 11.03.2026

### Hilfe-Dialog
- Neuer Tab "📸 Vorschau": Galerie aller 14 Seiten als Screenshots
- `grab_all_screenshots()` in MainWindow

### docs/BENUTZERANLEITUNG.md (neu)
- Vollständige Benutzeranleitung (17 Abschnitte)

---

## v3.4.0 – 11.03.2026

### Dienstliches
- Medikamentengabe als Tabelle (DB: `medikamente`-Tabelle)
- CASCADE-FK auf `patienten_id`

---

## v3.3.x – Früheres 2026

### Notizen-Modul (Kernfeature)
- `notizen.db` neu angelegt
- Dashboard: Notizen-Panel mit +5/-5-Tage-Fenster
- Archiv-Button + Archiv-Dialog
- "Erledigt rückgängig"-Button (`als_offen()`)
- `lade_fenster()`: heute-5 bis heute+10 nach `datum`-Feld

### Vorkommnisse
- Auto-Save beim E-Mail-Dialog wenn ungespeichert und Flugnummer vorhanden

### Handys-Modul
- `handys.db` komplett neu
- 4 Tabs: Übersicht, Historien, Berichte, Einstellungen
- Trigger für automatisches `geaendert_am`
- Word-Berichte, Excel-Export, E-Mail

### Turso-Sync
- Vollständige `TABLE_MAP` für alle DBs
- `push_row()` nach jedem Write
- `pull_all()` beim Start
- Auto-Sync Thread alle 30 Sekunden

---

## Version-Nummerierung

| Version | Bedeutung |
|---|---|
| 3.7.0 | Aktuell (Stand 13.05.2026) |
| 3.x.0 | Major Feature |
| 3.x.y | Bugfix / kleines Feature |
