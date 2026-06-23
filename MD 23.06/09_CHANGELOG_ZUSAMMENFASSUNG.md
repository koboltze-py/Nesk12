# Changelog-Zusammenfassung (v3.4.0 – v3.8.0)

> Vollständiger Changelog: `CHANGELOG.md` im Projektverzeichnis

---

## v3.8.0 (aktuell, 23.06.2026)

Aktuelle Version. Details in CHANGELOG.md.

---

## v3.7.0 (06.05.2026)

### Sonderaufgaben – Vorfeldmitarbeiter & verbesserter Druckdialog

**`gui/sonderaufgaben.py`**
- Neuer Abschnitt „👷 Vorfeldmitarbeiter": 3 Gruppen (09-14, 14-19, 19-00), je 3 Dropdowns
- Dropdowns aus aktuellem Dienstplan befüllt (Tag + Nacht dedupliziert)
- Suffix: Schichttyp (T/T10/N/N10), Bulmorfahrer (B), E-Mobby-Fahrer (EM)
- Excel-Export Vorfeldmitarbeiterliste: Querformat A4, s/w, Spalten: Datum/Uhrzeit | MA1 | MA2 | MA3
- Druckdialog überarbeitet: Anzahl-Spinner pro Dokument (Standard: SA=2×, VM-Liste=3×)
- Drucken via PowerShell COM-Automation (Excel.Application, wartet auf Abschluss)

---

## v3.6.0 (26.03.2026)

### Schulungen-Modul

**`functions/schulungen_db.py`**
- `lade_mitarbeiter_mit_schulungen()`: MA + aktuellste Einträge pro Typ

**`gui/schulungen_kalender.py`**
- Neuer Tab „👥 Mitarbeiter-Liste"
- Freitextsuche, Status-Filter (7 Optionen), Schulungs-Filter (14 Typen)
- Matrix-Tabelle mit Farbkodierung
- `_MitarbeiterDetailDialog` (Doppelklick)
- `_SchulungBearbeitenDialog`: auto Gültig-bis-Berechnung nach Typ

---

## v3.5.1 (21.03.2026)

### Tab-Design Harmonisierung

- Einheitliches Tab-Design `#1565a8` in allen 11 GUI-Dateien
- Full-Page-Tabs: `setDocumentMode(True)` + 3px Underline
- Nested-Tabs: 2px Underline + `#e8ecf0`-Hintergrund

### Fade-Animation

**`gui/main_window.py`**
- 180ms OutCubic Fade-In bei Seitenwechsel
- QGraphicsOpacityEffect + QPropertyAnimation (0→1)
- Effect nach Animation entfernt (kein Performance-Impact)

### Mitarbeiter-Reorganisation

**`gui/mitarbeiter.py`**
- Tab „Dokumente" → „🗂️ Verwaltung"
- Nur noch 2 Tabs: Verwaltung | Übersicht

**`gui/mitarbeiter_dokumente.py`**
- Sidebar: Emojis → Bullet Points
- Trennlinie + neue Einträge: 🖨️ Ausdrucke, 🤒 Krankmeldungen
- `_zeige_sonderkategorie()`: DokumentBrowserWidget für Ausdrucke/Krankmeldungen

### Sonderaufgaben – Ordner & Wiederherstellen

**`gui/sonderaufgaben.py`**
- Treeview-Überschrift: „Gespeicherte Aufgaben" → „📁 Dienstpläne"
- „📂 Ordner öffnen"-Button
- „↩️ Wiederherstellen"-Button: Dropdown + Formular-Befüllung
- `_restore_last()`, `_load_from_excel()`: liest alle Felder aus gespeicherter Excel

---

## v3.5.0 (21.03.2026)

### Passagieranfragen – Neues Sidebar-Modul

**`gui/passagieranfragen.py`** (neu)
- Outlook-Posteingang (letzte 75 E-Mails) direkt in der App
- Auto-Extraktion: Name, E-Mail, Flugnummer, Datum, Rückflug
- 4 Antwort-Szenarien (Bestätigung, Fehlend, Parkplatz, Info)
- Outlook-Entwurf via win32com, DRK-Logo als CID-Inline-Bild
- Betreff auto-erstellt: „PRM-Service – FH Köln/Bonn | Name | Flug | Datum"

**`gui/main_window.py`**
- PassagieranfragenWidget als Index 16 registriert

---

## v3.4.5 (20.03.2026)

### Sidebar – Animiertes Logo

**`gui/main_window.py`**
- `_NeskLogoWidget`: Teal-Ring (vorwärts) + Gold-Ring (rückwärts) + Shimmer
- QTimer 30ms (~33 FPS), Hintergrundfarbe exakt `#354a5e`
- Sidebar scrollbar (QScrollArea)

### Übergabe – HTML-E-Mail überarbeitet

**`gui/uebergabe.py`**
- HTML-E-Mail: DRK-roter Header-Banner, farbige Sektionen, Tabellen
- Info-Tabelle im Header: Datum, Schicht, Ersteller, Patienten
- Neue Sektion „Patienten DRK Station"
- Bugfix `NameError pat_html`

---

## v3.4.4 (14.03.2026)

### Dienstplan – Word-Export

**`gui/dienstplan.py`**
- Doppeltes Speichern entfernt
- Direkter Datei-Speicherdialog (kein Zwischen-Dialog)

---

## v3.4.3 (12.03.2026)

### Übergabe – Verspätungen

**`gui/uebergabe.py`**
- Nachtdienst: Vortag-Verspätungen nicht mehr automatisch angezeigt
- Blaue Einträge immer sichtbar (auch nach Speichern)
- Manueller ➕-Button speichert direkt in `verspaetungen.db`
- Datum-Feld, editierbare Sollzeit
- Bugfix: Duplikate nach `(mitarbeiter, dienstbeginn)` dedupliziert
- E-Mail: Datum pro Eintrag

### Backup-System

**`main.py`**
- Tages-Ordner `db_backups/YYYY-MM-DD/`
- Max. 5 Backups/DB/Tag
- Max. 7 Tages-Ordner

---

## v3.4.2 (12.03.2026)

### Übergabe – Bugfixes

- Auto-geladene Vortag-Verspätungen nach Speichern sichtbar
- Blaue Einträge werden in `uebergabe_verspaetungen` gespeichert
- Duplikate via `saved_keys`-Set verhindert
- E-Mail-Dialog: Datum-Filter für Verspätungen
- Sonderaufgaben: `_combo_to_line()` als eigenständige Methode

---

## v3.4.1 (11.03.2026)

### Hilfe-Dialog – Screenshots & Benutzeranleitung

**`gui/hilfe_dialog.py`**
- Tab „📸 Vorschau": 2-spaltige Kachelgalerie aller App-Seiten
- `_ScreenshotCard`, `_FullscreenPreview`
- Screenshots in `Daten/Hilfe/screenshots/{idx:02d}.png`

**`gui/main_window.py`**
- `grab_all_screenshots(callback)`: Timer-basierter Screenshot-Durchlauf (300ms/Seite)

**`docs/BENUTZERANLEITUNG.md`** (neu)
- Vollständige Benutzeranleitung (17 Abschnitte, Mermaid-Diagramme)

---

## v3.4.0 (11.03.2026)

### Dienstliches – Medikamentengabe als Tabelle

**`gui/dienstliches.py`**
- `medikamente`-Tabelle in `patienten_station.db` (id, patienten_id, medikament, dosis, applikation)
- `_build_grp_medikamente()`, Methoden für Hinzufügen/Entfernen
- Word-Export: Medikamenten-Tabelle in Abschnitt 7
