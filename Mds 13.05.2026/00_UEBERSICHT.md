# Nesk3 – Gesamtübersicht (Stand: 02.06.2026)

## Was ist Nesk3?

**Nesk3** ist eine Desktop-Verwaltungsanwendung für die **DRK Erste-Hilfe-Station Flughafen Köln/Bonn**.  
Sie verwaltet Mitarbeiter, Dienstpläne, Fahrzeuge, Passagieranfragen, Vorfälle, Schulungen, Handys und mehr.

- **Programmiersprache:** Python 3.13  
- **GUI-Framework:** PySide6 (Qt 6)  
- **Datenbank lokal:** SQLite (mehrere `.db`-Dateien in `database SQL/`)  
- **Cloud-Sync:** Turso (libSQL, read-replica)  
- **Design:** SAP Fiori inspiriert (Farben in `config.py`)  
- **Version:** 3.8.0  

---

## Dateienstruktur (Wurzelverzeichnis)

```
Nesk3/
├── main.py                  ← Einstiegspunkt (App-Start, Backups, Migrations)
├── config.py                ← Alle Pfade, Konstanten, API-Keys
├── requirements.txt         ← Python-Abhängigkeiten
├── Nesk3.spec               ← PyInstaller Build-Konfiguration
├── start_nesk.ps1           ← PowerShell-Startskript
│
├── gui/                     ← Alle PySide6-Widgets (Fenster, Dialoge, Tabs)
├── functions/               ← Business-Logik, DB-Zugriff für spezialisierte DBs
├── database/                ← ORM/Connection für nesk3.db (Haupt-DB)
├── backup/                  ← Backup-Manager (SQL, Gemeinsam, ZIP, DRK-Daten)
│
├── database SQL/            ← SQLite-Datenbankdateien (*.db)
├── Daten/                   ← Nutzdaten (Word, Excel, PDFs, Bilder)
├── json/                    ← JSON-Einstellungsdateien
├── docs/                    ← Technische Dokumentation
└── Mds 13.05.2026/          ← Diese Rekonstruktionsdokumentation
```

---

## Dokumentdateien in diesem Ordner

| Datei | Inhalt |
|---|---|
| `00_UEBERSICHT.md` | Diese Datei – Gesamtübersicht |
| `01_SETUP_UND_START.md` | Installation, Abhängigkeiten, Starten |
| `02_KONFIGURATION.md` | config.py vollständig erklärt |
| `03_ARCHITEKTUR.md` | Modul-Übersicht, Schichten, Datenfluss |
| `04_DATENBANKEN.md` | Alle SQLite-Schemata vollständig |
| `05_GUI_MODULE.md` | Alle GUI-Widgets und ihre Funktion |
| `06_FUNCTIONS_MODULE.md` | Alle functions/-Module und ihre API |
| `07_BACKUP_SYSTEM.md` | Backup-Logik, Startup-Backup, manuell |
| `08_TURSO_SYNC.md` | Cloud-Sync mit Turso |
| `09_FEATURES_DETAIL.md` | Detailbeschreibung aller Hauptfeatures |

---

## Technologie-Stack

| Komponente | Technologie |
|---|---|
| GUI | PySide6 ≥ 6.6.0 |
| Lokale DB | SQLite 3 (WAL-Modus) |
| Cloud-DB | Turso (libSQL) |
| Excel-Export | openpyxl ≥ 3.1.0 |
| Word-Dokumente | python-docx ≥ 1.0.0 |
| PDF-Verarbeitung | pypdf ≥ 4.0.0 |
| E-Mail (Outlook) | win32com.client (pywin32) |
| KI-Transkription | Google Gemini API |
| Build (EXE) | PyInstaller |

---

## Module-Überblick

### `gui/` – Alle Fenster und Dialoge

| Modul | Beschreibung |
|---|---|
| `main_window.py` | Hauptfenster mit Sidebar-Navigation |
| `splash_screen.py` | Animierter Startbildschirm |
| `dashboard.py` | Startseite: Kacheln, Kalender, Notizen, Termine |
| `mitarbeiter.py` | Mitarbeiterverwaltung (CRUD, Suche) |
| `dienstplan.py` | Dienstplan-Ansicht und Bearbeitung |
| `fahrzeuge.py` | Fahrzeugverwaltung (Status, Schäden, Termine) |
| `vorkommnisse.py` | Vorkommnisberichte (Word-Export, E-Mail) |
| `beschwerden.py` | Beschwerdemanagement |
| `passagiere.py` | Passagier-Datenbank |
| `passagieranfragen.py` | Anfragen-Tracking |
| `schulungen_kalender.py` | Schulungskalender |
| `handys_widget.py` | Handyverwaltung (4 Tabs, Berichte, E-Mail) |
| `uebergabe.py` | Schichtübergabe-Protokolle |
| `mitarbeiter_dokumente.py` | Dokumente je Mitarbeiter |
| `dokument_browser.py` | Dokumenten-Browser |
| `dienstliches.py` | Dienstliche Anweisungen |
| `aufgaben.py` / `aufgaben_tag.py` | Aufgaben-Verwaltung |
| `sonderaufgaben.py` | Sonderaufgaben |
| `checklisten.py` | Checklisten |
| `code19.py` | Code-19-Protokolle |
| `einstellungen.py` | App-Einstellungen |
| `backup_widget.py` | Backup-Verwaltung (UI) |
| `bericht.py` | Berichtsgenerierung |
| `telefonnummern.py` | Telefonbuch |
| `call_transcription.py` | Anruf-Transkription (Gemini AI) |
| `schadensbericht_dialog.py` | Dialog für Handyschadens-/Verlustberichte |
| `hilfe_dialog.py` | Hilfedialog |
| `slot_machine.py` / `slot_symbols.py` | Easter Egg (Wunderrad) |

### `functions/` – Business-Logik

| Modul | DB-Datei | Beschreibung |
|---|---|---|
| `notizen_db.py` | `notizen.db` | Dashboard-Notizen CRUD |
| `vorkommnisse_db.py` | `vorkommnisse.db` | Vorkommnisse CRUD + Turso-Push |
| `beschwerden_db.py` | `beschwerden.db` | Beschwerden CRUD + Turso-Push |
| `handys_db.py` | `handys.db` | Diensthandys CRUD |
| `handys_bericht.py` | – | Word-Berichte für Handys |
| `handys_email.py` | – | Outlook-E-Mail für Handys |
| `handys_excel_export.py` | – | Excel-Export Geräteübersicht |
| `verspaetung_db.py` | nesk3.db | Verspätungen |
| `verspaetung_functions.py` | – | Excel-/Word-Export Verspätungen |
| `mitarbeiter_functions.py` | nesk3.db | Mitarbeiter-Logik |
| `mitarbeiter_dokumente_functions.py` | – | Dokument-Verwaltung |
| `mitarbeiter_sync.py` | – | MA-Sync zwischen DBs |
| `dienstplan_functions.py` | nesk3.db | Dienstplan-Logik |
| `dienstplan_parser.py` | – | Word-Dienstplan einlesen |
| `dienstplan_html_export.py` | – | HTML-Export |
| `fahrzeug_functions.py` | nesk3.db | Fahrzeug-Logik |
| `schulungen_db.py` | nesk3.db | Schulungen |
| `uebergabe_functions.py` | nesk3.db | Übergabe-Protokolle |
| `archiv_functions.py` | `archiv.db` | Archiv-Modul |
| `stellungnahmen_db.py` | – | Stellungnahmen |
| `stellungnahmen_html_export.py` | – | HTML-Export |
| `mail_functions.py` | – | Gemeinsame E-Mail-Hilfsfunktionen |
| `settings_functions.py` | nesk3.db | App-Einstellungen lesen/schreiben |
| `telefonnummern_db.py` | nesk3.db | Telefonbuch |
| `emobby_functions.py` | – | E-Mobby Integration |
| `psa_db.py` | nesk3.db | PSA (Persönl. Schutzausrüstung) |
| `dokument_archiv.py` | – | Dokument-Archivierung |
| `dienstanweisungen_db.py` | – | Dienstanweisungen |
| `call_transcription_db.py` | – | Transkriptions-Datenbank |
| `staerkemeldung_export.py` | – | Stärkemeldung-Export |

### `database/` – Haupt-ORM

| Modul | Beschreibung |
|---|---|
| `connection.py` | SQLite-Verbindung zu nesk3.db |
| `migrations.py` | Schema-Erstellung + ALTER-TABLE-Migrationen |
| `models.py` | Datenzugriffsschicht (Queries) |
| `turso_sync.py` | Turso Push/Pull |
| `pax_db.py` | Passagier-Datenbank |
| `sanmat_db.py` | SanMat-Datenbank |
| `sonderaufgaben_db.py` | Sonderaufgaben-DB |
| `export_historie_db.py` | Export-Historienspeicherung |
