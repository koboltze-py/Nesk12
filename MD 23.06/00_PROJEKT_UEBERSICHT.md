# Nesk3 – Projektübersicht (Stand: 23.06.2026 / v3.8.0)

## Allgemein

| Eigenschaft | Wert |
|---|---|
| **Anwendungsname** | Nesk3 – DRK Flughafen Köln |
| **Version** | 3.8.0 |
| **Organisation** | Deutsches Rotes Kreuz – Erste-Hilfe-Station Flughafen Köln/Bonn |
| **Sprache** | Python 3.13 |
| **GUI-Framework** | PySide6 (Qt 6) |
| **Datenbank (lokal)** | SQLite (WAL-Modus) |
| **Datenbank (Cloud)** | Turso (libSQL / SSOT-Synchronisierung) |
| **Build-Tool** | PyInstaller |
| **Git Remote** | https://github.com/koboltze-py/Nesk12.git |

## Zweck

Nesk3 ist eine Desktop-Verwaltungssoftware für den operativen Betrieb der DRK-Erste-Hilfe-Station am Flughafen Köln/Bonn. Sie verwaltet:
- Mitarbeiter & Dienstpläne
- Übergabeprotokolle (Tag/Nacht)
- Fahrzeugflotte & Wartungstermine
- Einsätze & PAX-Betreuung (Passagiere)
- Vorkommnisse & Berichte
- Schulungsnachweise
- Beschwerden & Passagieranfragen (PRM-Service)
- Verspätungsdokumentation
- Telefonverzeichnis
- Checklisten & Sonderaufgaben
- Handyübersicht & E-Mails
- Dienstliche Dokumente
- Backup- und Synchronisierungssystem

## Technologiestack

```
Python 3.13
├── PySide6          – GUI (Qt 6, SAP Fiori Design)
├── SQLite (sqlite3) – Lokale Datenbanken
├── libsql-client    – Turso Cloud-Datenbank
├── openpyxl         – Excel-Export/Import
├── python-docx      – Word-Dokument-Export
├── pypdf            – PDF-Verarbeitung
├── win32com         – Outlook-Integration (E-Mail, COM)
├── PyInstaller      – EXE-Build
└── Gemini API       – KI-Integration (Transkription, Extraktion)
```

## Projektverzeichnisstruktur

```
Nesk3/
├── main.py                    ← Einstiegspunkt
├── config.py                  ← Alle Konfigurationswerte (Pfade, DB, Farben, API-Keys)
├── Nesk3.spec                 ← PyInstaller Build-Konfiguration
├── requirements.txt           ← Python-Abhängigkeiten
│
├── gui/                       ← PySide6 GUI-Module
│   ├── main_window.py         ← Hauptfenster + Sidebar-Navigation
│   ├── splash_screen.py       ← Ladebildschirm
│   ├── dashboard.py           ← Dashboard (Statistiken, Kalender)
│   ├── aufgaben.py            ← Nachtdienst-Sonderaufgaben
│   ├── aufgaben_tag.py        ← Tagdienst-Aufgaben
│   ├── aufgaben_haupt.py      ← Haupt-Aufgaben-Container
│   ├── sonderaufgaben.py      ← Sonderaufgaben (Vorfeldmitarbeiter, Druckdialog)
│   ├── fahrzeuge.py           ← Fahrzeugverwaltung
│   ├── mitarbeiter.py         ← Mitarbeiterverwaltung
│   ├── mitarbeiter_dokumente.py← Dokumentenbrowser
│   ├── telefonnummern.py      ← Telefonverzeichnis
│   ├── beschwerden.py         ← Beschwerden-Management
│   ├── passagiere.py          ← Passagiere-Container (Anfragen + Beschwerden)
│   ├── passagieranfragen.py   ← PRM-Passagieranfragen & Outlook-Integration
│   ├── dienstliches.py        ← Dienstliche Dokumente
│   ├── dienstplan.py          ← Dienstplan (Schichten, Word-Export)
│   ├── uebergabe.py           ← Übergabeprotokolle
│   ├── vorkommnisse.py        ← Vorkommnis-Berichte
│   ├── checklisten.py         ← Checklisten
│   ├── schulungen_kalender.py ← Schulungskalender & Mitarbeiter-Schulungsmatrix
│   ├── bericht.py             ← Berichte
│   ├── workflow.py            ← Workflow-Management
│   ├── code19.py              ← Code-19 (intern)
│   ├── einstellungen.py       ← App-Einstellungen
│   ├── backup_widget.py       ← Backup-UI
│   ├── hilfe_dialog.py        ← Hilfe/Info-Dialog
│   ├── handys_widget.py       ← Handy-Übersicht
│   ├── laufzettel_dialog.py   ← Laufzettel-Dialog
│   ├── schadensbericht_dialog.py← Schadensberichte
│   ├── dokument_browser.py    ← Datei-Browser-Widget
│   ├── call_transcription.py  ← Anruf-Transkription
│   ├── slot_machine.py        ← Easter-Egg Slot-Machine
│   ├── slot_symbols.py        ← Slot-Symbole
│   └── sanmat/                ← SANMAT-spezifische Module
│
├── database/                  ← Datenbankschicht
│   ├── connection.py          ← SQLite-Verbindungsmanagement
│   ├── migrations.py          ← Schema-Migrationen (CREATE TABLE ...)
│   ├── models.py              ← Datenklassen (dataclasses)
│   ├── pax_db.py              ← Passagier-Datenbankoperationen
│   ├── sanmat_db.py           ← SANMAT-Datenbankoperationen
│   ├── sonderaufgaben_db.py   ← Sonderaufgaben-DB
│   ├── turso_sync.py          ← Turso Cloud-Synchronisierung
│   ├── workflow_db.py         ← Workflow-DB
│   └── export_historie_db.py  ← Export-Historie-DB
│
├── functions/                 ← Geschäftslogik & DB-Funktionen
│   ├── mitarbeiter_functions.py
│   ├── fahrzeug_functions.py
│   ├── dienstplan_functions.py
│   ├── uebergabe_functions.py
│   ├── verspaetung_functions.py
│   ├── verspaetung_db.py
│   ├── vorkommnisse_db.py
│   ├── beschwerden_db.py
│   ├── schulungen_db.py
│   ├── telefonnummern_db.py
│   ├── handys_db.py
│   ├── handys_email.py
│   ├── handys_excel_export.py
│   ├── handys_bericht.py
│   ├── mail_functions.py
│   ├── laufzettel_functions.py
│   ├── mitarbeiter_dokumente_functions.py
│   ├── mitarbeiter_sync.py
│   ├── archiv_functions.py
│   ├── dokument_archiv.py
│   ├── dienstplan_html_export.py
│   ├── dienstplan_parser.py
│   ├── staerkemeldung_export.py
│   ├── staerkemeldung_dashboard_export.py
│   ├── stellungnahmen_db.py
│   ├── stellungnahmen_html_export.py
│   ├── emobby_functions.py
│   ├── call_transcription_db.py
│   ├── dienstanweisungen_db.py
│   ├── notizen_db.py
│   ├── psa_db.py
│   └── settings_functions.py
│
├── backup/                    ← Backup-System
│   └── backup_manager.py
│
├── database SQL/              ← SQLite-Datenbankdateien (Laufzeitdaten)
│   ├── nesk3.db               ← Hauptdatenbank
│   ├── mitarbeiter.db         ← Mitarbeiterdatenbank
│   ├── einsaetze.db           ← Einsatzdatenbank
│   ├── verspaetungen.db
│   ├── telefonnummern.db
│   ├── beschwerden.db
│   ├── vorkommnisse.db
│   ├── handys.db
│   ├── workflow.db
│   └── ...
│
├── Daten/                     ← Vorlagen, Logos, Exportdaten (in EXE eingebettet)
│   ├── Logo/                  ← DRK-Logo, Nesk3.ico
│   ├── Vordrucke/             ← Druckvorlagen
│   └── vorfeldmit/            ← Vorfeldmitarbeiter-Exports
│
├── json/                      ← JSON-Konfigurationsdaten (in EXE eingebettet)
│
├── G EXE/                     ← EXE-Build-Ausgabeverzeichnis
├── Backup Neu ab 20.03/       ← App-Code-Backups (ZIP + Ordner)
└── Backup Data/               ← Datenbankbackups
    ├── db_backups/            ← SQLite-DB-Kopien
    └── Dokumente/             ← Dokument-Archiv
```

## Startvoraussetzungen (Entwicklung)

```powershell
# Python 3.13 installiert
python3.13 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# App starten
python3.13 main.py
```

## Design-Prinzipien

- **SAP Fiori Design**: Einheitliches UI-System basierend auf SAP Fiori-Farbpalette
- **Sidebar-Navigation**: Dunkles Theme (`#354a5e`), 36 Seiten, QStackedWidget
- **Tab-Farbe**: Einheitliches `#1565a8` (DRK-Blau) in allen 11 GUI-Tabs
- **Fade-Animation**: 180ms OutCubic bei Seitenwechsel (QGraphicsOpacityEffect)
- **Offline-first**: Volle Funktionsfähigkeit ohne Internet; Turso-Sync im Hintergrund
- **Multi-PC**: OneDrive als Dateiablage; WAL-Modus für gleichzeitigen SQLite-Zugriff

## Farbpalette (SAP Fiori)

| Konstante | Hex | Verwendung |
|---|---|---|
| `FIORI_BLUE` | `#0a6ed1` | Primärfarbe Buttons |
| `FIORI_LIGHT_BLUE` | `#eef4fa` | Hintergründe |
| `FIORI_TEXT` | `#32363a` | Textfarbe |
| `FIORI_BORDER` | `#d9d9d9` | Rahmenlinien |
| `FIORI_SUCCESS` | `#107e3e` | Grün / Erfolg |
| `FIORI_WARNING` | `#e9730c` | Orange / Warnung |
| `FIORI_ERROR` | `#bb0000` | Rot / Fehler |
| `FIORI_SIDEBAR_BG` | `#354a5e` | Sidebar-Hintergrund |
| Tab-Blau | `#1565a8` | Alle Tab-Leisten |
