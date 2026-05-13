# Rekonstruktions-Anleitung (Stand: 13.05.2026)

> Diese Datei erklärt **Schritt für Schritt**, wie Nesk3 aus den Quellen komplett neu aufgebaut werden kann.

---

## Schritt 1: Ordnerstruktur anlegen

```
Nesk3/
├── gui/
├── functions/
├── database/
├── backup/
├── json/
├── Daten/
│   ├── Handys/
│   │   └── Berichte/
│   ├── Spät/
│   │   └── Monatsliste/
│   │       └── Auto Liste/
│   ├── Hilfe/
│   │   └── screenshots/
│   ├── Vorfeldschulung/
│   └── Vordrucke/
├── database SQL/
│   └── Backup Data/
│       └── db_backups/
└── Mds 13.05.2026/
```

---

## Schritt 2: Python-Umgebung

```powershell
# Python 3.13 installieren (microsoft.com/store oder python.org)
python --version  # sollte 3.13.x zeigen

# Abhängigkeiten installieren
pip install PySide6>=6.6.0
pip install openpyxl>=3.1.0
pip install python-docx>=1.0.0
pip install pypdf>=4.0.0
pip install pywin32           # für Outlook-Integration
pip install google-generativeai  # für Gemini-KI
```

---

## Schritt 3: Dateien erstellen (Reihenfolge beachten)

### 3.1 config.py (zuerst!)

Muss als erstes erstellt werden, da alle anderen Module davon abhängen.

Wichtige Variablen (siehe `02_KONFIGURATION.md`):
- `BASE_DIR` → automatisch aus Umgebungsvariable oder `__file__`
- Alle DB-Pfade
- `TURSO_URL`, `TURSO_TOKEN`
- `GEMINI_API_KEY`
- Fiori-Farben

### 3.2 database/ Ordner

Reihenfolge:
1. `database/__init__.py` (leer oder `from .connection import get_connection`)
2. `database/connection.py` → `get_connection()` für nesk3.db
3. `database/migrations.py` → vollständiges SQL-Schema (siehe `04_DATENBANKEN.md`)
4. `database/models.py` → Datenklassen
5. `database/turso_sync.py` → TABLE_MAP + push_row/pull_all (siehe `08_TURSO_SYNC.md`)
6. `database/pax_db.py`, `sanmat_db.py`, `sonderaufgaben_db.py`, `export_historie_db.py`

### 3.3 functions/ Ordner

Jedes Modul initialisiert seine DB selbst via `_init_db()` beim ersten Augriff.

Reihenfolge (Abhängigkeiten beachten):
1. `functions/settings_functions.py` (von anderen genutzt)
2. `functions/notizen_db.py` (Schema in `04_DATENBANKEN.md`)
3. `functions/vorkommnisse_db.py`
4. `functions/beschwerden_db.py`
5. `functions/handys_db.py`
6. alle anderen (unabhängig voneinander)

### 3.4 backup/ Ordner

1. `backup/__init__.py`
2. `backup/backup_manager.py`

### 3.5 gui/ Ordner

Reihenfolge (von unten nach oben):
1. `gui/splash_screen.py`
2. `gui/slot_machine.py` + `gui/slot_symbols.py` (Easter Egg)
3. `gui/dashboard.py`
4. Alle anderen Widget-Dateien
5. `gui/main_window.py` (zuletzt, importiert alle anderen)

### 3.6 main.py (zuletzt)

Orchestriert alles, ruft `run_migrations()`, startet QApplication.

---

## Schritt 4: Datenbank-Initialisierung

```python
# Beim ersten Start automatisch via main.py:
from database.migrations import run_migrations
run_migrations()  # erstellt alle Tabellen in nesk3.db

# Alle anderen DBs erstellen sich selbst:
from functions.notizen_db import _init_db as notizen_init
from functions.vorkommnisse_db import _init_db as vork_init
from functions.handys_db import _init_schema as handys_init
# etc. – passiert automatisch beim ersten GUI-Aufruf
```

---

## Schritt 5: Turso-Cloud vorbereiten (optional)

Falls Turso-Sync genutzt werden soll:

1. Turso-Account anlegen: https://turso.tech
2. Neue Datenbank erstellen: `nesk`
3. Alle Turso-Tabellen anlegen (Namensmuster: `<präfix>__<tabelle>`)
4. JWT-Token generieren
5. In `config.py` eintragen: `TURSO_URL`, `TURSO_TOKEN`

Tabellen-Namensmuster aus `TABLE_MAP` in `database/turso_sync.py` (siehe `08_TURSO_SYNC.md`).

---

## Schritt 6: Testen

```powershell
# App starten:
python main.py

# Zu erwartende Ausgaben:
# [OK] Datenbank bereit.
# [OK] Startup-Backup erstellt
# Splash-Screen erscheint → MainWindow lädt
```

---

## Schritt 7: EXE-Build (optional)

```powershell
# PyInstaller installieren:
pip install pyinstaller

# Build ausführen:
pyinstaller Nesk3.spec

# EXE liegt in: dist/Nesk3/Nesk3.exe
```

---

## Kritische Code-Abschnitte

### BASE_DIR-Ermittlung (config.py)
```python
# EXE-Modus: über OneDrive-Umgebungsvariable
# Script-Modus: __file__
# Wichtig: App läuft von OneDrive-Pfad, NICHT aus EXE-Verzeichnis
```

### PRAGMA-Einstellungen (alle DB-Verbindungen)
```python
conn.execute("PRAGMA journal_mode = WAL")
conn.execute("PRAGMA synchronous  = NORMAL")
conn.execute("PRAGMA busy_timeout  = 5000")
conn.execute("PRAGMA foreign_keys = ON")
```

### Datum-Format Konsistenz
```python
# notizen.datum:      dd.MM.yyyy   (z.B. "13.05.2026")
# SQLite DEFAULT:     datetime('now','localtime')  → "2026-05-13 12:00:00"
# Dienstplan:         YYYY-MM-DD
# Alle anderen:       gemischt – je nach Modul prüfen!
```

---

## Häufige Fehler bei Rekonstruktion

| Fehler | Ursache | Lösung |
|---|---|---|
| `ModuleNotFoundError: config` | sys.path fehlt | `sys.path.insert(0, BASE_DIR)` in main.py |
| `sqlite3.OperationalError: no such table` | DB nicht initialisiert | `run_migrations()` aufrufen |
| `win32com.client` fehlt | pywin32 nicht installiert | `pip install pywin32` |
| E-Mail-Funktionen schlagen fehl | Outlook nicht installiert | win32com benötigt Outlook-Desktop |
| EXE findet BASE_DIR nicht | OneDrive-Env-Var fehlt | Umgebungsvariable `OneDriveCommercial` prüfen |
| WAL/SHM-Dateien von OneDrive | Sync-Konflikt | `_cleanup_onedrive_artefakte()` in main.py |

---

## Dateien-Checkliste

```
✓ config.py
✓ main.py
✓ requirements.txt
✓ Nesk3.spec
✓ database/__init__.py
✓ database/connection.py
✓ database/migrations.py
✓ database/models.py
✓ database/turso_sync.py
✓ database/pax_db.py
✓ database/sanmat_db.py
✓ database/sonderaufgaben_db.py
✓ database/export_historie_db.py
✓ backup/__init__.py
✓ backup/backup_manager.py
✓ functions/notizen_db.py
✓ functions/vorkommnisse_db.py
✓ functions/beschwerden_db.py
✓ functions/handys_db.py
✓ functions/handys_bericht.py
✓ functions/handys_email.py
✓ functions/handys_excel_export.py
✓ functions/verspaetung_db.py
✓ functions/verspaetung_functions.py
✓ functions/mitarbeiter_functions.py
✓ functions/mitarbeiter_dokumente_functions.py
✓ functions/mitarbeiter_sync.py
✓ functions/dienstplan_functions.py
✓ functions/dienstplan_parser.py
✓ functions/dienstplan_html_export.py
✓ functions/fahrzeug_functions.py
✓ functions/schulungen_db.py
✓ functions/uebergabe_functions.py
✓ functions/archiv_functions.py
✓ functions/stellungnahmen_db.py
✓ functions/stellungnahmen_html_export.py
✓ functions/mail_functions.py
✓ functions/settings_functions.py
✓ functions/telefonnummern_db.py
✓ functions/emobby_functions.py
✓ functions/psa_db.py
✓ functions/dokument_archiv.py
✓ functions/dienstanweisungen_db.py
✓ functions/call_transcription_db.py
✓ functions/staerkemeldung_export.py
✓ functions/staerkemeldung_dashboard_export.py
✓ gui/main_window.py
✓ gui/splash_screen.py
✓ gui/dashboard.py
✓ gui/mitarbeiter.py
✓ gui/dienstplan.py
✓ gui/fahrzeuge.py
✓ gui/vorkommnisse.py
✓ gui/beschwerden.py
✓ gui/passagiere.py
✓ gui/passagieranfragen.py
✓ gui/schulungen_kalender.py
✓ gui/handys_widget.py
✓ gui/schadensbericht_dialog.py
✓ gui/uebergabe.py
✓ gui/mitarbeiter_dokumente.py
✓ gui/dokument_browser.py
✓ gui/dienstliches.py
✓ gui/aufgaben.py
✓ gui/aufgaben_tag.py
✓ gui/aufgaben_haupt.py
✓ gui/sonderaufgaben.py
✓ gui/checklisten.py
✓ gui/code19.py
✓ gui/einstellungen.py
✓ gui/backup_widget.py
✓ gui/bericht.py
✓ gui/telefonnummern.py
✓ gui/call_transcription.py
✓ gui/hilfe_dialog.py
✓ gui/slot_machine.py
✓ gui/slot_symbols.py
✓ gui/sanmat/ (Unterordner)
```
