# Konfiguration – config.py (Stand: 13.05.2026)

## Vollständiger Inhalt

```python
# config.py

import os
import sys

_ONEDRIVE_SUBPATH = os.path.join(
    "Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26",
    "Nesk", "Nesk3"
)

def _find_base_dir() -> str:
    if getattr(sys, "frozen", False):
        for var in ("OneDriveCommercial", "OneDrive"):
            od = os.environ.get(var, "")
            if od:
                candidate = os.path.join(od, _ONEDRIVE_SUBPATH)
                if os.path.isdir(candidate):
                    return candidate
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR  = _find_base_dir()

# ── SQLite Datenbank-Pfade ────────────────────────────────────────────────────
_DB_DIR             = os.path.join(BASE_DIR, "database SQL")
DB_PATH             = os.path.join(_DB_DIR, "nesk3.db")         # Haupt-DB
ARCHIV_DB_PATH      = os.path.join(_DB_DIR, "archiv.db")
MITARBEITER_DB_PATH = os.path.join(_DB_DIR, "mitarbeiter.db")
BESCHWERDEN_DB_PATH = os.path.join(_DB_DIR, "beschwerden.db")
SANMAT_DB_PATH      = os.path.join(_DB_DIR, "sanmat.db")
VORKOMMNISSE_DB_PATH= os.path.join(_DB_DIR, "vorkommnisse.db")
NOTIZEN_DB_PATH     = os.path.join(_DB_DIR, "notizen.db")
HANDYS_DB_PATH      = os.path.join(_DB_DIR, "handys.db")
os.makedirs(_DB_DIR, exist_ok=True)

# ── Anwendungseinstellungen ───────────────────────────────────────────────────
APP_NAME    = "Nesk3 – DRK Flughafen Köln"
APP_VERSION = "3.7.0"
APP_LANG    = "de"

# ── Backup ────────────────────────────────────────────────────────────────────
BACKUP_DIR      = "backup/exports"
BACKUP_MAX_KEEP = 30

# ── JSON-Einstellungen ────────────────────────────────────────────────────────
JSON_DIR = os.path.join(BASE_DIR, "json")

# ── Turso Cloud-Sync ──────────────────────────────────────────────────────────
TURSO_URL   = "https://nesk-koboltze.aws-eu-west-1.turso.io"
TURSO_TOKEN = "<JWT-Token – siehe Quellcode>"
TURSO_SYNC_INTERVAL = 30   # Sekunden zwischen auto. Syncs

# ── SAP Fiori Design-Farben ───────────────────────────────────────────────────
FIORI_BLUE        = "#0a6ed1"
FIORI_LIGHT_BLUE  = "#eef4fa"
FIORI_TEXT        = "#32363a"
FIORI_BORDER      = "#d9d9d9"
FIORI_SUCCESS     = "#107e3e"
FIORI_WARNING     = "#e9730c"
FIORI_ERROR       = "#bb0000"
FIORI_WHITE       = "#ffffff"
FIORI_SIDEBAR_BG  = "#354a5e"

# ── KI-Integration ────────────────────────────────────────────────────────────
GEMINI_API_KEY = "<Gemini API Key – siehe Quellcode>"

# ── Handys-Modul ──────────────────────────────────────────────────────────────
HANDYS_EXPORT_PATH = os.path.join(BASE_DIR, "Daten", "Handys")
```

---

## Variablen-Referenz

| Variable | Typ | Wert / Beschreibung |
|---|---|---|
| `BASE_DIR` | str | Wurzelpfad der App (OneDrive-Ordner oder Script-Verzeichnis) |
| `_DB_DIR` | str | `BASE_DIR/database SQL/` |
| `DB_PATH` | str | Pfad zu `nesk3.db` (Haupt-Datenbank) |
| `ARCHIV_DB_PATH` | str | Pfad zu `archiv.db` |
| `MITARBEITER_DB_PATH` | str | Pfad zu `mitarbeiter.db` |
| `BESCHWERDEN_DB_PATH` | str | Pfad zu `beschwerden.db` |
| `SANMAT_DB_PATH` | str | Pfad zu `sanmat.db` |
| `VORKOMMNISSE_DB_PATH` | str | Pfad zu `vorkommnisse.db` |
| `NOTIZEN_DB_PATH` | str | Pfad zu `notizen.db` |
| `HANDYS_DB_PATH` | str | Pfad zu `handys.db` |
| `APP_NAME` | str | `"Nesk3 – DRK Flughafen Köln"` |
| `APP_VERSION` | str | `"3.7.0"` |
| `BACKUP_DIR` | str | `"backup/exports"` |
| `BACKUP_MAX_KEEP` | int | `30` (max. gespeicherte Backups) |
| `JSON_DIR` | str | `BASE_DIR/json/` |
| `TURSO_URL` | str | Turso-Endpoint URL |
| `TURSO_TOKEN` | str | JWT-Token für Turso |
| `TURSO_SYNC_INTERVAL` | int | `30` Sekunden |
| `FIORI_SIDEBAR_BG` | str | `"#354a5e"` – Sidebar-Hintergrund |
| `GEMINI_API_KEY` | str | Google Gemini API Key |
| `HANDYS_EXPORT_PATH` | str | `BASE_DIR/Daten/Handys/` |

---

## main.py – Startup-Ablauf

```
1. sys.path.insert(0, BASE_DIR)
2. sys.excepthook = _excepthook   # alle Fehler sichtbar machen
3. _cleanup_onedrive_artefakte()  # WAL/SHM + Konfliktkopien entfernen
4. _db_startup_backup()           # alle *.db sichern → db_backups/HEUTE/
5. database.migrations.run_migrations()   # Tabellen anlegen / nachrüsten
6. QApplication erstellen
7. SplashScreen anzeigen (animiert)
8. MainWindow erstellen und anzeigen
9. _verspaetung_monats_auto_export()  # 1. des Monats: Vormonats-Excel
10. _taeglich_gemeinsam_backup()      # tägl. Gemeinsam.26-Backup (Thread)
11. Turso-Sync-Thread starten
12. app.exec()
```

---

## Startup-Backup Logik (`_db_startup_backup()`)

- Sichert alle `*.db` im `database SQL/` Ordner
- Ziel: `database SQL/Backup Data/db_backups/YYYY-MM-DD/<name>_HHMMSS.db`
- Pro Tag max. **5 Snapshots** je Datenbank
- Max. **7 Tages-Ordner** gesamt (älteste werden gelöscht)
- Nutzt `sqlite3.Connection.backup()` → konsistente Kopie auch bei WAL

---

## OneDrive-Bereinigung (`_cleanup_onedrive_artefakte()`)

**Problem 1 – WAL/SHM:**
SQLite legt im WAL-Modus `.db-wal` und `.db-shm` Dateien an.  
OneDrive synchronisiert diese und kann bei 568+ Dateien einen "Elemente löschen?"-Dialog erzeugen.  
→ Lösung: `PRAGMA wal_checkpoint(TRUNCATE)` + Dateien löschen.

**Problem 2 – Konfliktkopien:**
Wenn 2 PCs gleichzeitig die DB bearbeiten, erstellt OneDrive Kopien wie `nesk3-W11-8.db`.  
→ Lösung: Regex `^.+-[A-Za-z0-9]+-\d+(-\d+)?\.db$` → löschen.
