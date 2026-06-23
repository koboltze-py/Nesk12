# config.py – Konfigurationsdokumentation

## Übersicht

`config.py` ist die zentrale Konfigurationsdatei für Nesk3. Alle Pfade, Datenbankverbindungen, API-Schlüssel und Design-Konstanten werden hier definiert.

## Wichtige Konstanten

### Basispfad (`BASE_DIR`)

```python
BASE_DIR = _find_base_dir()
```

**Logik:**
- Im **EXE-Modus** (`sys.frozen = True`): Liest `OneDriveCommercial` oder `OneDrive` aus den Windows-Umgebungsvariablen und kombiniert mit dem relativen Unterpfad:
  ```
  <OneDrive>\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3
  ```
- Im **Script-Modus**: `os.path.dirname(os.path.abspath(__file__))` → Verzeichnis von `config.py`

**Wichtig**: Der BASE_DIR zeigt IMMER auf den OneDrive-Ordner des angemeldeten Nutzers – unabhängig vom PC. Damit funktioniert die EXE auf W11-262011, W11-262013 und anderen DRK-PCs ohne Konfigurationsänderung.

### Datenbankpfade

| Variable | Datei | Inhalt |
|---|---|---|
| `DB_PATH` | `database SQL/nesk3.db` | Hauptdatenbank (Dienstplan, Fahrzeuge, Übergabe, Settings) |
| `ARCHIV_DB_PATH` | `database SQL/archiv.db` | Archivdaten |
| `MITARBEITER_DB_PATH` | `database SQL/mitarbeiter.db` | Mitarbeiterstammdaten |
| `BESCHWERDEN_DB_PATH` | `database SQL/beschwerden.db` | Beschwerdemanagement |
| `SANMAT_DB_PATH` | `database SQL/sanmat.db` | SANMAT-Daten |
| `VORKOMMNISSE_DB_PATH` | `database SQL/vorkommnisse.db` | Vorkommnis-Berichte |
| `NOTIZEN_DB_PATH` | `database SQL/notizen.db` | Notizen |
| `HANDYS_DB_PATH` | `database SQL/handys.db` | Handy-Verwaltung |
| `WORKFLOW_DB_PATH` | `database SQL/workflow.db` | Workflow-Daten |

### App-Einstellungen

```python
APP_NAME    = "Nesk3 – DRK Flughafen Köln"
APP_VERSION = "3.8.0"
APP_LANG    = "de"
```

### Backup-Einstellungen

```python
BACKUP_DIR      = "backup/exports"    # relativ zu BASE_DIR
BACKUP_MAX_KEEP = 30                  # Maximale Backups
```

### JSON-Verzeichnis

```python
JSON_DIR = os.path.join(BASE_DIR, "json")
```

### Turso-Konfiguration

```python
TURSO_URL            = "https://nesk-koboltze.aws-eu-west-1.turso.io"
TURSO_TOKEN          = "<JWT-Token>"          # Bearer-Auth für Turso HTTP API
TURSO_SYNC_INTERVAL  = 30                     # Sekunden zwischen Auto-Syncs
```

### SAP Fiori Design-Farben

```python
FIORI_BLUE       = "#0a6ed1"    # Primärblau
FIORI_LIGHT_BLUE = "#eef4fa"    # Hintergrund
FIORI_TEXT       = "#32363a"    # Standard-Textfarbe
FIORI_BORDER     = "#d9d9d9"    # Rahmenlinie
FIORI_SUCCESS    = "#107e3e"    # Grün/Erfolg
FIORI_WARNING    = "#e9730c"    # Orange/Warnung
FIORI_ERROR      = "#bb0000"    # Rot/Fehler
FIORI_WHITE      = "#ffffff"    # Weiß
FIORI_SIDEBAR_BG = "#354a5e"    # Sidebar-Hintergrund
```

### KI-Integration

```python
GEMINI_API_KEY = "<Google Gemini API Key>"
```

Wird in `gui/call_transcription.py` für automatische Anruf-Transkription und Datenextraktion verwendet.

### Handys-Modul

```python
HANDYS_EXPORT_PATH       = os.path.join(BASE_DIR, "Daten", "Handys")
HANDYS_EMAIL_EMPFAENGER  = "erste-hilfe-station-flughafen@drk-koeln.de"
```

## Vollständiger Quellcode (Wiederherstellbar)

```python
"""
Konfigurationsdatei für Nesk3
SQLite Datenbankverbindung und App-Einstellungen
"""
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

_DB_DIR             = os.path.join(BASE_DIR, "database SQL")
DB_PATH             = os.path.join(_DB_DIR, "nesk3.db")
ARCHIV_DB_PATH      = os.path.join(_DB_DIR, "archiv.db")
MITARBEITER_DB_PATH = os.path.join(_DB_DIR, "mitarbeiter.db")
os.makedirs(_DB_DIR, exist_ok=True)

APP_NAME    = "Nesk3 – DRK Flughafen Köln"
APP_VERSION = "3.8.0"
APP_LANG    = "de"

BACKUP_DIR      = "backup/exports"
BACKUP_MAX_KEEP = 30

JSON_DIR = os.path.join(BASE_DIR, "json")

TURSO_URL   = "https://nesk-koboltze.aws-eu-west-1.turso.io"
TURSO_TOKEN = (
    "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9"
    ".eyJhIjoicnciLCJnaWQiOiI5MmYxMzNiMS1jYmVkLTQ1NWEtOGU0MS00ZTUxYjYxMjQ0YTYi"
    "LCJpYXQiOjE3NzM4ODgxNjQsInJpZCI6ImY5YTc0NzA1LTE4ZjktNGE2Ny1iNzkyLTM4Yzg4"
    "MTY4N2E3NSJ9"
    ".JSGexxBNRkcbdlAVPGAr8-P0mIiiDuaMWg4elSKf853-xGI5CzcZBxH-ozRLbVjTeM5EhZ6h"
    "N0_OcOvqdVl0Cg"
)
TURSO_SYNC_INTERVAL = 30

FIORI_BLUE       = "#0a6ed1"
FIORI_LIGHT_BLUE = "#eef4fa"
FIORI_TEXT       = "#32363a"
FIORI_BORDER     = "#d9d9d9"
FIORI_SUCCESS    = "#107e3e"
FIORI_WARNING    = "#e9730c"
FIORI_ERROR      = "#bb0000"
FIORI_WHITE      = "#ffffff"
FIORI_SIDEBAR_BG = "#354a5e"

GEMINI_API_KEY = "AIzaSyAoO7bSaxupDJszFv3oS3POA4b0AGMatRQ"

BESCHWERDEN_DB_PATH   = os.path.join(_DB_DIR, "beschwerden.db")
SANMAT_DB_PATH        = os.path.join(_DB_DIR, "sanmat.db")
VORKOMMNISSE_DB_PATH  = os.path.join(_DB_DIR, "vorkommnisse.db")
NOTIZEN_DB_PATH       = os.path.join(_DB_DIR, "notizen.db")
HANDYS_DB_PATH        = os.path.join(_DB_DIR, "handys.db")
WORKFLOW_DB_PATH      = os.path.join(_DB_DIR, "workflow.db")

HANDYS_EXPORT_PATH       = os.path.join(BASE_DIR, "Daten", "Handys")
HANDYS_EMAIL_EMPFAENGER  = "erste-hilfe-station-flughafen@drk-koeln.de"
```
