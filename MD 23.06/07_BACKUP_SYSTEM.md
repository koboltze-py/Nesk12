# Backup-System

## Übersicht

Nesk3 hat drei Backup-Ebenen:

| Ebene | Typ | Speicherort | Automatisch |
|---|---|---|---|
| 1 | SQLite DB-Kopien (täglich) | `database SQL/db_backups/YYYY-MM-DD/` | ✅ App-Start |
| 2 | ZIP-Archiv (Code + Daten) | `Backup Neu ab 20.03/` | ❌ Manuell |
| 3 | JSON-Export (Tabellen) | `backup/exports/` | ❌ Manuell via UI |

---

## Ebene 1: Automatische DB-Backups (App-Start)

```python
# main.py - _erstelle_db_backup()
# Beim jedem App-Start:
#
# database SQL/
# └── db_backups/
#     ├── 2026-06-23/
#     │   ├── nesk3.db
#     │   ├── mitarbeiter.db
#     │   └── ...
#     └── 2026-06-24/ (aktuell)
#         ├── nesk3.db
#         └── ...
```

**Limits:**
- Max. **5 Kopien** pro Datenbank pro Tag
- Max. **7 Tages-Ordner** (älteste werden automatisch gelöscht)

### Restore (manuell)
1. App schließen
2. Gewünschte `.db`-Datei aus `db_backups/YYYY-MM-DD/` kopieren
3. In `database SQL/` einfügen (bestehende Datei überschreiben)
4. App starten

---

## Ebene 2: ZIP-Archiv (Code + Konfiguration)

### Speicherort
```
Backup Neu ab 20.03/
├── bak_20260608_215844/     ← Entpackter Code-Snapshot
├── Nesk3_backup_YYYYMMDD_HHMMSS.zip  ← ZIP-Archiv
└── ...
```

### ZIP erstellen (manuell)
```powershell
# PowerShell: Komplettes Nesk3-Verzeichnis als ZIP sichern
$datum = Get-Date -Format "yyyyMMdd_HHmmss"
$quelle = "C:\Users\DRKairport\OneDrive - ...\Nesk\Nesk3"
$ziel = "$quelle\Backup Neu ab 20.03\Nesk3_backup_$datum.zip"
Compress-Archive -Path $quelle -DestinationPath $ziel -CompressionLevel Optimal
```

### Restore aus ZIP
1. ZIP-Datei entpacken
2. Gewünschte Dateien kopieren (`.py`, `config.py`, etc.)
3. App neu starten

---

## Ebene 3: JSON-Export (`backup/backup_manager.py`)

### Funktion: `create_backup(typ="manuell")`
- Exportiert alle Datenbank-Tabellen als JSON
- Dateiname: `nesk3_backup_YYYYMMDD_HHMMSS.json`
- Speicherort: `backup/exports/`
- Max. `BACKUP_MAX_KEEP = 30` Dateien

### Funktion: `list_backups()`
```python
# Gibt zurück:
[
    {
        "dateiname": "nesk3_backup_20260624_120000.json",
        "pfad": "/full/path/to/file.json",
        "groesse_kb": 512.3,
        "erstellt": "24.06.2026 12:00"
    },
    ...
]
```

---

## `backup/backup_manager.py` – Long-Path-Support

Da der OneDrive-Pfad über 260 Zeichen lang ist, hat der Backup-Manager spezielle Hilfsfunktionen:

```python
def _lp(p: str) -> str:
    """Fügt Windows Long-Path-Präfix (\\\\?\\) hinzu wenn Pfad > 259 Zeichen."""
    if sys.platform == 'win32' and len(p) > 259 and not p.startswith('\\\\?\\'):
        return '\\\\?\\' + p
    return p

def _rmtree_lp(path: str):
    """shutil.rmtree mit Long-Path-Support für OneDrive-Pfade."""
    ...

def _makedirs_lp(path: str):
    """os.makedirs mit Long-Path-Support für OneDrive-Pfade."""
    ...
```

---

## Backup-Übersicht der vorhandenen Backups

### Code-Snapshots (Backup Neu ab 20.03/)
| Datum | Art | Inhalt |
|---|---|---|
| 20260305 | Ordner + ZIP | Vollständiger App-Snapshot mit .venv, dist |
| 20260308 | Ordner | Code-Snapshot |
| 20260311 | Ordner + ZIP | Code-Snapshot |
| 20260313 | Ordner + ZIP | Code-Snapshot |
| 20260314 | Ordner | Code-Snapshot |
| 20260319 | Ordner + ZIP | Code-Snapshot (letzter vor 20.03-Umbau) |
| 20260321 | 2× Ordner | Vollständiger Code-Snapshot |
| 20260501 | Ordner | Code-Snapshot |
| 20260602 | Ordner | Code-Snapshot |
| 20260608 | Ordner | Code-Snapshot (aktuellster) |

### SQL/Datenbank-Backups
| Speicherort | Inhalt |
|---|---|
| `Backup Data/database SQL 08.03/` | SQL-Export 08.03.2026 |
| `Backup Data/db_backups/` | SQLite-DB-Dateien inkl. `pre_*`-Vor-Migrationssicherungen |
| `Backup Data/sql_backups/2026-03-19_213656/` | SQL-Export 19.03.2026 |
| `Database SQL Backup/` | 10 SQL-Backups 09./10.03.2026 |
| `database SQL/` | Aktive DB-Exporte (W11-262011, W11-262013) |

---

## Empfohlener Backup-Workflow

```
Änderungen machen
  ↓
git add . && git commit -m "Beschreibung"
  ↓
git push origin main (→ https://github.com/koboltze-py/Nesk12.git)
  ↓
ZIP-Backup erstellen (Backup Neu ab 20.03/)
  ↓
EXE bauen: python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE" ...
  ↓
MDs aktualisieren (MD 23.06/)
```
