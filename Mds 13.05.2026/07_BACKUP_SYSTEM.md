# Backup-System (Stand: 13.05.2026)

## Übersicht

Das Backup-System besteht aus **3 Schichten**:

| Schicht | Wo | Wann | Was |
|---|---|---|---|
| Startup-Backup | `database SQL/Backup Data/db_backups/` | Bei jedem App-Start | Alle `*.db`-Dateien |
| Manuelles Backup | `backup/exports/` | Auf Knopfdruck | JSON-Dump (Tabellen) |
| Gemeinsam-Backup | Lokaler Pfad | Täglich (Hintergrund) | Gesamter Gemeinsam.26-Ordner |

---

## 1. Startup-Backup (`_db_startup_backup()` in main.py)

### Ziel-Struktur
```
database SQL/
└── Backup Data/
    └── db_backups/
        ├── 2026-05-13/
        │   ├── nesk3_120000.db
        │   ├── nesk3_143015.db
        │   ├── handys_120000.db
        │   └── ...
        ├── 2026-05-12/
        └── ... (max. 7 Tages-Ordner)
```

### Logik
```python
def _db_startup_backup():
    db_dir = "database SQL/"
    backup_root = "database SQL/Backup Data/db_backups/"
    
    heute = datetime.today().strftime("%Y-%m-%d")
    tag_pfad = backup_root / heute
    os.makedirs(tag_pfad)
    
    for db_file in glob.glob(db_dir + "/*.db"):
        name = os.path.basename(db_file)  # z.B. "nesk3.db"
        ts = datetime.now().strftime("%H%M%S")  # z.B. "143015"
        ziel = tag_pfad / f"{name.replace('.db','')}_{ts}.db"
        
        # SQLite backup() → konsistente Kopie auch bei WAL-Modus
        with sqlite3.connect(db_file) as src:
            dst = sqlite3.connect(ziel)
            src.backup(dst)
            dst.close()
    
    # Alte Tages-Ordner bereinigen (max. 7)
    tage = sorted(glob.glob(backup_root + "/*"))
    while len(tage) > 7:
        shutil.rmtree(tage.pop(0))
    
    # Alte Snapshots je DB bereinigen (max. 5 pro Tag)
    for db_name in ["nesk3", "handys", ...]:
        snapshots = sorted(glob.glob(tag_pfad + f"/{db_name}_*.db"))
        while len(snapshots) > 5:
            os.remove(snapshots.pop(0))
```

### Grenzen
- Max. **5 Snapshots** pro Datenbank pro Tag
- Max. **7 Tages-Ordner** gesamt
- Nutzt `sqlite3.Connection.backup()` für konsistente Kopie

---

## 2. JSON-Backup (`backup/backup_manager.py`)

### Funktionen
```python
def create_backup(typ: str = "manuell") -> str:
    # Erstellt JSON-Dump aller Tabellen
    # Speichert in: BASE_DIR/backup/exports/backup_TIMESTAMP.json
    # Gibt den Dateipfad zurück

def list_backups() -> list[dict]:
    # Gibt Liste aller .json-Backups zurück
    # Jeder Eintrag: {dateiname, pfad, groesse_kb, erstellt}

def restore_backup(filepath: str) -> int:
    # Stellt Backup wieder her
    # Gibt Anzahl wiederhergestellter Datensätze zurück

def _cleanup_old_backups(backup_dir: str):
    # Löscht älteste .json-Backups wenn > BACKUP_MAX_KEEP (30)
```

### Long-Path-Support
```python
def _lp(p: str) -> str:
    # Fügt \\?\ Präfix für Pfade > 259 Zeichen hinzu (Windows-Limit)
    if sys.platform == 'win32' and len(p) > 259:
        return '\\\\?\\' + p
    return p
```

---

## 3. Tägliches Gemeinsam-Backup (`_taeglich_gemeinsam_backup()` in main.py)

```python
def _taeglich_gemeinsam_backup():
    # Läuft in eigenem Thread (Hintergrund)
    # Prüft: Ist heute schon ein Backup gemacht worden? (drk_daten_backup_log)
    # Wenn nein: Kopiert gesamten Gemeinsam.26-Ordner lokal
    # Loggt Ergebnis in drk_daten_backup_log-Tabelle
```

---

## 4. OneDrive-Bereinigung (`_cleanup_onedrive_artefakte()`)

**Problem 1 – WAL/SHM-Dateien:**
```python
# SQLite im WAL-Modus erzeugt .db-wal und .db-shm
# OneDrive synchronisiert diese mit
# Bei 568+ Änderungen erscheint "Elemente löschen?"-Dialog
# Lösung:
for db_file in glob.glob("database SQL/*.db"):
    with sqlite3.connect(db_file) as conn:
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    for ext in ["-wal", "-shm"]:
        p = db_file + ext
        if os.path.exists(p):
            os.remove(p)
```

**Problem 2 – Konfliktkopien:**
```python
# OneDrive erstellt bei Sync-Konflikten Kopien wie "nesk3-W11-8.db"
# Regex: ^.+-[A-Za-z0-9]+-\d+(-\d+)?\.db$
import re
pattern = re.compile(r'^.+-[A-Za-z0-9]+-\d+(-\d+)?\.db$')
for f in os.listdir("database SQL/"):
    if pattern.match(f):
        os.remove(os.path.join("database SQL/", f))
```

---

## 5. Manueller Backup-Restore (backup_widget.py)

Im GUI (`BackupWidget`):
1. Button "Backup erstellen" → `create_backup()`
2. Backup-Liste anzeigen mit Datum und Größe
3. Button "Wiederherstellen" → `restore_backup(pfad)`
4. Bestätigungsdialog vor Restore

---

## 6. Verspätungs-Auto-Export (`_verspaetung_monats_auto_export()`)

```python
def _verspaetung_monats_auto_export():
    # Wird am 1. jeden Monats beim App-Start ausgeführt
    # Erstellt Excel-Datei für den Vormonat
    # Speichert in: Daten/Spät/Monatsliste/Auto Liste/
    # Nutzt: functions/verspaetung_functions.py
```

---

## 7. ZIP-Backup des gesamten Nesk3-Ordners

```python
# In backup_manager.py (für manuelle vollständige Sicherung)
# Erstellt ZIP-Archiv des gesamten BASE_DIR
# Nutzt _lp() für Long-Path-Support
# Exclude: __pycache__, *.pyc, build/, dist/
```
