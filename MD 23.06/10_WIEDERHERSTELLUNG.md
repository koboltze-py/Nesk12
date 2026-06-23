# Wiederherstellungsanleitung

## Gesamtwiederherstellung aus MD-Dokumentation

Diese Anleitung beschreibt, wie Nesk3 vollständig aus den MD-Dokumenten und dem Git-Repository wiederhergestellt werden kann.

---

## Schritt 1: Repository klonen

```powershell
git clone https://github.com/koboltze-py/Nesk12.git Nesk3
cd Nesk3
```

---

## Schritt 2: Python-Umgebung einrichten

```powershell
python3.13 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**Benötigte Pakete** (falls `requirements.txt` fehlt):
```
PySide6
openpyxl
python-docx
pypdf
pywin32
pyinstaller
```

---

## Schritt 3: Verzeichnisstruktur erstellen

```powershell
# Pflicht-Verzeichnisse anlegen
New-Item -ItemType Directory -Force -Path "database SQL"
New-Item -ItemType Directory -Force -Path "Daten\Logo"
New-Item -ItemType Directory -Force -Path "Daten\Vordrucke"
New-Item -ItemType Directory -Force -Path "json"
New-Item -ItemType Directory -Force -Path "backup\exports"
```

---

## Schritt 4: Datenbank-Tabellen erstellen

```python
# Tabellen werden automatisch beim ersten Start angelegt
# via database/migrations.py → init_database()
python3.13 main.py
```

Alternativ manuell:
```python
from database.migrations import init_database
init_database()
```

---

## Schritt 5: Konfiguration prüfen (`config.py`)

```python
# Sicherstellen dass BASE_DIR korrekt aufgelöst wird:
from config import BASE_DIR, DB_PATH
print(BASE_DIR)   # Muss auf Nesk3-Verzeichnis zeigen
print(DB_PATH)    # Muss auf database SQL/nesk3.db zeigen
```

---

## Schritt 6: Datenbankdaten wiederherstellen

### Aus automatischem DB-Backup:
```powershell
# Letztes Backup-Datum finden
ls "database SQL\db_backups\" | Sort-Object Name -Descending | Select -First 1

# DB wiederherstellen
Copy-Item "database SQL\db_backups\2026-06-24\nesk3.db.120000" "database SQL\nesk3.db"
```

### Aus Turso (Cloud):
```python
from database.turso_sync import pull_all
pull_all()   # Holt alle Daten aus Turso
```

### Aus SQL-Backup:
```powershell
# SQLite-Dump einlesen
sqlite3 "database SQL\nesk3.db" ".read backup.sql"
```

---

## Schritt 7: App starten

```powershell
python3.13 main.py
```

---

## EXE aus Quellcode neu erstellen

```powershell
# Im Nesk3-Verzeichnis:
python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE" --workpath build_tmp --noconfirm
```

EXE wird erstellt als: `G EXE\Nesk3.exe`

---

## Wichtige Konfigurationswerte (für Neuerstellung von config.py)

| Variable | Wert |
|---|---|
| `APP_VERSION` | `3.8.0` |
| `TURSO_URL` | `https://nesk-koboltze.aws-eu-west-1.turso.io` |
| `FIORI_SIDEBAR_BG` | `#354a5e` |
| `FIORI_BLUE` | `#0a6ed1` |
| Tab-Farbe | `#1565a8` |
| `HANDYS_EMAIL_EMPFAENGER` | `erste-hilfe-station-flughafen@drk-koeln.de` |

Vollständige Konfiguration: `01_CONFIG.md`

---

## Modulabhängigkeiten (Import-Reihenfolge)

```
main.py
  ├── config.py                  (keine Abhängigkeiten)
  ├── database/
  │   ├── connection.py          ← config
  │   ├── migrations.py          ← connection
  │   ├── models.py              (keine)
  │   └── turso_sync.py          ← config
  ├── functions/
  │   └── *.py                   ← connection, config
  └── gui/
      ├── main_window.py         ← config, alle GUI-Module
      ├── splash_screen.py       (keine)
      ├── dashboard.py           ← config, fahrzeug_functions
      └── *.py                   ← config, functions/*
```

---

## Häufige Probleme & Lösungen

### Problem: DB-Pfad nicht gefunden (EXE-Modus)

```
Fehler: "database SQL\nesk3.db" nicht gefunden
```

**Lösung**: OneDrive-Umgebungsvariable prüfen:
```powershell
echo $env:OneDriveCommercial
# Muss zeigen auf: C:\Users\...\OneDrive - Deutsches Rotes Kreuz ...
```

### Problem: WAL-Dateien blockieren OneDrive

**Lösung**: App schließen, manuell bereinigen:
```powershell
Remove-Item "database SQL\*.db-wal" -Force
Remove-Item "database SQL\*.db-shm" -Force
```

### Problem: Import-Fehler beim Start

**Lösung**: Alle Abhängigkeiten neu installieren:
```powershell
pip install -r requirements.txt --force-reinstall
```

### Problem: PySide6 nicht gefunden in EXE

**Lösung**: `hiddenimports` in `Nesk3.spec` ergänzen und neu bauen.

### Problem: Excel-Export funktioniert nicht

**Lösung**: `openpyxl` installiert? `pip install openpyxl`

### Problem: Outlook-Entwurf nicht erstellt

**Lösung**: `pywin32` installiert? `pip install pywin32`  
Outlook muss auf dem PC installiert und konfiguriert sein.
