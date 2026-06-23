# Build & Deployment – EXE erstellen

## Voraussetzungen

- Python 3.13 (Windows)
- PyInstaller installiert (`pip install pyinstaller`)
- Alle Abhängigkeiten installiert (`pip install -r requirements.txt`)
- Arbeitspfad: `C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3`

---

## Build-Befehl

```powershell
python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE" --workpath build_tmp --noconfirm
```

| Parameter | Beschreibung |
|---|---|
| `Nesk3.spec` | Build-Konfigurationsdatei |
| `--distpath "G EXE"` | EXE-Ausgabeverzeichnis |
| `--workpath build_tmp` | Temporäre Build-Dateien (kann danach gelöscht werden) |
| `--noconfirm` | Keine Rückfrage bei Überschreiben |

**Ausgabe**: `G EXE/Nesk3.exe` (Single-File EXE, ~150-250 MB)

---

## `Nesk3.spec` – Build-Konfiguration

```python
# -*- mode: python ; coding: utf-8 -*-
import os

BASE    = r"C:\Users\DRKairport\OneDrive - ...\Nesk3"
OUT_DIR = r"C:\Users\DRKairport\OneDrive - ...\Nesk3\G EXE"

# Daten-Ordner einsammeln (AOCC-Ordner ausgeschlossen)
_SKIP_DIRS = {"AOCC"}
_daten_files = []
_daten_root = os.path.join(BASE, "Daten")
for _dirpath, _dirnames, _filenames in os.walk(_daten_root):
    _dirnames[:] = [d for d in _dirnames if d not in _SKIP_DIRS]
    for _fn in _filenames:
        _src = os.path.join(_dirpath, _fn)
        if not os.path.isfile(_src):   # OneDrive-Placeholder überspringen
            continue
        _rel = os.path.relpath(os.path.dirname(_src), BASE)
        _daten_files.append((_src, _rel))

a = Analysis(
    [os.path.join(BASE, "main.py")],
    pathex=[BASE],
    binaries=[],
    datas=_daten_files + [
        (os.path.join(BASE, "json"), "json"),
    ],
    hiddenimports=[
        "PySide6.QtPrintSupport",
        "PySide6.QtSvgWidgets",
        "win32com", "win32com.client", "win32com.server",
        "pywintypes", "pythoncom",
        "docx",
        "openpyxl", "openpyxl.styles", "openpyxl.utils",
        "pypdf", "pypdf._reader", "pypdf._writer",
        "pypdf.filters", "pypdf.generic",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Nesk3",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,                          # Kein Konsolenfenster
    disable_windowed_traceback=False,
    icon=os.path.join(BASE, "Daten", "Logo", "nesk3.ico"),
)
```

### Wichtige Punkte:

1. **`console=False`**: Kein schwarzes Terminal-Fenster beim Start
2. **`upx=False`**: UPX-Kompression deaktiviert (verhindert Antivirus-Fehlalarme)
3. **`noarchive=False`**: Alle Dateien in der EXE eingebettet (Single-File)
4. **`datas`**: Enthält alle Dateien aus `Daten/` (außer `AOCC/`) und `json/`
5. **`hiddenimports`**: Pakete, die PyInstaller sonst nicht findet

---

## In der EXE enthaltene Daten

- `Daten/Logo/` – DRK-Logo (PNG + ICO)
- `Daten/Vordrucke/` – Druckvorlagen
- `Daten/vorfeldmit/` – Vorfeldmitarbeiter-Templates
- `json/` – JSON-Konfigurationsdateien

**Nicht in der EXE** (werden zur Laufzeit aus OneDrive gelesen):
- `database SQL/*.db` – Datenbankdateien
- `Backup Data/` – Backups
- Alle Benutzerdaten

---

## Pfadauflösung im EXE-Modus

Die EXE kann auf **beliebigen DRK-PCs** starten, die OneDrive synchronisiert haben:

```python
# config.py - _find_base_dir()
if getattr(sys, "frozen", False):    # EXE-Modus
    for var in ("OneDriveCommercial", "OneDrive"):
        od = os.environ.get(var, "")
        if od:
            candidate = os.path.join(od, _ONEDRIVE_SUBPATH)
            if os.path.isdir(candidate):
                return candidate
    # Fallback: Ordner neben der EXE
    return os.path.dirname(sys.executable)
```

**Unterstützte PCs**: W11-262011, W11-262013 und alle weiteren DRK-PCs mit aktivem OneDrive.

---

## Task (VS Code)

```json
{
    "label": "Nesk3 starten",
    "type": "shell",
    "command": "C:/Users/DRKairport/AppData/Local/Microsoft/WindowsApps/python3.13.exe main.py",
    "isBackground": false,
    "group": "build"
}
```

---

## requirements.txt (Abhängigkeiten)

```
PySide6
openpyxl
python-docx
pypdf
pywin32
pyinstaller
```

Vollständige Liste: `requirements.txt` im Projektverzeichnis.

---

## Deployment auf DRK-PCs

1. EXE bauen: `python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE" --workpath build_tmp --noconfirm`
2. `G EXE/Nesk3.exe` per OneDrive synchronisiert automatisch auf alle DRK-PCs
3. Benutzer starten `Nesk3.exe` direkt aus dem OneDrive-Ordner oder via Desktop-Verknüpfung
4. Datenbankpfad wird automatisch über die `OneDriveCommercial`-Umgebungsvariable aufgelöst

---

## Build-Verzeichnisse

| Verzeichnis | Inhalt |
|---|---|
| `G EXE/` | Fertige EXE-Datei(en) |
| `build_tmp/` | Temporäre PyInstaller-Dateien (kann gelöscht werden) |
| `build/` | Ältere Build-Artefakte |
