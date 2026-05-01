# Nesk3 – Deployment & EXE-Build

**Stand:** 28.04.2026 – v3.11.0

---

## Voraussetzungen

| Komponente | Version | Hinweis |
|-----------|---------|---------|
| Python | 3.13+ | `python3.13` im PATH |
| PySide6 | aktuell | GUI-Framework |
| openpyxl | aktuell | Excel |
| python-docx | aktuell | Word-Export |
| pywin32 | aktuell | Outlook-Integration |
| pypdf | aktuell | PDF-Verarbeitung |
| PyInstaller | aktuell | EXE-Build |

```powershell
pip install PySide6 openpyxl python-docx pywin32 pypdf PyInstaller
```

---

## App starten (Entwicklung)

```powershell
cd "...\Nesk\Nesk3"
python3.13 main.py
```

Oder per VS Code Task: **„Nesk3 starten"**

---

## EXE bauen

```powershell
cd "...\Nesk\Nesk3"
python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE"
```

### Spec-Datei `Nesk3.spec`

- **Einstiegspunkt:** `main.py`
- **Ausgabe:** `G EXE/Nesk3.exe` (Single-File-EXE)
- **Eingebundene Daten:**
  - `Daten/` (rekursiv, ohne `AOCC/`)
  - `json/`
- **Hidden Imports:** `PySide6.QtPrintSupport`, `PySide6.QtSvgWidgets`, `win32com`, `docx`, `openpyxl`, `pypdf`
- **Ausgeschlossen:** `.venv/`, `build/`, `build_tmp/`, Backup-Ordner

### Build-Ausgabe

```
G EXE/
└── Nesk3.exe   ← Fertige Standalone-EXE
```

Die EXE liegt auf OneDrive und ist von allen Rechnern erreichbar.

### Wichtige Hinweise

- Die EXE ermittelt `BASE_DIR` über `%OneDriveCommercial%` / `%OneDrive%` → die Datenbanken liegen immer in `database SQL/` relativ zum OneDrive-Ordner
- Keine Installation notwendig – Doppelklick auf `Nesk3.exe` startet die App
- Beim ersten Start auf einem neuen PC: Sicherstellen dass OneDrive synchronisiert ist und der Freigabe-Ordner existiert

---

## Workflow nach Code-Änderungen

```
1. Code ändern
2. Syntaxcheck:   python3.13 -c "import ast; ast.parse(open('gui/xyz.py').read())"
3. Git commit:    git add . && git commit -m "feat: ..."
4. Git push:      git push                        # → origin (nesk5)
                  git push nesk-neue-word main    # → nesk-neue-word
5. Backup-ZIP:    C:\Daten\Backup Nesk3\Nesk3_backup_YYYYMMDD_HHMMSS.zip
6. MDs updaten:   MD 13.03/ (README, CHANGELOG, FUNKTIONEN, DATENBANKEN)
7. EXE bauen:     python3.13 -m PyInstaller Nesk3.spec --distpath "G EXE"
```

---

## Git-Remotes

| Remote | URL | Zweck |
|--------|-----|-------|
| `origin` / `nesk5` | https://github.com/koboltze-py/nesk5.git | Haupt-Remote |
| `nesk-neue-word` | https://github.com/koboltze-py/nesk-neue-word.git | Spiegel |
| `nesk-word` | https://github.com/koboltze-py/Nesk-Word.git | Älterer Spiegel |
| `nesk4` | https://github.com/koboltze-py/Nesk4.git | Archiv v4 |
| `nesk6` | https://github.com/koboltze-py/Nesk6.git | Zukünftig |

---

## Backup

### Code-Backup (ZIP)
- **Ziel:** `C:\Daten\Backup Nesk3\Nesk3_backup_YYYYMMDD_HHMMSS.zip`
- Enthält: alle Python-Dateien, `json/`, `MD 13.03/`, `Nesk3.spec`, `requirements.txt`
- Ausgeschlossen: `.git/`, `build/`, `build_tmp/`, `Backup*/`, `Daten/`, `G EXE/`, `__pycache__/`

### Datenbank-Backup
- Automatisch via `backup/backup_manager.py` (Button in der App unter „Backup")
- Alternativ manuell: `database SQL/*.db` Dateien kopieren

---

## Neue PCs einrichten

1. Python 3.13 installieren
2. Pakete installieren (siehe oben)
3. OneDrive synchronisieren (Freigabe-Ordner muss vorhanden sein)
4. `Nesk3.exe` aus `G EXE/` starten – fertig

Kein Clone, kein Build notwendig. Die EXE findet die Datenbanken automatisch über OneDrive.
