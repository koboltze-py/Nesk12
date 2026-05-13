# Setup und Start (Stand: 13.05.2026)

## Voraussetzungen

| Anforderung | Details |
|---|---|
| Python | 3.13 (empfohlen: Microsoft Store oder python.org) |
| Betriebssystem | Windows 10/11 (OneDrive-Pfad wird für EXE-Modus genutzt) |
| OneDrive | Muss aktiv und synchronisiert sein |
| Microsoft Outlook | Für E-Mail-Funktionen (win32com) |
| Internet | Für Turso-Cloud-Sync (optional, App läuft auch offline) |

---

## Python-Abhängigkeiten installieren

```powershell
cd "C:\...\Nesk\Nesk3"
pip install -r requirements.txt
```

**requirements.txt:**
```
PySide6>=6.6.0
openpyxl>=3.1.0
python-docx>=1.0.0
pypdf>=4.0.0
# sqlite3 ist in Python standardmäßig enthalten
```

**Zusätzliche Pakete (nicht in requirements.txt, aber benötigt):**
```powershell
pip install pywin32        # Outlook-Integration (win32com.client)
pip install google-generativeai  # Gemini-KI für Transkription
pip install pyinstaller    # nur für EXE-Build
```

---

## App starten (Entwicklungsmodus)

```powershell
cd "C:\Users\DRKairport\OneDrive - ...\Nesk\Nesk3"
python main.py
```

Oder via VS Code Task:
- Task: **"Nesk3 starten"** → führt `python3.13.exe main.py` aus

Oder via PowerShell-Skript:
```powershell
.\start_nesk.ps1
```

---

## Erste-Start-Verhalten

Beim ersten Start passiert automatisch:
1. `config.py` → `BASE_DIR` wird ermittelt (OneDrive-Pfad)
2. `database SQL/` Ordner wird erstellt (falls nicht vorhanden)
3. `database/migrations.py` → Alle SQLite-Tabellen werden angelegt
4. Jede spezialisierte DB (`notizen.db`, `vorkommnisse.db`, etc.) initialisiert sich selbst beim ersten Zugriff
5. Startup-Backup aller `.db`-Dateien in `database SQL/Backup Data/db_backups/YYYY-MM-DD/`
6. OneDrive WAL/SHM-Bereinigung

---

## EXE-Build (PyInstaller)

```powershell
pyinstaller Nesk3.spec
```

Die fertige EXE liegt in `dist/Nesk3/Nesk3.exe`.

**Im EXE-Modus:**
- `BASE_DIR` wird über `OneDriveCommercial` oder `OneDrive` Umgebungsvariable ermittelt
- Alle Daten liegen weiterhin auf OneDrive, nicht in der EXE
- Die EXE kann auf jedem PC genutzt werden, solange OneDrive synchronisiert ist

---

## Pfad-Logik (wichtig für Rekonstruktion)

Die App nutzt **keinen festen Pfad** – stattdessen:

```python
# In config.py:
def _find_base_dir() -> str:
    if getattr(sys, "frozen", False):
        # EXE-Modus
        for var in ("OneDriveCommercial", "OneDrive"):
            od = os.environ.get(var, "")
            if od:
                candidate = os.path.join(od, _ONEDRIVE_SUBPATH)
                if os.path.isdir(candidate):
                    return candidate
    else:
        # Script-Modus
        return os.path.dirname(os.path.abspath(__file__))
```

`_ONEDRIVE_SUBPATH`:
```
Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3
```

---

## Datenbankdateien (müssen vorhanden sein / werden auto-erstellt)

Alle Dateien in `database SQL/`:

| Datei | Erstellt durch |
|---|---|
| `nesk3.db` | `database/migrations.py` → `run_migrations()` |
| `archiv.db` | `functions/archiv_functions.py` |
| `mitarbeiter.db` | `functions/mitarbeiter_functions.py` |
| `beschwerden.db` | `functions/beschwerden_db.py` → `_init_db()` |
| `sanmat.db` | `database/sanmat_db.py` |
| `vorkommnisse.db` | `functions/vorkommnisse_db.py` → `_init_db()` |
| `notizen.db` | `functions/notizen_db.py` → `_init_db()` |
| `handys.db` | `functions/handys_db.py` → `_init_schema()` |

---

## Ordnerstruktur (muss vorhanden sein)

```
Nesk3/
├── database SQL/           ← wird auto-erstellt
│   └── Backup Data/
│       └── db_backups/     ← wird auto-erstellt
├── Daten/
│   ├── Handys/
│   │   └── Berichte/
│   ├── Spät/
│   │   └── Monatsliste/
│   │       └── Auto Liste/
│   └── Vorfeldschulung/
├── json/                   ← Einstellungsdateien
└── Mds 13.05.2026/         ← Diese Dokumentation
```
