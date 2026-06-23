# main.py – Einstiegspunkt

## Übersicht

`main.py` ist der Einstiegspunkt der Anwendung. Er initialisiert alle Module, bereinigt Artefakte, erstellt Backups und startet das Hauptfenster.

---

## Vollständige Startsequenz

```
1. _cleanup_onedrive_artefakte()     ← WAL/SHM + OneDrive-Konfliktkopien bereinigen
2. init_database()                   ← SQLite-Schema erstellen/migrieren
3. pull_all() (Turso)                ← Neueste Daten aus Cloud holen
4. _erstelle_db_backup()             ← Lokale DB-Kopien anlegen
5. QApplication erstellen            ← PySide6-App initialisieren
6. SplashScreen anzeigen             ← Ladebildschirm
7. MainWindow erstellen & anzeigen   ← Hauptfenster
8. Background-Sync-Thread starten    ← Alle 30s Turso-Pull
9. app.exec()                        ← Haupt-Ereignisschleife
```

---

## Vollständiger Quellcode (Einstiegspunkt)

```python
"""
Nesk3 – DRK Flughafen Köln
Mitarbeiter- und Dienstplanverwaltung
Einstiegspunkt der Anwendung
"""
import sys
import os
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

def _excepthook(exc_type, exc_value, exc_tb):
    print("\n=== UNBEHANDELTER FEHLER ===", file=sys.stderr)
    traceback.print_exception(exc_type, exc_value, exc_tb)
    print("============================\n", file=sys.stderr)

sys.excepthook = _excepthook

import sqlite3
import shutil
import glob
import threading
import time
from datetime import datetime
from PySide6.QtWidgets import QApplication, QToolTip
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPalette, QColor, QIcon
from gui.main_window import MainWindow
from gui.splash_screen import SplashScreen


def _cleanup_onedrive_artefakte():
    """
    Bereinigt OneDrive-Sync-Artefakte:
    1. WAL/SHM-Dateien (PRAGMA wal_checkpoint(TRUNCATE) → löschen)
    2. OneDrive-Konfliktkopien (Pattern: name-PC-Zahl.db)
    """
    try:
        from config import DB_PATH
        import re
        db_dir = os.path.dirname(DB_PATH)

        # WAL-Checkpoint + WAL/SHM löschen
        geloescht_wal = 0
        for wal_path in glob.glob(os.path.join(db_dir, "*.db-wal")):
            db_path = wal_path[:-4]
            if os.path.isfile(db_path):
                try:
                    con = sqlite3.connect(db_path, timeout=3)
                    con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                    con.close()
                except Exception:
                    pass
            try:
                os.remove(wal_path)
                geloescht_wal += 1
            except Exception:
                pass
        for shm_path in glob.glob(os.path.join(db_dir, "*.db-shm")):
            try:
                os.remove(shm_path)
            except Exception:
                pass

        # OneDrive-Konfliktkopien löschen
        konflikt_muster = re.compile(r'^.+-[A-Za-z0-9]+-\d+(-\d+)?\.db$')
        for f in os.listdir(db_dir):
            if konflikt_muster.match(f):
                try:
                    os.remove(os.path.join(db_dir, f))
                except Exception:
                    pass
    except Exception as e:
        print(f"[WARN] Artefakt-Bereinigung fehlgeschlagen: {e}")


def _erstelle_db_backup():
    """
    Erstellt tägliche SQLite-Backups in database SQL/db_backups/YYYY-MM-DD/.
    Max. 5 Kopien pro DB pro Tag, max. 7 Tages-Ordner.
    """
    try:
        from config import BASE_DIR
        db_dir   = os.path.join(BASE_DIR, "database SQL")
        bak_root = os.path.join(db_dir, "db_backups")
        heute    = datetime.now().strftime("%Y-%m-%d")
        bak_dir  = os.path.join(bak_root, heute)
        os.makedirs(bak_dir, exist_ok=True)

        for db_file in glob.glob(os.path.join(db_dir, "*.db")):
            name = os.path.basename(db_file)
            kopien = sorted(glob.glob(os.path.join(bak_dir, name + ".*")))
            if len(kopien) >= 5:
                continue
            ts = datetime.now().strftime("%H%M%S")
            shutil.copy2(db_file, os.path.join(bak_dir, f"{name}.{ts}"))

        # Alte Tages-Ordner bereinigen (max. 7)
        alle_tage = sorted(os.listdir(bak_root))
        while len(alle_tage) > 7:
            shutil.rmtree(os.path.join(bak_root, alle_tage.pop(0)), ignore_errors=True)
    except Exception as e:
        print(f"[WARN] DB-Backup fehlgeschlagen: {e}")


def main():
    _cleanup_onedrive_artefakte()

    from database.migrations import init_database
    init_database()

    try:
        from database.turso_sync import pull_all
        pull_all()
    except Exception as e:
        print(f"[WARN] Turso-Pull fehlgeschlagen: {e}")

    _erstelle_db_backup()

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Standard-Schrift
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Tooltip-Stil
    QToolTip.setFont(QFont("Segoe UI", 9))

    # Splash Screen
    splash = SplashScreen()
    splash.show()
    app.processEvents()

    # Hauptfenster
    window = MainWindow()

    # Splash beenden, Hauptfenster anzeigen
    splash.finish(window)
    window.showMaximized()

    # Hintergrund-Sync
    def _bg_sync():
        from config import TURSO_SYNC_INTERVAL
        from database.turso_sync import pull_all as _pull
        while True:
            time.sleep(TURSO_SYNC_INTERVAL)
            try:
                _pull()
            except Exception:
                pass

    t = threading.Thread(target=_bg_sync, daemon=True)
    t.start()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
```

---

## Fehlerbehandlung

```python
# Globaler Exception-Hook – zeigt unbehandelte Fehler im Terminal
def _excepthook(exc_type, exc_value, exc_tb):
    print("\n=== UNBEHANDELTER FEHLER ===", file=sys.stderr)
    traceback.print_exception(exc_type, exc_value, exc_tb)
    print("============================\n", file=sys.stderr)

sys.excepthook = _excepthook
```

Verhindert lautloses Versagen: Alle unbehandelten Exceptions werden sichtbar.
