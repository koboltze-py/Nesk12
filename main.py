"""
Nesk3 – DRK Flughafen Köln
Mitarbeiter- und Dienstplanverwaltung
Einstiegspunkt der Anwendung
"""
import sys
import os
import traceback

# Projektverzeichnis in den Python-Pfad aufnehmen
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# Alle unbehandelten Exceptions im Terminal sichtbar machen
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
    Bereinigt zwei OneDrive-Sync-Artefakte im database SQL-Ordner
    die sonst zu OneDrive-Lösch-Dialogen ("568 Elemente löschen?") führen:

    1. WAL / SHM – Hilfsdateien die SQLite beim WAL-Modus anlegt.
       Werden per PRAGMA wal_checkpoint(TRUNCATE) sauber geleert und danach gelöscht.
       Solange Nesk nicht läuft, enthalten sie keine ungespeicherten Daten.

    2. OneDrive-Konfliktkopien – Entstehen wenn zwei PCs gleichzeitig die gleiche
       .db-Datei bearbeitet haben. OneDrive legt dann Kopien nach dem Muster
       'nesk3-W11-<Nummer>.db' an. Diese werden automatisch entfernt, da die
       echten Daten bereits in nesk3.db und Turso gesichert sind.
    """
    try:
        from config import DB_PATH
        import re
        db_dir = os.path.dirname(DB_PATH)

        # ── 1. WAL-Checkpoint + WAL/SHM löschen ─────────────────────────────
        geloescht_wal = 0
        for wal_path in glob.glob(os.path.join(db_dir, "*.db-wal")):
            db_path = wal_path[:-4]          # *.db-wal → *.db
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
        if geloescht_wal:
            print(f"[OK] WAL-Bereinigung: {geloescht_wal} Datei(en) entfernt.")

        # ── 2. OneDrive-Konfliktkopien löschen ──────────────────────────────
        # Muster: <name>-<COMPUTERNAME>-<Zahl>.db  (z. B. nesk3-W11-262013-8.db)
        konflikt_muster = re.compile(r'^.+-[A-Za-z0-9]+-\d+(-\d+)?\.db$')
        geloescht_kf = 0
        for fname in os.listdir(db_dir):
            if konflikt_muster.match(fname):
                try:
                    os.remove(os.path.join(db_dir, fname))
                    geloescht_kf += 1
                except Exception:
                    pass
        if geloescht_kf:
            print(f"[OK] OneDrive-Konfliktkopien entfernt: {geloescht_kf} Datei(en).")

    except Exception as e:
        print(f"[WARNUNG] OneDrive-Bereinigung fehlgeschlagen: {e}")


def _db_startup_backup():
    """Erstellt beim Programmstart ein SQLite-Backup aller Datenbanken.
    Struktur: db_backups/YYYY-MM-DD/<name>_HHMMSS.db
    Pro Tag max. 5 Backups je Datenbank, max. 7 Tages-Ordner insgesamt."""
    try:
        from config import DB_PATH
        from datetime import date as _date, timedelta as _td
        db_dir = os.path.dirname(DB_PATH)
        base_backup = os.path.join(db_dir, "Backup Data", "db_backups")
        jetzt = datetime.now()
        tag_ordner = os.path.join(base_backup, jetzt.strftime("%Y-%m-%d"))
        os.makedirs(tag_ordner, exist_ok=True)
        zeitstempel = jetzt.strftime("%H%M%S")

        # Alle .db-Dateien direkt im database SQL-Ordner sichern (keine Unterordner)
        for db_path in glob.glob(os.path.join(db_dir, "*.db")):
            name = os.path.splitext(os.path.basename(db_path))[0]
            backup_path = os.path.join(tag_ordner, f"{name}_{zeitstempel}.db")
            src = sqlite3.connect(db_path)
            dst = sqlite3.connect(backup_path)
            src.backup(dst)
            dst.close()
            src.close()
            # Pro Tag nur die letzten 5 Backups je Datenbank behalten
            tages_backups = sorted(glob.glob(os.path.join(tag_ordner, f"{name}_*.db")))
            for alt in tages_backups[:-5]:
                try:
                    os.remove(alt)
                except Exception:
                    pass
            print(f"[OK] DB-Backup: {jetzt.strftime('%Y-%m-%d')}/{os.path.basename(backup_path)}")

        # Nur die letzten 7 Tages-Ordner behalten
        alle_tage = sorted([
            d for d in os.listdir(base_backup)
            if os.path.isdir(os.path.join(base_backup, d)) and len(d) == 10 and d.count("-") == 2
        ])
        for alter_tag in alle_tage[:-7]:
            try:
                shutil.rmtree(os.path.join(base_backup, alter_tag))
                print(f"[DEL] Alter Tages-Ordner gelöscht: {alter_tag}")
            except Exception:
                pass
    except Exception as e:
        print(f"[WARNUNG] DB-Backup fehlgeschlagen: {e}")


def _verspaetung_monats_auto_export():
    """
    Wird beim Start geprüft: Ist heute der 1. eines Monats und wurde der
    automatische Vormonats-Export noch nicht erstellt, wird er jetzt erstellt.
    Speicherpfad: Daten/Spät/Monatsliste/Auto Liste/
    """
    try:
        heute = datetime.now()
        if heute.day != 1:
            return  # nur am 1. des Monats

        # Vormonat berechnen
        import calendar
        if heute.month == 1:
            vm_jahr, vm_monat = heute.year - 1, 12
        else:
            vm_jahr, vm_monat = heute.year, heute.month - 1
        letzter_tag = calendar.monthrange(vm_jahr, vm_monat)[1]

        vormonat_name = f"Verspaetungen_{vm_jahr}-{vm_monat:02d}_Auto.xlsx"
        auto_dir = (
            "C:\\Users\\DRKairport\\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\\"
            "Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\\"
            "Nesk\\Nesk3\\Daten\\Spät\\Monatsliste\\Auto Liste"
        )
        ziel_pfad = os.path.join(auto_dir, vormonat_name)

        if os.path.isfile(ziel_pfad):
            print(f"[INFO] Vormonats-Excel bereits vorhanden: {vormonat_name}")
            return

        from datetime import datetime as _dt
        datum_von = _dt(vm_jahr, vm_monat, 1)
        datum_bis = _dt(vm_jahr, vm_monat, letzter_tag, 23, 59, 59)

        from functions.verspaetung_db import lade_verspaetungen
        from gui.mitarbeiter_dokumente import _verspaetungen_als_excel_speichern

        eintraege = lade_verspaetungen(monat=vm_monat, jahr=vm_jahr)
        eintraege.sort(
            key=lambda e: (
                lambda d: _dt.strptime(d, "%d.%m.%Y") if d else _dt.min
            )(e.get("datum", "")),
            reverse=True,
        )

        _verspaetungen_als_excel_speichern(eintraege, ziel_pfad)
        print(f"[OK] Automatischer Vormonats-Export: {vormonat_name} ({len(eintraege)} Einträge)")
    except Exception as e:
        print(f"[WARNUNG] Automatischer Excel-Export fehlgeschlagen: {e}")


def _taeglich_gemeinsam_backup():
    """
    Wird einmal täglich im Hintergrund ausgeführt (nach Fenster-Start).
    Ablauf:
      1. Prüft ob heute bereits ein Gemeinsam.26-Backup existiert → ggf. abbrechen
         (prüft sowohl OneDrive-Ziel als auch lokale Kopie C:\\Daten\\Backup Gemeinsam\\)
      2. Nesk3-Code-Backup (ZIP des App-Ordners)
      3. Gemeinsam.26-Backup (ZIP; altes vorheriges wird automatisch gelöscht)
    Alle Fehler werden als Warnung geloggt, nie als Exception weitergegeben.
    """
    try:
        from backup.backup_manager import (
            create_zip_backup, create_gemeinsam_backup,
            _GEMEINSAM_BACKUP_DIR, _GEMEINSAM_BACKUP_LOKAL,
        )
        heute = datetime.now().strftime("%Y%m%d")
        prefix = f"gemeinsam_{heute}"

        # Prüfen ob heute bereits ein Backup existiert (OneDrive oder lokal)
        for backup_dir in (_GEMEINSAM_BACKUP_DIR, _GEMEINSAM_BACKUP_LOKAL):
            if not os.path.isdir(backup_dir):
                continue
            for fname in os.listdir(backup_dir):
                if fname.startswith(prefix) and fname.endswith('.zip'):
                    print(f"[INFO] Tägliches Gemeinsam-Backup bereits vorhanden: {fname}")
                    return

        print("[INFO] Tägliches Backup wird gestartet …")

        # Schritt 1: Nesk3 Code-Backup erstellen
        print("[INFO] 1/2 – Nesk3 Code-Backup …")
        try:
            zip_path = create_zip_backup()
            print(f"[OK]   Nesk3 Code-Backup: {os.path.basename(zip_path)}")
        except Exception as e:
            print(f"[WARNUNG] Nesk3 Code-Backup fehlgeschlagen: {e}")

        # Schritt 2: Gemeinsam.26-Backup (altes vorheriges wird automatisch gelöscht)
        print("[INFO] 2/2 – Gemeinsam.26-Backup …")
        result = create_gemeinsam_backup(inkrementell=False)
        if result.get('erfolg'):
            erste_zeile = result.get('meldung', '').splitlines()[0]
            print(f"[OK]   Gemeinsam-Backup: {result.get('dateien_count')} Dateien – {erste_zeile}")
            fehler_liste = result.get('fehler_liste', [])
            if fehler_liste:
                print(f"[WARNUNG] Gemeinsam-Backup: {len(fehler_liste)} Datei(en) nicht gesichert:")
                for f in fehler_liste[:10]:
                    print(f"          • {f}")
                if len(fehler_liste) > 10:
                    print(f"          ... und {len(fehler_liste) - 10} weitere")
        else:
            print(f"[WARNUNG] Gemeinsam-Backup fehlgeschlagen: {result.get('meldung')}")

    except Exception as e:
        print(f"[WARNUNG] Tägliches Gemeinsam-Backup fehlgeschlagen: {e}")


def main():
    # High-DPI Unterstützung für Qt6 (PySide6)
    # QT_AUTO_SCREEN_SCALE_FACTOR ist nur Qt5 – in Qt6 ist High-DPI standardmäßig aktiv.
    # PassThrough gibt den echten Skalierungsfaktor (z.B. 1.25 bei 125%) direkt weiter
    # statt ihn auf ganze Zahlen zu runden → schärfere Darstellung auf allen Displays.
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("Nesk3")
    app.setOrganizationName("DRK Flughafen Köln")
    app.setStyle("Fusion")

    # -----------------------------------------------------------------------
    # FONT – plattformunabhängig, identisch auf allen PCs / EXE-Installationen
    # Schrift-Priorität: Segoe UI (Windows) → Arial → Helvetica Neue → OS-default
    # -----------------------------------------------------------------------
    app_font = QFont()
    app_font.setFamilies(["Segoe UI", "Arial", "Helvetica Neue"])
    app_font.setPointSize(10)
    app_font.setStyleHint(QFont.StyleHint.SansSerif)
    app.setFont(app_font)

    # -----------------------------------------------------------------------
    # PALETTE – Textfarbe explizit SCHWARZ; kein Einfluss von Dark Mode /
    # Windows-Theme / High-Contrast-Modus möglich.
    # -----------------------------------------------------------------------
    BLACK    = QColor("#000000")
    WHITE    = QColor("#FFFFFF")
    BG       = QColor("#F5F5F5")   # Fensterhintergrund
    BASE     = QColor("#FFFFFF")   # Eingabefelder, Listen
    ALT      = QColor("#EAEAEA")   # Alternating rows
    BTN      = QColor("#E0E0E0")   # Schaltflächen
    DISABLED = QColor("#A0A0A0")   # Inaktive Elemente
    HILIGHT  = QColor("#0078D4")   # Auswahl / Fokus (DRK-Blau)
    HI_TEXT  = QColor("#FFFFFF")   # Text auf Auswahl

    pal = QPalette()
    # Aktiv & Inaktiv – gleiche Farben (kein "graues Fenster im Hintergrund")
    for grp in (QPalette.ColorGroup.Active,
                QPalette.ColorGroup.Inactive,
                QPalette.ColorGroup.Normal):
        pal.setColor(grp, QPalette.ColorRole.Window,          BG)
        pal.setColor(grp, QPalette.ColorRole.WindowText,      BLACK)
        pal.setColor(grp, QPalette.ColorRole.Base,            BASE)
        pal.setColor(grp, QPalette.ColorRole.AlternateBase,   ALT)
        pal.setColor(grp, QPalette.ColorRole.Text,            BLACK)
        pal.setColor(grp, QPalette.ColorRole.BrightText,      BLACK)
        pal.setColor(grp, QPalette.ColorRole.ButtonText,      BLACK)
        pal.setColor(grp, QPalette.ColorRole.Button,          BTN)
        pal.setColor(grp, QPalette.ColorRole.Highlight,       HILIGHT)
        pal.setColor(grp, QPalette.ColorRole.HighlightedText, HI_TEXT)
        pal.setColor(grp, QPalette.ColorRole.ToolTipBase,     QColor("#FFFFC4"))
        pal.setColor(grp, QPalette.ColorRole.ToolTipText,     BLACK)
        pal.setColor(grp, QPalette.ColorRole.PlaceholderText, QColor("#808080"))
        pal.setColor(grp, QPalette.ColorRole.Link,            QColor("#0078D4"))
        pal.setColor(grp, QPalette.ColorRole.LinkVisited,     QColor("#5C2D91"))
    # Disabled
    pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, DISABLED)
    pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text,       DISABLED)
    pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, DISABLED)
    pal.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base,       QColor("#F0F0F0"))
    app.setPalette(pal)

    # -----------------------------------------------------------------------
    # GLOBAL STYLESHEET
    # -----------------------------------------------------------------------
    app.setStyleSheet("""
        QWidget {
            font-family: 'Segoe UI', Arial, 'Helvetica Neue', sans-serif;
            font-size: 10pt;
            color: #000000;
        }
        QToolTip {
            color: #1b3a5c;
            background-color: #fffde7;
            border: 1px solid #90a4ae;
            border-radius: 4px;
            font-size: 9.5pt;
            padding: 5px 10px;
            opacity: 255;
        }
    """)

    # Tooltip-Palette explizit setzen (belt-and-suspenders gegen Qt-Windows-Bug)
    _tt_pal = QPalette()
    _tt_pal.setColor(QPalette.ColorRole.ToolTipBase, QColor("#fffde7"))
    _tt_pal.setColor(QPalette.ColorRole.ToolTipText, QColor("#1b3a5c"))
    QToolTip.setPalette(_tt_pal)

    # App-Icon setzen
    _icon_path = os.path.join(BASE_DIR, "Daten", "Logo", "nesk3.ico")
    if os.path.exists(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))

    # ── Splash Screen ──────────────────────────────────────────────────────
    from config import APP_VERSION
    splash = SplashScreen(version=APP_VERSION)
    splash.show()
    QApplication.processEvents()

    _init_done = threading.Event()

    def _do_init():
        # OneDrive-Artefakte bereinigen (WAL-Dateien + Konfliktkopien)
        splash.set_status("OneDrive-Bereinigung …")
        _cleanup_onedrive_artefakte()

        # DB-Backup
        splash.set_status("Datenbank-Backup wird erstellt …")
        _db_startup_backup()

        # Migrationen
        splash.set_status("Datenbanktabellen werden geprüft …")
        try:
            from database.migrations import run_migrations
            run_migrations()
        except Exception as e:
            print(f"[WARNUNG] Datenbankinitialisierung fehlgeschlagen: {e}")
            print("[INFO] Bitte Datenbankverbindung in config.py konfigurieren.")

        # Mitarbeiter-DB
        try:
            from database.connection import init_mitarbeiter_db
            init_mitarbeiter_db()
        except Exception as e:
            print(f"[WARNUNG] Mitarbeiter-DB Initialisierung fehlgeschlagen: {e}")

        # Handys-DB
        try:
            from functions.handys_db import init_handys_db
            init_handys_db()
        except Exception as e:
            print(f"[WARNUNG] Handys-DB Initialisierung fehlgeschlagen: {e}")

        # Turso-Sync
        try:
            from database.turso_sync import (
                ensure_turso_schema, pull_all, start_background_sync,
                push_all_local_to_turso, init_sync_ts,
            )
            from backup.backup_manager import is_restore_pending, clear_restore_pending
            splash.set_status("Verbindung wird geprüft …")
            ensure_turso_schema()
            if is_restore_pending():
                splash.set_status("Daten werden übertragen …")
                push_all_local_to_turso()
                clear_restore_pending()
            else:
                splash.set_status("Verbindung wird hergestellt …")
                pull_all()
            init_sync_ts()
            start_background_sync()
        except Exception as e:
            print(f"[WARNUNG] Turso-Sync konnte nicht gestartet werden: {e}")
            print("[INFO] App läuft weiter mit lokalen Datenbanken.")

        _init_done.set()

    threading.Thread(target=_do_init, daemon=True).start()

    # Hauptthread: Splash direkt neu zeichnen → Ringe drehen garantiert
    while not _init_done.is_set():
        splash.repaint()
        QApplication.processEvents()
        time.sleep(0.016)   # 60 fps

    # ── Hauptfenster (muss im Hauptthread erstellt werden) ─────────────────
    splash.set_status("Oberfläche wird geladen …")
    QApplication.processEvents()
    window = MainWindow()
    splash.finish(window)
    window.show()

    # Tägliches Gemeinsam-Backup im Hintergrund (nach Fenster-Start, blockiert UI nicht)
    threading.Thread(target=_taeglich_gemeinsam_backup, daemon=True).start()

    # Automatischer Vormonats-Excel-Export am 1. des Monats
    threading.Thread(target=_verspaetung_monats_auto_export, daemon=True).start()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
