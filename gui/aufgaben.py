"""
Aufgaben-Widget
Kombiniert Checklisten (Drucksachen), Sonderaufgaben und AOCC Lagebericht.
"""
import os
import sys
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QFrame, QMessageBox, QComboBox,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont
from datetime import datetime as _dt

from gui.checklisten    import ChecklistenWidget
from gui.sonderaufgaben import SonderaufgabenWidget
from gui.aufgaben_tag   import _Code19MailTab
from config import FIORI_BLUE, FIORI_TEXT
from functions.settings_functions import get_setting

_AOCC_DEFAULT = (
    r'C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - '
    r'Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - '
    r'DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3\Daten\AOCC\AOCC Lagebericht.xlsm'
)


class _AoccWidget(QWidget):
    """Einfacher Tab zum Öffnen des AOCC Lageberichts."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(16)

        info = QLabel("📣 AOCC Lagebericht")
        info.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        info.setStyleSheet(f"color: {FIORI_TEXT};")
        info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info)

        self._pfad_lbl = QLabel()
        self._pfad_lbl.setStyleSheet("color: #666; font-size: 11px;")
        self._pfad_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._pfad_lbl.setWordWrap(True)
        layout.addWidget(self._pfad_lbl)

        self._status_lbl = QLabel()
        self._status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status_lbl.setStyleSheet("font-size: 11px;")
        layout.addWidget(self._status_lbl)

        btn = QPushButton("📄  AOCC Lagebericht öffnen")
        btn.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        btn.setFixedHeight(50)
        btn.setFixedWidth(340)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {FIORI_BLUE};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 24px;
            }}
            QPushButton:hover {{ background-color: #0855a9; }}
        """)
        btn.clicked.connect(self._open)
        layout.addWidget(btn, 0, Qt.AlignmentFlag.AlignCenter)

        self._refresh()

    def _refresh(self):
        pfad = get_setting('aocc_datei', _AOCC_DEFAULT)
        self._pfad_lbl.setText(pfad)
        if os.path.isfile(pfad):
            self._status_lbl.setText("✅ Datei gefunden")
            self._status_lbl.setStyleSheet("color: #107e3e; font-size: 11px;")
        else:
            self._status_lbl.setText("⚠️ Datei nicht gefunden – Pfad in Einstellungen prüfen")
            self._status_lbl.setStyleSheet("color: #bb6600; font-size: 11px;")

    def _open(self):
        from PySide6.QtWidgets import QMessageBox
        pfad = get_setting('aocc_datei', _AOCC_DEFAULT)
        if not os.path.isfile(pfad):
            QMessageBox.warning(self, "Datei nicht gefunden",
                f"Die Datei wurde nicht gefunden:\n{pfad}\n\n"
                "Bitte den Pfad in den Einstellungen anpassen.")
            return
        try:
            os.startfile(pfad)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Öffnen:\n{exc}")


# ── Sonderaufgaben Historie ────────────────────────────────────────────────

class _SonderaufgabenHistorieWidget(QWidget):
    """Tab: Archiv aller in der Datenbank gespeicherten Sonderaufgaben-Snapshots."""

    # Referenz auf den SonderaufgabenWidget (wird nach der Erstellung gesetzt)
    _sonder_widget = None

    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(10)

        # ── Kopfzeile ──────────────────────────────────────────────────────
        header = QHBoxLayout()
        title = QLabel("🗂️ Sonderaufgaben Historie")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {FIORI_TEXT};")
        header.addWidget(title)
        header.addStretch()

        btn_refresh = QPushButton("🔄 Aktualisieren")
        btn_refresh.setFixedHeight(30)
        btn_refresh.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_refresh.setStyleSheet("""
            QPushButton {
                background: #e8ecf0;
                border: 1px solid #c0cad4;
                border-radius: 4px;
                padding: 4px 14px;
                font-size: 12px;
                color: #333;
            }
            QPushButton:hover { background: #dde4ec; }
        """)
        btn_refresh.clicked.connect(self._refresh_filter)
        header.addWidget(btn_refresh)

        self._btn_laden = QPushButton("↩️ In Formular laden")
        self._btn_laden.setFixedHeight(30)
        self._btn_laden.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_laden.setStyleSheet(f"""
            QPushButton {{
                background: {FIORI_TEXT};
                border: none;
                border-radius: 4px;
                padding: 4px 16px;
                font-size: 12px;
                color: white;
            }}
            QPushButton:hover {{ background: #555; }}
            QPushButton:disabled {{ background: #b0bec5; }}
        """)
        self._btn_laden.setEnabled(False)
        self._btn_laden.clicked.connect(self._laden_selected)
        header.addWidget(self._btn_laden)

        self._btn_excel = QPushButton("📄 Excel öffnen")
        self._btn_excel.setFixedHeight(30)
        self._btn_excel.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_excel.setStyleSheet(f"""
            QPushButton {{
                background: #17a2b8;
                border: none;
                border-radius: 4px;
                padding: 4px 16px;
                font-size: 12px;
                color: white;
            }}
            QPushButton:hover {{ background: #117a8b; }}
            QPushButton:disabled {{ background: #b0bec5; }}
        """)
        self._btn_excel.setEnabled(False)
        self._btn_excel.clicked.connect(self._open_excel)
        header.addWidget(self._btn_excel)

        root.addLayout(header)

        # ── Trennlinie ──────────────────────────────────────────────────────
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #d0d8e4;")
        root.addWidget(line)

        # ── Filter-Zeile: Jahr + Monat ───────────────────────────────────────
        _combo_style = """
            QComboBox {
                border: 1px solid #c8d2dc;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
                background: white;
                min-width: 90px;
            }
            QComboBox::drop-down { border: none; width: 18px; }
            QComboBox:focus { border: 1px solid #1565a8; }
        """
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        lbl_jahr = QLabel("Jahr:")
        lbl_jahr.setStyleSheet("color: #555; font-size: 12px;")
        filter_row.addWidget(lbl_jahr)

        self._cb_jahr = QComboBox()
        self._cb_jahr.setStyleSheet(_combo_style)
        filter_row.addWidget(self._cb_jahr)

        lbl_monat = QLabel("Monat:")
        lbl_monat.setStyleSheet("color: #555; font-size: 12px;")
        filter_row.addWidget(lbl_monat)

        self._cb_monat = QComboBox()
        self._cb_monat.setStyleSheet(_combo_style)
        # Alle Monate + "Alle"
        _MONATE = [
            "Alle",
            "Januar", "Februar", "März", "April",
            "Mai", "Juni", "Juli", "August",
            "September", "Oktober", "November", "Dezember",
        ]
        self._cb_monat.addItems(_MONATE)
        filter_row.addWidget(self._cb_monat)

        filter_row.addStretch()
        root.addLayout(filter_row)

        # ── Info-Label ──────────────────────────────────────────────────────
        self._info_lbl = QLabel()
        self._info_lbl.setStyleSheet("color: #666; font-size: 11px;")
        root.addWidget(self._info_lbl)

        # ── Eintrags-Liste ──────────────────────────────────────────────────
        self._list = QListWidget()
        self._list.setStyleSheet("""
            QListWidget {
                border: 1px solid #d0d8e4;
                border-radius: 4px;
                background: white;
                font-size: 12px;
                font-family: 'Segoe UI';
            }
            QListWidget::item {
                padding: 8px 10px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:selected {
                background: #e3eaf5;
                color: #1565a8;
            }
            QListWidget::item:hover:!selected {
                background: #f5f8fc;
            }
        """)
        self._list.itemSelectionChanged.connect(self._on_selection_changed)
        self._list.itemDoubleClicked.connect(self._laden_selected)
        root.addWidget(self._list, 1)

        # ── Hinweis ─────────────────────────────────────────────────────────
        hint = QLabel("💡 Doppelklick auf einen Eintrag lädt die Eingaben in das Sonderaufgaben-Formular.")
        hint.setStyleSheet("color: #888; font-size: 10px; font-style: italic;")
        root.addWidget(hint)

        # Jahres-Filter befüllen + auf aktuellen Monat setzen (nach Aufbau aller Widgets)
        self._init_filter()

    # ── Verbindung zum Sonderaufgaben-Widget ────────────────────────────────

    def set_sonder_widget(self, widget) -> None:
        """Setzt die Referenz auf den SonderaufgabenWidget (nach Erstellung)."""
        self._sonder_widget = widget

    # ── Filter initialisieren ────────────────────────────────────────────────

    def _init_filter(self) -> None:
        """Jahres-Dropdown befüllen und Standard auf aktuellen Monat setzen."""
        from database import sonderaufgaben_db
        jetzt = _dt.now()
        aktuelles_jahr = str(jetzt.year)
        aktueller_monat = jetzt.month   # 1-basiert → Index in cb_monat

        # Verfügbare Jahre aus DB holen
        try:
            alle = sonderaufgaben_db.get_snapshots(limit=5000)
            jahre = sorted(
                {s["gespeichert_am"][:4] for s in alle if s.get("gespeichert_am")},
                reverse=True,
            )
        except Exception:
            jahre = []

        # Aktuelles Jahr sicherstellen
        if aktuelles_jahr not in jahre:
            jahre.insert(0, aktuelles_jahr)

        # Combo befüllen (kein Signal feuern während des Befüllens)
        self._cb_jahr.blockSignals(True)
        self._cb_jahr.clear()
        self._cb_jahr.addItem("Alle")
        self._cb_jahr.addItems(jahre)
        # Aktuelles Jahr vorwählen
        idx = self._cb_jahr.findText(aktuelles_jahr)
        self._cb_jahr.setCurrentIndex(idx if idx >= 0 else 0)
        self._cb_jahr.blockSignals(False)

        # Aktuellen Monat vorwählen (Index 0 = "Alle", 1 = Januar …)
        self._cb_monat.setCurrentIndex(aktueller_monat)

        # Signale verbinden (einmalig nach Aufbau, damit kein doppeltes Laden)
        self._cb_jahr.currentIndexChanged.connect(self._load_liste)
        self._cb_monat.currentIndexChanged.connect(self._load_liste)

        self._load_liste()

    def _refresh_filter(self) -> None:
        """Jahres-Dropdown neu befüllen und Liste aktualisieren."""
        from database import sonderaufgaben_db
        jetzt = _dt.now()
        aktuelles_jahr = str(jetzt.year)
        try:
            alle = sonderaufgaben_db.get_snapshots(limit=5000)
            jahre = sorted(
                {s["gespeichert_am"][:4] for s in alle if s.get("gespeichert_am")},
                reverse=True,
            )
        except Exception:
            jahre = []
        if aktuelles_jahr not in jahre:
            jahre.insert(0, aktuelles_jahr)

        aktuell_jahr_text  = self._cb_jahr.currentText()
        aktuell_monat_idx  = self._cb_monat.currentIndex()

        self._cb_jahr.blockSignals(True)
        self._cb_jahr.clear()
        self._cb_jahr.addItem("Alle")
        self._cb_jahr.addItems(jahre)
        idx = self._cb_jahr.findText(aktuell_jahr_text)
        self._cb_jahr.setCurrentIndex(idx if idx >= 0 else 0)
        self._cb_jahr.blockSignals(False)

        self._cb_monat.setCurrentIndex(aktuell_monat_idx)
        self._load_liste()

    # ── Daten laden ─────────────────────────────────────────────────────────

    def _load_liste(self) -> None:
        from database import sonderaufgaben_db
        self._list.clear()
        self._btn_laden.setEnabled(False)
        self._btn_excel.setEnabled(False)

        # Filter auslesen
        filter_jahr  = self._cb_jahr.currentText()   # "Alle" oder "2026"
        filter_monat = self._cb_monat.currentIndex() # 0 = Alle, 1-12 = Monat

        try:
            snapshots = sonderaufgaben_db.get_snapshots(limit=5000)
        except Exception as exc:
            self._info_lbl.setText(f"⚠️ Datenbankfehler: {exc}")
            return

        # Filtern: gespeichert_am ist "YYYY-MM-DD HH:MM:SS"
        if filter_jahr != "Alle":
            snapshots = [s for s in snapshots
                         if s.get("gespeichert_am", "").startswith(filter_jahr)]
        if filter_monat > 0:
            monat_str = f"-{filter_monat:02d}-"
            snapshots = [s for s in snapshots
                         if monat_str in s.get("gespeichert_am", "")]

        self._info_lbl.setText(
            f"{len(snapshots)} Einträge gefunden  |  "
            "Doppelklick = in Formular laden"
        )
        for snap in snapshots:
            aktion_icon = "🖨️" if snap.get("aktion") == "drucken" else "💾"
            datum       = snap.get("datum", "")
            zeit        = snap.get("gespeichert_am", "")[-8:]          # HH:MM:SS
            bemerkung   = (snap.get("bemerkung") or "").replace("\n", " ")
            bem_kurz    = (bemerkung[:50] + "…") if len(bemerkung) > 50 else bemerkung
            hat_excel   = bool(snap.get("excel_datei", ""))
            excel_hint  = "  📊" if hat_excel else ""

            text = (
                f"{aktion_icon}  {datum}  ·  {zeit}{excel_hint}"
                + (f"   —  {bem_kurz}" if bem_kurz else "")
            )
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, snap.get("id"))
            item.setToolTip(
                f"Aktion: {snap.get('aktion', '')}  |  "
                f"Datum: {datum}  |  Gespeichert: {snap.get('gespeichert_am', '')}\n"
                f"Excel: {snap.get('excel_datei', '') or '—'}\n"
                f"Bemerkung: {bemerkung or '—'}"
            )
            self._list.addItem(item)

    def _on_selection_changed(self) -> None:
        items = self._list.selectedItems()
        has_sel = len(items) > 0
        self._btn_laden.setEnabled(has_sel)
        if has_sel:
            snap_id = items[0].data(Qt.ItemDataRole.UserRole)
            from database import sonderaufgaben_db
            try:
                snap = sonderaufgaben_db.get_snapshot_by_id(snap_id)
                has_excel = bool(snap and snap.get("excel_datei") and
                                 os.path.isfile(snap["excel_datei"]))
            except Exception:
                has_excel = False
            self._btn_excel.setEnabled(has_excel)
        else:
            self._btn_excel.setEnabled(False)

    # ── Aktionen ────────────────────────────────────────────────────────────

    def _laden_selected(self) -> None:
        """Snapshot in das Sonderaufgaben-Formular laden."""
        items = self._list.selectedItems()
        if not items:
            return
        snap_id = items[0].data(Qt.ItemDataRole.UserRole)
        from database import sonderaufgaben_db
        try:
            snap = sonderaufgaben_db.get_snapshot_by_id(snap_id)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Datenbankfehler:\n{exc}")
            return
        if snap is None:
            QMessageBox.warning(self, "Nicht gefunden", "Eintrag wurde nicht gefunden.")
            return
        if self._sonder_widget is None:
            QMessageBox.warning(self, "Fehler", "Sonderaufgaben-Widget nicht verbunden.")
            return

        # Daten laden und zum ersten Sub-Tab wechseln
        self._sonder_widget.load_from_db_snapshot(snap)
        # Tab zurück auf "Sonderaufgaben" (Index 0) wechseln
        inner_tabs = self.parent()
        if inner_tabs is not None:
            parent_widget = inner_tabs.parent()
            if hasattr(parent_widget, "_inner_tabs"):
                parent_widget._inner_tabs.setCurrentIndex(0)

        QMessageBox.information(
            self, "Geladen",
            f"Snapshot vom {snap.get('datum', '')} "
            f"({snap.get('gespeichert_am', '')}) wurde in das Formular geladen."
        )

    def _open_excel(self) -> None:
        """Excel-Datei des ausgewählten Snapshots öffnen."""
        items = self._list.selectedItems()
        if not items:
            return
        snap_id = items[0].data(Qt.ItemDataRole.UserRole)
        from database import sonderaufgaben_db
        try:
            snap = sonderaufgaben_db.get_snapshot_by_id(snap_id)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Datenbankfehler:\n{exc}")
            return
        if snap is None:
            return
        pfad = snap.get("excel_datei", "")
        if not pfad or not os.path.isfile(pfad):
            QMessageBox.warning(self, "Nicht gefunden",
                                "Die Excel-Datei dieses Eintrags wurde nicht gefunden.")
            return
        try:
            os.startfile(pfad)
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Öffnen:\n{exc}")


class _SonderaufgabenContainerWidget(QWidget):
    """Container mit zwei Sub-Tabs: Sonderaufgaben + Sonderaufgaben Historie."""

    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        inner_tabs = QTabWidget()
        inner_tabs.setFont(QFont("Segoe UI", 10))
        inner_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background: #f5f6f7;
            }
            QTabBar::tab {
                padding: 6px 16px;
                font-size: 11px;
                font-family: 'Segoe UI';
                color: #666;
                background: #eef1f5;
                border-bottom: 2px solid transparent;
                border-radius: 4px 4px 0 0;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #f5f6f7;
                color: #1565a8;
                font-weight: bold;
                border-bottom: 2px solid #1565a8;
            }
            QTabBar::tab:hover:!selected {
                background: #dde4ec;
                color: #1565a8;
            }
        """)

        self._sonderaufgaben = SonderaufgabenWidget()
        inner_tabs.addTab(self._sonderaufgaben, "📝  Sonderaufgaben")

        self._historie = _SonderaufgabenHistorieWidget()
        inner_tabs.addTab(self._historie, "🗂️  Sonderaufgaben Historie")

        # Referenz verdrahten – Historie kann jetzt Snapshots in das Formular laden
        self._historie.set_sonder_widget(self._sonderaufgaben)

        # Tab-Wechsel → Historie aktualisieren
        inner_tabs.currentChanged.connect(self._on_tab_changed)

        self._inner_tabs = inner_tabs
        root.addWidget(inner_tabs, 1)

    def _on_tab_changed(self, index: int) -> None:
        """Wenn auf den Historie-Tab gewechselt wird, Liste neu laden."""
        if index == 1:
            self._historie._refresh_filter()


class AufgabenWidget(QWidget):
    """Kombinierte Seite: Tabs für Checklisten, Sonderaufgaben und AOCC."""

    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Titel-Bar ──────────────────────────────────────────────────────
        title_bar = QWidget()
        title_bar.setStyleSheet("background-color: white; border-bottom: 1px solid #e0e0e0;")
        title_bar.setFixedHeight(52)
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(20, 0, 20, 0)
        lbl = QLabel("🌙 Aufgaben Nacht")
        lbl.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color: {FIORI_TEXT};")
        title_layout.addWidget(lbl)
        title_layout.addStretch()
        root.addWidget(title_bar)

        # ── Tab-Widget ─────────────────────────────────────────────────────
        tabs = QTabWidget()
        tabs.setDocumentMode(False)
        tabs.setFont(QFont("Segoe UI", 11))
        tabs.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background: #f5f6f7;
            }
            QTabBar::tab {
                padding: 8px 20px;
                font-size: 12px;
                font-family: 'Segoe UI';
                color: #666;
                background: #e8ecf0;
                border-bottom: 2px solid transparent;
                border-radius: 4px 4px 0 0;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #f5f6f7;
                color: #1565a8;
                font-weight: bold;
                border-bottom: 2px solid #1565a8;
            }
            QTabBar::tab:hover:!selected {
                background: #dde4ec;
                color: #1565a8;
            }
        """)

        # Tab 1 – Checklisten
        self._checklisten = ChecklistenWidget()
        tabs.addTab(self._checklisten, "📋  Checklisten & Drucksachen")

        # Tab 2 – Sonderaufgaben (mit Sub-Tabs: Sonderaufgaben + Historie)
        self._sonderaufgaben_container = _SonderaufgabenContainerWidget()
        tabs.addTab(self._sonderaufgaben_container, "📝  Sonderaufgaben")

        # Tab 3 – AOCC Lagebericht
        self._aocc = _AoccWidget()
        tabs.addTab(self._aocc, "📣  AOCC Lagebericht")

        # Tab 4 – Code 19 Mail
        self._code19_mail = _Code19MailTab()
        tabs.addTab(self._code19_mail, "📋  Code 19 Mail")

        self._tabs = tabs
        root.addWidget(tabs, 1)

    # ── Navigation-Refresh ─────────────────────────────────────────────────
    def refresh(self):
        """Wird beim Tab-Wechsel in der Sidebar aufgerufen."""
        self._checklisten.refresh()
        self._aocc._refresh()
