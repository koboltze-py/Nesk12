"""
Vorkommnisse-Widget
Erfassung und Verwaltung von Vorkommnis-Berichten (PRM, Passagiere, Chronologie).
Export als Word-Dokument mit Firmenlogo und Anschrift.
"""
from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime, date
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import Qt, QDate, QTime, QPoint
from PySide6.QtGui import QFont, QColor, QAction
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QDateEdit, QTimeEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QScrollArea, QFrame,
    QSplitter, QSizePolicy, QMessageBox, QFileDialog, QTabWidget, QMenu,
    QCheckBox,
)

from config import BASE_DIR, FIORI_BLUE, FIORI_TEXT

# ── Vorkommnis-Typen ───────────────────────────────────────────────────────────
VORKOMMNIS_TYPEN = [
    "PRM-Betreuung",
    "Medizinischer Notfall",
    "Sicherheitsvorfall",
    "Verspätung/Offblock",
    "Fahrzeugschaden",
    "Kommunikationsfehler",
    "Sonstiges",
]

# ── Bereich-Optionen (feingranulare Kategorie) ─────────────────────────────────
BEREICH_OPTIONEN = [
    "Flüge",
    "Fahrzeuge",
    "Personal",
    "Sachschäden",
    "Infrastruktur",
    "Sonstiges",
]

# ── Bereich → Bezeichnungs-Label (Formular & Word) ────────────────────────────
_BEREICH_LABEL: dict[str, str] = {
    "Flüge":        "Flugnummer:",
    "Fahrzeuge":    "Fahrzeug / Kennzeichen:",
    "Personal":     "Name / Dienstnummer:",
    "Sachschäden":  "Bezeichnung:",
    "Infrastruktur":"Bereich / Bezeichnung:",
    "Sonstiges":    "Bezeichnung:",
}

# ── PRM-Kategorien (IATA) ──────────────────────────────────────────────────────
PRM_KATEGORIEN = [
    "WCHS", "WCHR", "WCHC", "BLND", "DEAF", "DPNA",
    "UMNR", "STCR", "MEDA", "Sonstiges",
]

# ── Mitarbeiter-Rollen ─────────────────────────────────────────────────────────
MITARBEITER_ROLLEN = [
    "Disposition",
    "PRM-Begleitung",
    "Fahrer",
    "Servicepoint",
    "Leitung",
    "Sonstiges",
]

# ═══════════════════════════════════════════════════════════════════════════════
#  Stil-Helfer
# ═══════════════════════════════════════════════════════════════════════════════

_BTN_PRIM = f"""
    QPushButton {{
        background-color: {FIORI_BLUE};
        color: white;
        border: none;
        border-radius: 4px;
        padding: 6px 14px;
        font-size: 12px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: #1a6fc4;
    }}
    QPushButton:disabled {{
        background-color: #aaa;
    }}
"""

_BTN_DANGER = """
    QPushButton {
        background-color: #e74c3c;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 5px 12px;
        font-size: 12px;
    }
    QPushButton:hover {
        background-color: #c0392b;
    }
"""

_BTN_SEC = """
    QPushButton {
        background-color: #4a6480;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 6px 14px;
        font-size: 12px;
    }
    QPushButton:hover {
        background-color: #5a7490;
    }
"""

_TABLE_STYLE = """
    QTableWidget {
        border: 1px solid #d0d8e4;
        border-radius: 4px;
        gridline-color: #e0e6ef;
        background: white;
        font-size: 12px;
    }
    QTableWidget::item:selected {
        background-color: #dbeafe;
        color: #1b3a5c;
    }
    QHeaderView::section {
        background-color: #1565a8;
        color: white;
        font-weight: bold;
        font-size: 12px;
        padding: 5px;
        border: none;
    }
"""

_GROUP_STYLE = """
    QGroupBox {
        font-weight: bold;
        font-size: 13px;
        color: #1b3a5c;
        border: 1px solid #c5d0de;
        border-radius: 6px;
        margin-top: 10px;
        padding-top: 8px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 4px;
    }
"""

_LINE_EDIT_STYLE = """
    QLineEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit {
        border: 1px solid #c5d0de;
        border-radius: 4px;
        padding: 4px 8px;
        font-size: 12px;
        background: white;
    }
    QLineEdit:focus, QTextEdit:focus, QComboBox:focus,
    QDateEdit:focus, QTimeEdit:focus {
        border: 1.5px solid #1565a8;
    }
"""


# ═══════════════════════════════════════════════════════════════════════════════
#  Hilfsklasse: Editierbare Tabelle mit Hinzufügen/Löschen
# ═══════════════════════════════════════════════════════════════════════════════

class _EditTable(QWidget):
    """Generische editierbare Tabelle mit +/−-Buttons und Custom-Kontextmenü."""

    def __init__(
        self,
        headers: list[str],
        combo_columns: dict[int, list[str]] | None = None,
        column_widths: dict[int, int] | None = None,
        conditional_columns: dict[int, tuple[int, list[str]]] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._combo_columns = combo_columns or {}
        # conditional_columns: {ziel_col: (quell_col, [aktivierende Werte])}
        self._conditional_columns = conditional_columns or {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # Tabelle
        self._table = QTableWidget(0, len(headers))
        self._table.setHorizontalHeaderLabels(headers)
        self._table.setStyleSheet(_TABLE_STYLE)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked
                                    | QAbstractItemView.EditTrigger.SelectedClicked
                                    | QAbstractItemView.EditTrigger.AnyKeyPressed)
        self._table.verticalHeader().setVisible(False)
        self._table.setMinimumHeight(120)

        # Spaltenbreiten
        hdr = self._table.horizontalHeader()
        hdr.setStretchLastSection(True)
        for col in range(len(headers) - 1):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
            w = (column_widths or {}).get(col, 160)
            self._table.setColumnWidth(col, w)

        # Custom-Kontextmenü (ersetzt das OS-Standard-Menü)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._show_context_menu)

        layout.addWidget(self._table)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        self._btn_add = QPushButton("+ Hinzufügen")
        self._btn_add.setStyleSheet(_BTN_PRIM)
        self._btn_add.clicked.connect(self._add_row)
        self._btn_del = QPushButton("− Löschen")
        self._btn_del.setStyleSheet(_BTN_DANGER)
        self._btn_del.clicked.connect(self._del_row)
        btn_row.addWidget(self._btn_add)
        btn_row.addWidget(self._btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

    def _show_context_menu(self, pos: QPoint):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background: white; border: 1px solid #c5d0de;
                font-size: 12px;
            }
            QMenu::item { padding: 6px 20px; }
            QMenu::item:selected { background: #dbeafe; color: #1b3a5c; }
            QMenu::separator { height: 1px; background: #e0e6ef; margin: 2px 8px; }
        """)
        act_add  = QAction("✚  Zeile hinzufügen", self)
        act_del  = QAction("✖  Ausgewählte Zeile löschen", self)
        act_up   = QAction("↑  Zeile nach oben", self)
        act_down = QAction("↓  Zeile nach unten", self)
        act_add.triggered.connect(self._add_row)
        act_del.triggered.connect(self._del_row)
        act_up.triggered.connect(self._move_up)
        act_down.triggered.connect(self._move_down)
        menu.addAction(act_add)
        menu.addAction(act_del)
        menu.addSeparator()
        menu.addAction(act_up)
        menu.addAction(act_down)
        menu.exec(self._table.viewport().mapToGlobal(pos))

    def _move_up(self):
        row = self._table.currentRow()
        if row <= 0:
            return
        self._swap_rows(row, row - 1)
        self._table.selectRow(row - 1)

    def _move_down(self):
        row = self._table.currentRow()
        if row < 0 or row >= self._table.rowCount() - 1:
            return
        self._swap_rows(row, row + 1)
        self._table.selectRow(row + 1)

    def _swap_rows(self, a: int, b: int):
        cols = self._table.columnCount()
        for col in range(cols):
            wa = self._table.cellWidget(a, col)
            wb = self._table.cellWidget(b, col)
            if wa and wb and isinstance(wa, QComboBox) and isinstance(wb, QComboBox):
                va, vb = wa.currentText(), wb.currentText()
                ia = wa.findText(vb); wa.setCurrentIndex(ia if ia >= 0 else 0)
                ib = wb.findText(va); wb.setCurrentIndex(ib if ib >= 0 else 0)
            else:
                ia_item = self._table.item(a, col)
                ib_item = self._table.item(b, col)
                ta = ia_item.text() if ia_item else ""
                tb = ib_item.text() if ib_item else ""
                self._table.setItem(a, col, QTableWidgetItem(tb))
                self._table.setItem(b, col, QTableWidgetItem(ta))

    def _add_row(self):
        row = self._table.rowCount()
        self._table.insertRow(row)
        # Ziel-Spalten die standardmäßig deaktiviert starten
        conditional_targets = {tc for tc in self._conditional_columns}
        for col in range(self._table.columnCount()):
            if col in self._combo_columns:
                combo = QComboBox()
                combo.addItems(self._combo_columns[col])
                if col in conditional_targets:
                    combo.setEnabled(False)
                self._table.setCellWidget(row, col, combo)
            else:
                self._table.setItem(row, col, QTableWidgetItem(""))
        # Quell-Combos verbinden
        for target_col, (src_col, enabled_vals) in self._conditional_columns.items():
            src_w = self._table.cellWidget(row, src_col)
            if isinstance(src_w, QComboBox):
                src_w.currentTextChanged.connect(
                    lambda text, r=row, tc=target_col, ev=enabled_vals:
                        self._update_conditional(r, tc, text, ev)
                )
        self._table.scrollToBottom()
        self._table.selectRow(row)

    def _update_conditional(self, row: int, target_col: int, text: str, enabled_vals: list[str]):
        widget = self._table.cellWidget(row, target_col)
        if widget:
            enabled = text in enabled_vals
            widget.setEnabled(enabled)
            if not enabled:
                widget.setCurrentIndex(0)

    def _del_row(self):
        rows = sorted(
            {idx.row() for idx in self._table.selectedIndexes()},
            reverse=True,
        )
        for row in rows:
            self._table.removeRow(row)

    def get_data(self) -> list[list[str]]:
        result = []
        for row in range(self._table.rowCount()):
            cols = []
            for col in range(self._table.columnCount()):
                widget = self._table.cellWidget(row, col)
                if isinstance(widget, QComboBox):
                    cols.append(widget.currentText())
                else:
                    item = self._table.item(row, col)
                    cols.append(item.text().strip() if item else "")
            result.append(cols)
        return result

    def set_data(self, data: list[list[str]]):
        self._table.setRowCount(0)
        conditional_targets = {tc for tc in self._conditional_columns}
        for row_data in data:
            row = self._table.rowCount()
            self._table.insertRow(row)
            for col, value in enumerate(row_data):
                if col in self._combo_columns:
                    combo = QComboBox()
                    combo.addItems(self._combo_columns[col])
                    idx = combo.findText(value)
                    if idx >= 0:
                        combo.setCurrentIndex(idx)
                    if col in conditional_targets:
                        combo.setEnabled(False)  # vorerst; wird unten per Signal gesetzt
                    self._table.setCellWidget(row, col, combo)
                else:
                    self._table.setItem(row, col, QTableWidgetItem(value))
            # Quell-Combos verbinden und initialen Zustand setzen
            for target_col, (src_col, enabled_vals) in self._conditional_columns.items():
                src_w = self._table.cellWidget(row, src_col)
                if isinstance(src_w, QComboBox):
                    src_w.currentTextChanged.connect(
                        lambda text, r=row, tc=target_col, ev=enabled_vals:
                            self._update_conditional(r, tc, text, ev)
                    )
                    # Initialen Zustand anwenden
                    self._update_conditional(row, target_col, src_w.currentText(), enabled_vals)


# ═══════════════════════════════════════════════════════════════════════════════
#  Haupt-Widget: VorkommnisseWidget
# ═══════════════════════════════════════════════════════════════════════════════

class VorkommnisseWidget(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_id: int | None = None   # None = neuer Datensatz
        self.setStyleSheet("background-color: #f5f6f7;")
        self._build_ui()
        self._aktualisiere_liste()

    # ── UI-Aufbau ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header-Leiste
        header = QWidget()
        header.setFixedHeight(56)
        header.setStyleSheet(f"background-color: {FIORI_BLUE};")
        hdr_layout = QHBoxLayout(header)
        hdr_layout.setContentsMargins(20, 0, 20, 0)
        title_lbl = QLabel("⚠️  Vorkommnisse")
        title_lbl.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title_lbl.setStyleSheet("color: white;")
        hdr_layout.addWidget(title_lbl)
        hdr_layout.addStretch()
        self._btn_neu = QPushButton("🔄  Neues Formular")
        self._btn_neu.setStyleSheet(_BTN_SEC)
        self._btn_neu.clicked.connect(self._neues_formular)
        self._btn_speichern = QPushButton("💾  Speichern")
        self._btn_speichern.setStyleSheet(_BTN_PRIM)
        self._btn_speichern.clicked.connect(self._speichern)
        self._btn_export = QPushButton("📄  Word exportieren")
        self._btn_export.setStyleSheet(_BTN_SEC)
        self._btn_export.clicked.connect(self._export_word)
        self._btn_ordner = QPushButton("📁  Berichte öffnen")
        self._btn_ordner.setStyleSheet(_BTN_SEC)
        self._btn_ordner.setToolTip("Ordner mit allen Vorkommnis-Berichten öffnen")
        self._btn_ordner.clicked.connect(self._ordner_oeffnen)
        self._btn_email = QPushButton("✉️  E-Mail erstellen")
        self._btn_email.setStyleSheet(_BTN_SEC)
        self._btn_email.setToolTip("E-Mail direkt in Outlook öffnen")
        self._btn_email.clicked.connect(self._email_entwurf_dialog)
        hdr_layout.addWidget(self._btn_neu)
        hdr_layout.addWidget(self._btn_speichern)
        hdr_layout.addWidget(self._btn_export)
        hdr_layout.addWidget(self._btn_ordner)
        hdr_layout.addWidget(self._btn_email)
        root.addWidget(header)

        # Tab-Widget
        self._tabs = QTabWidget()
        self._tabs.setStyleSheet("""
            QTabWidget::pane { border: none; background: #f5f6f7; }
            QTabBar::tab {
                background: #354a5e; color: #cdd5e0;
                padding: 8px 22px; font-size: 13px;
                border: none; margin-right: 2px;
            }
            QTabBar::tab:selected { background: #1565a8; color: white; font-weight: bold; }
            QTabBar::tab:hover    { background: #4a6480; color: white; }
        """)
        self._tabs.addTab(self._build_erfassung_tab(), "✏️  Erfassung")
        self._tabs.addTab(self._build_liste_tab(),     "📋  Gespeicherte Vorkommnisse")
        self._tabs.currentChanged.connect(self._on_tab_changed)
        root.addWidget(self._tabs)

    # ── Tab 1: Erfassung ───────────────────────────────────────────────────────

    def _build_erfassung_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: #f5f6f7; }")
        content = QWidget()
        content.setStyleSheet("background: #f5f6f7;")
        c_layout = QVBoxLayout(content)
        c_layout.setContentsMargins(20, 16, 20, 20)
        c_layout.setSpacing(14)

        self._status_lbl = QLabel("🟢  Neuer Datensatz")
        self._status_lbl.setStyleSheet(
            "color: #27ae60; font-size: 12px; font-weight: bold;"
        )
        c_layout.addWidget(self._status_lbl)

        c_layout.addWidget(self._build_grunddaten())

        grp_pax = QGroupBox("1. Betroffene Personen")
        grp_pax.setStyleSheet(_GROUP_STYLE)
        pax_layout = QVBoxLayout(grp_pax)
        pax_layout.setContentsMargins(10, 14, 10, 10)
        self._pax_table = _EditTable(
            headers=["Person (Name, Vorname)", "Typ", "Kategorie", "Anmerkung"],
            combo_columns={
                1: ["Passagier", "PRM Passagier", "Patient", "Mitarbeiter", "Sonstige"],
                2: PRM_KATEGORIEN,
            },
            column_widths={0: 200, 1: 130, 2: 110},
            conditional_columns={2: (1, ["PRM Passagier"])},
        )
        pax_layout.addWidget(self._pax_table)
        c_layout.addWidget(grp_pax)

        grp_personal = QGroupBox("2. Eingeteiltes Personal")
        grp_personal.setStyleSheet(_GROUP_STYLE)
        pers_layout = QVBoxLayout(grp_personal)
        pers_layout.setContentsMargins(10, 14, 10, 10)
        self._personal_table = _EditTable(
            headers=["Name (Mitarbeiter)", "Funktion/Rolle", "Anmerkung"],
            combo_columns={1: MITARBEITER_ROLLEN},
            column_widths={0: 200, 1: 150},
        )
        pers_layout.addWidget(self._personal_table)
        c_layout.addWidget(grp_personal)

        grp_chrono = QGroupBox("3. Chronologischer Ablauf")
        grp_chrono.setStyleSheet(_GROUP_STYLE)
        chrono_layout = QVBoxLayout(grp_chrono)
        chrono_layout.setContentsMargins(10, 14, 10, 10)
        self._chrono_table = _EditTable(
            headers=["Uhrzeit", "Ereignis"],
            column_widths={0: 110},
        )
        chrono_layout.addWidget(self._chrono_table)
        c_layout.addWidget(grp_chrono)

        grp_ursache = QGroupBox("4. Ursachenanalyse")
        grp_ursache.setStyleSheet(_GROUP_STYLE)
        ursache_layout = QVBoxLayout(grp_ursache)
        ursache_layout.setContentsMargins(10, 14, 10, 10)
        self._ursache_edit = QTextEdit()
        self._ursache_edit.setPlaceholderText(
            "Ursachen des Vorkommnisses stichpunktartig beschreiben ..."
        )
        self._ursache_edit.setMinimumHeight(100)
        self._ursache_edit.setStyleSheet(_LINE_EDIT_STYLE)
        ursache_layout.addWidget(self._ursache_edit)
        c_layout.addWidget(grp_ursache)

        grp_ergebnis = QGroupBox("5. Ergebnis")
        grp_ergebnis.setStyleSheet(_GROUP_STYLE)
        erg_layout = QVBoxLayout(grp_ergebnis)
        erg_layout.setContentsMargins(10, 14, 10, 10)
        self._ergebnis_edit = QTextEdit()
        self._ergebnis_edit.setPlaceholderText(
            "Ergebnis und Ausgang des Vorkommnisses ..."
        )
        self._ergebnis_edit.setMinimumHeight(80)
        self._ergebnis_edit.setStyleSheet(_LINE_EDIT_STYLE)
        erg_layout.addWidget(self._ergebnis_edit)
        c_layout.addWidget(grp_ergebnis)

        c_layout.addStretch()
        scroll.setWidget(content)
        return scroll

    _MONATE = [
        "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember",
    ]

    # ── Tab 2: Übersicht ───────────────────────────────────────────────────────

    def _build_liste_tab(self) -> QWidget:
        from datetime import date as _date
        from PySide6.QtWidgets import QSpinBox
        _heute = _date.today()

        w = QWidget()
        w.setStyleSheet("background: #f5f6f7;")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        # ── Filter-Leiste ─────────────────────────────────────────────────────
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        lbl_monat = QLabel("Monat:")
        lbl_monat.setStyleSheet("font-weight: bold; color: #1b3a5c;")
        self._filter_monat = QComboBox()
        self._filter_monat.addItem("Alle Monate", 0)
        for i, name in enumerate(self._MONATE, start=1):
            self._filter_monat.addItem(name, i)
        self._filter_monat.setCurrentIndex(_heute.month)  # 1-12 → Index 1-12
        self._filter_monat.setFixedWidth(140)
        self._filter_monat.setStyleSheet(_LINE_EDIT_STYLE)

        lbl_jahr = QLabel("Jahr:")
        lbl_jahr.setStyleSheet("font-weight: bold; color: #1b3a5c;")
        self._filter_jahr = QSpinBox()
        self._filter_jahr.setRange(2020, 2040)
        self._filter_jahr.setValue(_heute.year)
        self._filter_jahr.setFixedWidth(80)
        self._filter_jahr.setStyleSheet(_LINE_EDIT_STYLE)

        btn_alle = QPushButton("🔄  Alle anzeigen")
        btn_alle.setStyleSheet(_BTN_SEC)
        btn_alle.setToolTip("Monatsfilter aufheben – alle Vorkommnisse anzeigen")
        btn_alle.clicked.connect(self._filter_alle_anzeigen)

        self._filter_zaehler = QLabel("")
        self._filter_zaehler.setStyleSheet("color: #555; font-size: 11px;")

        filter_row.addWidget(lbl_monat)
        filter_row.addWidget(self._filter_monat)
        filter_row.addSpacing(12)
        filter_row.addWidget(lbl_jahr)
        filter_row.addWidget(self._filter_jahr)
        filter_row.addSpacing(8)
        filter_row.addWidget(btn_alle)
        filter_row.addStretch()
        filter_row.addWidget(self._filter_zaehler)
        layout.addLayout(filter_row)

        self._filter_monat.currentIndexChanged.connect(self._aktualisiere_liste)
        self._filter_jahr.valueChanged.connect(self._aktualisiere_liste)

        # ── Toolbar ───────────────────────────────────────────────────────────
        toolbar = QHBoxLayout()
        btn_laden = QPushButton("📥  Öffnen / Bearbeiten")
        btn_laden.setStyleSheet(_BTN_PRIM)
        btn_laden.setToolTip("Ausgewähltes Vorkommnis in den Erfassungs-Tab laden")
        btn_laden.clicked.connect(self._laden_aus_liste)
        btn_loeschen = QPushButton("🗑  Löschen")
        btn_loeschen.setStyleSheet(_BTN_DANGER)
        btn_loeschen.setToolTip("Ausgewähltes Vorkommnis dauerhaft löschen")
        btn_loeschen.clicked.connect(self._loeschen_aus_liste)
        btn_eigener = QPushButton("➕  Eigener Eintrag")
        btn_eigener.setStyleSheet(_BTN_SEC)
        btn_eigener.setToolTip("Eigenen Eintrag anlegen und optional Word-Datei verknüpfen")
        btn_eigener.clicked.connect(self._eigener_eintrag_dialog)
        btn_word_oeffnen = QPushButton("📄  Word öffnen")
        btn_word_oeffnen.setStyleSheet(_BTN_SEC)
        btn_word_oeffnen.setToolTip("Verknüpfte Word-Datei des ausgewählten Eintrags öffnen")
        btn_word_oeffnen.clicked.connect(self._word_datei_oeffnen)
        toolbar.addWidget(btn_laden)
        toolbar.addWidget(btn_loeschen)
        toolbar.addSpacing(16)
        toolbar.addWidget(btn_eigener)
        toolbar.addWidget(btn_word_oeffnen)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        self._liste_table = QTableWidget(0, 8)
        self._liste_table.setHorizontalHeaderLabels(
            ["ID", "Flug/Ref.", "Typ", "Bereich", "Datum", "Ort", "Word", "Erstellt von"]
        )
        self._liste_table.setStyleSheet(_TABLE_STYLE)
        self._liste_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._liste_table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self._liste_table.verticalHeader().setVisible(False)
        self._liste_table.horizontalHeader().setStretchLastSection(True)
        self._liste_table.setColumnWidth(0, 45)
        self._liste_table.setColumnWidth(1, 90)
        self._liste_table.setColumnWidth(2, 160)
        self._liste_table.setColumnWidth(3, 110)
        self._liste_table.setColumnWidth(4, 100)
        self._liste_table.setColumnWidth(5, 140)
        self._liste_table.setColumnWidth(6, 65)
        self._liste_table.doubleClicked.connect(self._liste_doppelklick)
        self._liste_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._liste_table.customContextMenuRequested.connect(
            self._liste_context_menu
        )
        layout.addWidget(self._liste_table)
        return w

    def _liste_context_menu(self, pos: QPoint):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background: white; border: 1px solid #c5d0de; font-size: 12px; }"
            "QMenu::item { padding: 6px 20px; }"
            "QMenu::item:selected { background: #dbeafe; color: #1b3a5c; }"
        )
        act_open   = QAction("📥  Öffnen / Bearbeiten", self)
        act_del    = QAction("🗑  Löschen", self)
        act_word   = QAction("📄  Word-Datei öffnen", self)
        act_link   = QAction("🔗  Word-Datei verknüpfen …", self)
        act_unlink = QAction("✖  Word-Verknüpfung entfernen", self)
        act_open.triggered.connect(self._laden_aus_liste)
        act_del.triggered.connect(self._loeschen_aus_liste)
        act_word.triggered.connect(self._word_datei_oeffnen)
        act_link.triggered.connect(self._word_datei_verknuepfen)
        act_unlink.triggered.connect(self._word_datei_entfernen)
        menu.addAction(act_open)
        menu.addAction(act_del)
        menu.addSeparator()
        menu.addAction(act_word)
        menu.addAction(act_link)
        menu.addAction(act_unlink)
        menu.exec(self._liste_table.viewport().mapToGlobal(pos))

    def _build_grunddaten(self) -> QGroupBox:
        grp = QGroupBox("Grunddaten des Vorkommnisses")
        grp.setStyleSheet(_GROUP_STYLE)
        layout = QFormLayout(grp)
        layout.setContentsMargins(14, 18, 14, 12)
        layout.setSpacing(10)
        layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        style = _LINE_EDIT_STYLE

        self._flug_edit = QLineEdit()
        self._flug_edit.setPlaceholderText("z. B. XQ983 / PRM-Betreuung")
        self._flug_edit.setStyleSheet(style)
        layout.addRow("Flugnummer / Vorkommnis:", self._flug_edit)
        # Label-Referenz für dynamische Aktualisierung
        self._flug_lbl = layout.labelForField(self._flug_edit)

        self._typ_combo = QComboBox()
        self._typ_combo.addItems(VORKOMMNIS_TYPEN)
        self._typ_combo.setStyleSheet(style)
        layout.addRow("Vorkommnis-Typ:", self._typ_combo)

        self._bereich_combo = QComboBox()
        self._bereich_combo.addItems(BEREICH_OPTIONEN)
        self._bereich_combo.setStyleSheet(style)
        self._bereich_combo.setToolTip(
            "Kategorie des Vorkommnisses – wird im Word-Bericht verwendet"
        )
        layout.addRow("Bereich:", self._bereich_combo)
        self._bereich_combo.currentTextChanged.connect(self._on_bereich_changed)

        self._datum_edit = QDateEdit(QDate.currentDate())
        self._datum_edit.setDisplayFormat("dd.MM.yyyy")
        self._datum_edit.setCalendarPopup(True)
        self._datum_edit.setStyleSheet(style)
        layout.addRow("Datum:", self._datum_edit)

        self._ort_edit = QLineEdit()
        self._ort_edit.setPlaceholderText("z. B. Köln/Bonn (CGN)")
        self._ort_edit.setStyleSheet(style)
        layout.addRow("Ort:", self._ort_edit)

        self._offblock_plan_edit = QTimeEdit()
        self._offblock_plan_edit.setDisplayFormat("HH:mm")
        self._offblock_plan_edit.setStyleSheet(style)
        self._offblock_plan_edit.setEnabled(False)
        self._offblock_plan_chk = QCheckBox("angeben")
        self._offblock_plan_chk.setChecked(False)
        self._offblock_plan_chk.toggled.connect(self._offblock_plan_edit.setEnabled)
        _plan_row = QWidget()
        _plan_hl  = QHBoxLayout(_plan_row)
        _plan_hl.setContentsMargins(0, 0, 0, 0)
        _plan_hl.addWidget(self._offblock_plan_edit)
        _plan_hl.addWidget(self._offblock_plan_chk)
        layout.addRow("Geplanter Offblock:", _plan_row)

        self._offblock_ist_edit = QTimeEdit()
        self._offblock_ist_edit.setDisplayFormat("HH:mm")
        self._offblock_ist_edit.setStyleSheet(style)
        self._offblock_ist_edit.setEnabled(False)
        self._offblock_ist_chk = QCheckBox("angeben")
        self._offblock_ist_chk.setChecked(False)
        self._offblock_ist_chk.toggled.connect(self._offblock_ist_edit.setEnabled)
        _ist_row = QWidget()
        _ist_hl  = QHBoxLayout(_ist_row)
        _ist_hl.setContentsMargins(0, 0, 0, 0)
        _ist_hl.addWidget(self._offblock_ist_edit)
        _ist_hl.addWidget(self._offblock_ist_chk)
        layout.addRow("Tatsächlicher Offblock:", _ist_row)

        self._erstellt_von_edit = QLineEdit()
        self._erstellt_von_edit.setPlaceholderText("Name des Verfassers")
        self._erstellt_von_edit.setStyleSheet(style)
        layout.addRow("Erstellt von:", self._erstellt_von_edit)

        return grp

    # ── DB-Aktionen ────────────────────────────────────────────────────────────

    def _speichern(self):
        from functions.vorkommnisse_db import speichern, aktualisieren
        daten = self._sammle_daten()
        if not daten["flug"] and not daten["bereich"]:
            QMessageBox.warning(self, "Fehler", "Bitte Bezeichnung / Flugnummer eingeben.")
            return
        try:
            if self._current_id is None:
                self._current_id = speichern(daten)
                QMessageBox.information(
                    self, "Gespeichert",
                    f"Vorkommnis »{daten['flug'] or daten['bereich']}« wurde gespeichert."
                )
            else:
                aktualisieren(self._current_id, daten)
                QMessageBox.information(
                    self, "Aktualisiert",
                    f"Vorkommnis #{self._current_id} wurde aktualisiert."
                )
            self._aktualisiere_status()
            self._aktualisiere_liste()
        except Exception as exc:
            QMessageBox.critical(self, "Fehler beim Speichern", str(exc))

    def _neues_formular(self):
        self._current_id = None
        self._formular_leeren()
        self._aktualisiere_status()
        self._tabs.setCurrentIndex(0)

    def _laden_aus_liste(self):
        from functions.vorkommnisse_db import lade_ein
        row = self._liste_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst ein Vorkommnis auswählen.")
            return
        id_item = self._liste_table.item(row, 0)
        if not id_item:
            return
        vid = int(id_item.text())
        daten = lade_ein(vid)
        if daten is None:
            QMessageBox.warning(self, "Fehler", "Datensatz nicht gefunden.")
            return
        self._current_id = vid
        self._formular_befuellen(daten)
        self._aktualisiere_status()
        self._tabs.setCurrentIndex(0)

    def _loeschen_aus_liste(self):
        from functions.vorkommnisse_db import loeschen
        row = self._liste_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst ein Vorkommnis auswählen.")
            return
        id_item  = self._liste_table.item(row, 0)
        flg_item = self._liste_table.item(row, 1)
        if not id_item:
            return
        vid  = int(id_item.text())
        flug = flg_item.text() if flg_item else str(vid)
        antwort = QMessageBox.question(
            self, "Löschen bestätigen",
            f"Vorkommnis Flug »{flug}« wirklich löschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if antwort != QMessageBox.StandardButton.Yes:
            return
        loeschen(vid)
        if self._current_id == vid:
            self._current_id = None
            self._formular_leeren()
            self._aktualisiere_status()
        self._aktualisiere_liste()

    def _filter_alle_anzeigen(self):
        """Monatsfilter auf \"Alle\" setzen."""
        self._filter_monat.setCurrentIndex(0)   # "Alle Monate"

    def _aktualisiere_liste(self):
        try:
            from functions.vorkommnisse_db import lade_alle
            alle = lade_alle()
        except Exception:
            alle = []

        # Aktiven Filter auslesen
        sel_monat = self._filter_monat.currentData()   # 0 = alle, 1-12 = Monat
        sel_jahr  = self._filter_jahr.value()

        gefiltert = []
        for d in alle:
            if sel_monat == 0:
                gefiltert.append(d)
                continue
            # Datum im Format dd.MM.yyyy parsen
            try:
                from datetime import datetime as _dt
                dat = _dt.strptime(d["datum"], "%d.%m.%Y")
                if dat.month == sel_monat and dat.year == sel_jahr:
                    gefiltert.append(d)
            except ValueError:
                pass  # unbekanntes Datumsformat → überspringen

        # Zähler aktualisieren
        gesamt = len(alle)
        angezeigt = len(gefiltert)
        if sel_monat == 0:
            self._filter_zaehler.setText(f"{gesamt} Einträge")
        else:
            monat_name = self._MONATE[sel_monat - 1]
            self._filter_zaehler.setText(
                f"{angezeigt} von {gesamt} • {monat_name} {sel_jahr}"
            )

        self._liste_table.setRowCount(0)
        for d in gefiltert:
            r = self._liste_table.rowCount()
            self._liste_table.insertRow(r)
            # Spalten 0-5: ID, Flug, Typ, Bereich, Datum, Ort
            for col, val in enumerate([
                str(d["id"]), d["flug"], d["typ"], d.get("bereich", ""),
                d["datum"], d["ort"],
            ]):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if d["id"] == self._current_id:
                    item.setBackground(QColor("#dbeafe"))
                self._liste_table.setItem(r, col, item)
            # Spalte 6: Word-Datei Indikator
            word_datei = d.get("word_datei", "")
            word_item = QTableWidgetItem("📄" if word_datei else "−")
            word_item.setFlags(word_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            word_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if word_datei:
                abs_pfad = self._berichte_basis_dir() / word_datei
                word_item.setToolTip(str(abs_pfad))
                word_item.setData(Qt.ItemDataRole.UserRole, word_datei)
            if d["id"] == self._current_id:
                word_item.setBackground(QColor("#dbeafe"))
            self._liste_table.setItem(r, 6, word_item)
            # Spalte 7: Erstellt von
            ev_item = QTableWidgetItem(d["erstellt_von"])
            ev_item.setFlags(ev_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if d["id"] == self._current_id:
                ev_item.setBackground(QColor("#dbeafe"))
            self._liste_table.setItem(r, 7, ev_item)

    def _aktualisiere_status(self):
        if self._current_id is None:
            self._status_lbl.setText("🟢  Neuer Datensatz")
            self._status_lbl.setStyleSheet(
                "color: #27ae60; font-size: 12px; font-weight: bold;"
            )
        else:
            self._status_lbl.setText(f"🟡  Bearbeitung: Datensatz #{self._current_id}")
            self._status_lbl.setStyleSheet(
                "color: #e67e22; font-size: 12px; font-weight: bold;"
            )

    def _on_tab_changed(self, idx: int):
        if idx == 1:
            self._aktualisiere_liste()

    def _on_bereich_changed(self, bereich: str):
        """Passt den Label-Text des Flug/Referenz-Feldes an den Bereich an."""
        lbl = _BEREICH_LABEL.get(bereich, "Bezeichnung / Flugnummer:")
        if self._flug_lbl:
            self._flug_lbl.setText(lbl)

    # ── Formular-Helfer ────────────────────────────────────────────────────────

    def _liste_doppelklick(self, index):
        """Doppelklick auf Word-Spalte öffnet die Datei, sonst wird der Eintrag geladen."""
        if index.column() == 6:
            self._word_datei_oeffnen()
        else:
            self._laden_aus_liste()

    def _word_datei_oeffnen(self):
        """Öffnet die verknüpfte Word-Datei des ausgewählten Eintrags."""
        row = self._liste_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst einen Eintrag auswählen.")
            return
        word_item = self._liste_table.item(row, 6)
        word_datei = word_item.data(Qt.ItemDataRole.UserRole) if word_item else None
        if not word_datei:
            QMessageBox.information(self, "Keine Word-Datei", "Für diesen Eintrag ist keine Word-Datei verknüpft.")
            return
        abs_pfad = self._berichte_basis_dir() / word_datei
        if not abs_pfad.exists():
            QMessageBox.warning(self, "Datei nicht gefunden",
                                f"Die verknüpfte Datei wurde nicht gefunden:\n{abs_pfad}")
            return
        os.startfile(str(abs_pfad))

    def _word_datei_verknuepfen(self):
        """Wählt eine Word-Datei per Windows-Dialog aus, kopiert sie in den Berichte-Ordner
        und verknüpft sie mit dem ausgewählten Eintrag."""
        row = self._liste_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst einen Eintrag auswählen.")
            return
        id_item = self._liste_table.item(row, 0)
        if not id_item:
            return
        vid = int(id_item.text())
        datum_item = self._liste_table.item(row, 4)
        datum = datum_item.text() if datum_item else datetime.today().strftime("%d.%m.%Y")

        pfad, _ = QFileDialog.getOpenFileName(
            self, "Word-Datei auswählen", "",
            "Word-Dokument (*.docx *.doc)",
        )
        if not pfad:
            return

        try:
            ziel_dir = self._monats_dir(datum)
            dateiname = os.path.basename(pfad)
            ziel_pfad = ziel_dir / dateiname
            # Überschreiben vermeiden: Zähler anhängen
            if ziel_pfad.exists():
                try:
                    same = os.path.samefile(pfad, str(ziel_pfad))
                except OSError:
                    same = False
                if not same:
                    stem, suffix = Path(dateiname).stem, Path(dateiname).suffix
                    i = 1
                    while ziel_pfad.exists():
                        ziel_pfad = ziel_dir / f"{stem}_{i}{suffix}"
                        i += 1
                    shutil.copy2(pfad, ziel_pfad)
            else:
                shutil.copy2(pfad, ziel_pfad)
            word_rel = ziel_pfad.relative_to(self._berichte_basis_dir()).as_posix()
        except Exception as exc:
            QMessageBox.critical(self, "Fehler beim Kopieren", str(exc))
            return

        from functions.vorkommnisse_db import setze_word_datei
        try:
            setze_word_datei(vid, word_rel)
            self._aktualisiere_liste()
            QMessageBox.information(self, "Verknüpft",
                                    f"Word-Datei wurde verknüpft:\n{os.path.basename(str(ziel_pfad))}")
        except Exception as exc:
            QMessageBox.critical(self, "Fehler", str(exc))

    def _word_datei_entfernen(self):
        """Entfernt die Word-Verknüpfung vom ausgewählten Eintrag (Datei bleibt erhalten)."""
        row = self._liste_table.currentRow()
        if row < 0:
            return
        id_item = self._liste_table.item(row, 0)
        if not id_item:
            return
        vid = int(id_item.text())
        antwort = QMessageBox.question(
            self, "Verknüpfung entfernen",
            "Die Word-Verknüpfung für diesen Eintrag wirklich entfernen?\n"
            "(Die Datei selbst bleibt erhalten.)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if antwort != QMessageBox.StandardButton.Yes:
            return
        from functions.vorkommnisse_db import setze_word_datei
        setze_word_datei(vid, "")
        self._aktualisiere_liste()

    def _eigener_eintrag_dialog(self):
        """Öffnet einen Dialog zum Anlegen eines eigenen Eintrags mit optionaler Word-Datei."""
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QFormLayout, QDialogButtonBox,
        )

        dlg = QDialog(self)
        dlg.setWindowTitle("Eigener Eintrag anlegen")
        dlg.setMinimumWidth(520)
        dlg.setStyleSheet("background: #f5f6f7;")

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(10)

        form = QFormLayout()
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)

        bezeichnung = QLineEdit()
        bezeichnung.setPlaceholderText("z. B. XQ983 / Fahrzeug C58 / …")
        bezeichnung.setStyleSheet(_LINE_EDIT_STYLE)

        typ_combo = QComboBox()
        typ_combo.addItems(VORKOMMNIS_TYPEN)
        typ_combo.setStyleSheet(_LINE_EDIT_STYLE)

        bereich_combo = QComboBox()
        bereich_combo.addItems(BEREICH_OPTIONEN)
        bereich_combo.setStyleSheet(_LINE_EDIT_STYLE)

        datum_edit = QDateEdit(QDate.currentDate())
        datum_edit.setCalendarPopup(True)
        datum_edit.setDisplayFormat("dd.MM.yyyy")
        datum_edit.setStyleSheet(_LINE_EDIT_STYLE)

        ort_edit = QLineEdit()
        ort_edit.setPlaceholderText("z. B. Köln/Bonn (CGN)")
        ort_edit.setStyleSheet(_LINE_EDIT_STYLE)

        erstellt_von_edit = QLineEdit()
        erstellt_von_edit.setPlaceholderText("Name des Verfassers")
        erstellt_von_edit.setStyleSheet(_LINE_EDIT_STYLE)

        # Word-Datei Auswahl
        word_pfad_edit = QLineEdit()
        word_pfad_edit.setReadOnly(True)
        word_pfad_edit.setPlaceholderText("(keine Word-Datei ausgewählt)")
        word_pfad_edit.setStyleSheet(_LINE_EDIT_STYLE)
        btn_word_waehlen = QPushButton("📂  Auswählen …")
        btn_word_waehlen.setStyleSheet(_BTN_SEC)

        word_row_widget = QWidget()
        word_hl = QHBoxLayout(word_row_widget)
        word_hl.setContentsMargins(0, 0, 0, 0)
        word_hl.setSpacing(6)
        word_hl.addWidget(word_pfad_edit, 1)
        word_hl.addWidget(btn_word_waehlen)

        form.addRow("Bezeichnung / Flugnr.:", bezeichnung)
        form.addRow("Typ:", typ_combo)
        form.addRow("Bereich:", bereich_combo)
        form.addRow("Datum:", datum_edit)
        form.addRow("Ort:", ort_edit)
        form.addRow("Erstellt von:", erstellt_von_edit)
        form.addRow("Word-Datei (optional):", word_row_widget)
        layout.addLayout(form)

        selected_file: list[str | None] = [None]

        def _waehle_word():
            pfad, _ = QFileDialog.getOpenFileName(
                dlg, "Word-Datei auswählen", "",
                "Word-Dokument (*.docx *.doc)",
            )
            if pfad:
                selected_file[0] = pfad
                word_pfad_edit.setText(os.path.basename(pfad))

        btn_word_waehlen.clicked.connect(_waehle_word)

        btn_box = QDialogButtonBox()
        btn_speichern = btn_box.addButton("💾  Speichern", QDialogButtonBox.ButtonRole.AcceptRole)
        btn_ab = btn_box.addButton("Abbrechen", QDialogButtonBox.ButtonRole.RejectRole)
        btn_box.rejected.connect(dlg.reject)

        def _do_speichern():
            if not bezeichnung.text().strip():
                QMessageBox.warning(dlg, "Fehler", "Bitte Bezeichnung / Flugnummer eingeben.")
                return
            dlg.accept()

        btn_speichern.clicked.connect(_do_speichern)
        layout.addWidget(btn_box)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        datum_str = datum_edit.date().toString("dd.MM.yyyy")

        # Word-Datei in Berichte-Ordner kopieren
        word_rel = ""
        if selected_file[0] and os.path.isfile(selected_file[0]):
            try:
                ziel_dir = self._monats_dir(datum_str)
                dateiname = os.path.basename(selected_file[0])
                ziel_pfad = ziel_dir / dateiname
                if ziel_pfad.exists():
                    try:
                        same = os.path.samefile(selected_file[0], str(ziel_pfad))
                    except OSError:
                        same = False
                    if not same:
                        stem, suffix = Path(dateiname).stem, Path(dateiname).suffix
                        i = 1
                        while ziel_pfad.exists():
                            ziel_pfad = ziel_dir / f"{stem}_{i}{suffix}"
                            i += 1
                        shutil.copy2(selected_file[0], ziel_pfad)
                else:
                    shutil.copy2(selected_file[0], ziel_pfad)
                word_rel = ziel_pfad.relative_to(self._berichte_basis_dir()).as_posix()
            except Exception as exc:
                QMessageBox.warning(self, "Hinweis",
                                    f"Word-Datei konnte nicht kopiert werden:\n{exc}")

        daten = {
            "flug":          bezeichnung.text().strip(),
            "typ":           typ_combo.currentText(),
            "bereich":       bereich_combo.currentText(),
            "datum":         datum_str,
            "ort":           ort_edit.text().strip(),
            "erstellt_von":  erstellt_von_edit.text().strip(),
            "offblock_plan": "",
            "offblock_ist":  "",
            "passagiere":    [],
            "personal":      [],
            "chronologie":   [],
            "ursache":       "",
            "ergebnis":      "",
            "word_datei":    word_rel,
        }

        from functions.vorkommnisse_db import speichern
        try:
            new_id = speichern(daten)
            self._aktualisiere_liste()
            QMessageBox.information(self, "Gespeichert",
                                    f"Eigener Eintrag #{new_id} wurde angelegt.")
        except Exception as exc:
            QMessageBox.critical(self, "Fehler beim Speichern", str(exc))

    # ── Formular-Helfer ────────────────────────────────────────────────────────

    def _formular_leeren(self):
        self._flug_edit.clear()
        self._typ_combo.setCurrentIndex(0)
        self._bereich_combo.setCurrentIndex(0)
        self._datum_edit.setDate(QDate.currentDate())
        self._ort_edit.clear()
        self._offblock_plan_chk.setChecked(False)
        self._offblock_plan_edit.setTime(QTime(0, 0))
        self._offblock_ist_chk.setChecked(False)
        self._offblock_ist_edit.setTime(QTime(0, 0))
        self._erstellt_von_edit.clear()
        self._pax_table.set_data([])
        self._personal_table.set_data([])
        self._chrono_table.set_data([])
        self._ursache_edit.clear()
        self._ergebnis_edit.clear()

    def _formular_befuellen(self, d: dict):
        self._flug_edit.setText(d.get("flug", ""))
        idx = self._typ_combo.findText(d.get("typ", ""))
        if idx >= 0:
            self._typ_combo.setCurrentIndex(idx)
        b_idx = self._bereich_combo.findText(d.get("bereich", ""))
        if b_idx >= 0:
            self._bereich_combo.setCurrentIndex(b_idx)
        qdate = QDate.fromString(d.get("datum", ""), "dd.MM.yyyy")
        if qdate.isValid():
            self._datum_edit.setDate(qdate)
        self._ort_edit.setText(d.get("ort", ""))
        def _pt(s: str) -> QTime:
            t = QTime.fromString(s.split(" ")[0], "HH:mm")
            return t if t.isValid() else QTime(0, 0)
        _plan_raw = d.get("offblock_plan", "")
        _ist_raw  = d.get("offblock_ist",  "")
        _plan_has = bool(_plan_raw and _plan_raw not in ("", "00:00", "00:00 Uhr"))
        _ist_has  = bool(_ist_raw  and _ist_raw  not in ("", "00:00", "00:00 Uhr"))
        self._offblock_plan_chk.setChecked(_plan_has)
        self._offblock_plan_edit.setTime(_pt(_plan_raw) if _plan_has else QTime(0, 0))
        self._offblock_ist_chk.setChecked(_ist_has)
        self._offblock_ist_edit.setTime(_pt(_ist_raw) if _ist_has else QTime(0, 0))
        self._erstellt_von_edit.setText(d.get("erstellt_von", ""))
        self._pax_table.set_data(d.get("passagiere", []))
        self._personal_table.set_data(d.get("personal", []))
        self._chrono_table.set_data(d.get("chronologie", []))
        self._ursache_edit.setPlainText(d.get("ursache", ""))
        self._ergebnis_edit.setPlainText(d.get("ergebnis", ""))

    def _sammle_daten(self) -> dict:
        if self._offblock_plan_chk.isChecked():
            plan = self._offblock_plan_edit.time().toString("HH:mm")
            plan_min = (
                self._offblock_plan_edit.time().hour() * 60
                + self._offblock_plan_edit.time().minute()
            )
            offblock_plan_str = f"{plan} Uhr"
        else:
            plan_min = 0
            offblock_plan_str = ""

        if self._offblock_ist_chk.isChecked():
            ist = self._offblock_ist_edit.time().toString("HH:mm")
            ist_min = (
                self._offblock_ist_edit.time().hour() * 60
                + self._offblock_ist_edit.time().minute()
            )
            if self._offblock_plan_chk.isChecked():
                versp = ist_min - plan_min
                if versp < 0:
                    versp += 1440
                offblock_ist_str = f"{ist} Uhr"
                if versp > 0:
                    offblock_ist_str += f" (+{versp} Min.)"
            else:
                offblock_ist_str = f"{ist} Uhr"
        else:
            offblock_ist_str = ""
        return {
            "flug":          self._flug_edit.text().strip(),
            "typ":           self._typ_combo.currentText(),
            "bereich":       self._bereich_combo.currentText(),
            "datum":         self._datum_edit.date().toString("dd.MM.yyyy"),
            "ort":           self._ort_edit.text().strip(),
            "offblock_plan": offblock_plan_str,
            "offblock_ist":  offblock_ist_str,
            "erstellt_von":  self._erstellt_von_edit.text().strip(),
            "passagiere":    self._pax_table.get_data(),
            "personal":      self._personal_table.get_data(),
            "chronologie":   self._chrono_table.get_data(),
            "ursache":       self._ursache_edit.toPlainText().strip(),
            "ergebnis":      self._ergebnis_edit.toPlainText().strip(),
        }

    # ── Word-Export ────────────────────────────────────────────────────────────

    def _berichte_basis_dir(self) -> Path:
        """Basisordner für alle Vorkommnis-Berichte."""
        return Path(BASE_DIR) / "Daten" / "Vorkommnis Berichte"

    def _vk_dateiname(self, flug: str, datum_kurz: str) -> str:
        """Erstellt einen sicheren Dateinamen für Vorkommnisberichte.
        Kürzt den Flugnamen wenn der Gesamtpfad das Windows MAX_PATH (260) überschreiten würde."""
        sicher_flug = flug.replace("/", "-").replace(" ", "_")
        suffix = f"_{datum_kurz}.docx" if datum_kurz else ".docx"
        prefix = "Vorkommnisbericht_"
        # Reservierte Länge: Basispfad + Separator + Monatsordner (max ~18 Zeichen) + Separator
        basis_len = len(str(self._berichte_basis_dir())) + 1 + 18 + 1
        max_flug = 259 - basis_len - len(prefix) - len(suffix)
        if max_flug < 10:
            max_flug = 10
        if len(sicher_flug) > max_flug:
            sicher_flug = sicher_flug[:max_flug]
        return f"{prefix}{sicher_flug}{suffix}"

    def _monats_dir(self, datum: str) -> Path:
        """Gibt den Monatsordner zurück und legt ihn bei Bedarf an.
        datum erwartet Format 'dd.MM.yyyy'."""
        try:
            d = datetime.strptime(datum, "%d.%m.%Y")
        except ValueError:
            d = datetime.today()
        # Ordnername z. B. "2026-04 April"
        monat_name = d.strftime("%B")  # lokalisierter Monatsname
        ordner = self._berichte_basis_dir() / f"{d.year}-{d.month:02d} {monat_name}"
        ordner.mkdir(parents=True, exist_ok=True)
        return ordner

    def _ordner_oeffnen(self):
        """Öffnet den Berichte-Basisordner im Windows Explorer."""
        basis = self._berichte_basis_dir()
        basis.mkdir(parents=True, exist_ok=True)
        os.startfile(str(basis))

    # ── E-Mail Entwurf ─────────────────────────────────────────────────────────
    def _ermittle_word_pfad(self, flug: str, datum: str) -> "str | None":
        """Sucht die passende Word-Datei zum Bericht anhand Flug und Datum.
        Gibt den Pfad zurück, oder None wenn keine gefunden wurde."""
        import glob
        basis = self._berichte_basis_dir()
        sicher_flug = flug.replace("/", "-").replace(" ", "_")
        # Datum in JJJJMMTT für den Dateinamen
        try:
            from datetime import datetime as _dt
            datum_kurz = _dt.strptime(datum, "%d.%m.%Y").strftime("%d-%m-%Y")
        except Exception:
            datum_kurz = ""
        # Exakter Dateiname wie beim Export (mit Datum)
        kandidat_name = self._vk_dateiname(flug, datum_kurz)
        kandidat_pfad = self._monats_dir(datum) / kandidat_name
        if kandidat_pfad.exists():
            return str(kandidat_pfad)
        # Breiteren Scan: irgendeine .docx mit dem Flugnamen
        alle = [p for p in glob.glob(str(basis / "**" / f"*{sicher_flug}*.docx"), recursive=True)
                if os.path.isfile(p)]  # OneDrive-Placeholder überspringen
        if alle:
            # Neueste nehmen
            return max(alle, key=os.path.getmtime)
        return None

    def _exportiere_word_automatisch(self, daten: dict) -> "str | None":
        """Erstellt die Word-Datei ohne Dialog und gibt den Pfad zurück."""
        from datetime import datetime as _dt
        try:
            datum_kurz = _dt.strptime(daten.get("datum", ""), "%d.%m.%Y").strftime("%d-%m-%Y")
        except Exception:
            datum_kurz = ""
        default_name = self._vk_dateiname(daten.get("flug", ""), datum_kurz)
        pfad = str(self._monats_dir(daten["datum"]) / default_name)
        try:
            self._erstelle_word(pfad, daten)
            return pfad
        except Exception as exc:
            QMessageBox.critical(self, "Fehler beim Word-Export", str(exc))
            return None

    def _email_entwurf_dialog(self):
        """Zeigt Dialog zum Erstellen einer E-Mail mit optionalem Anhang.
        Die passende Word-Datei wird automatisch gefunden oder exportiert."""
        try:
            self._email_entwurf_dialog_intern()
        except Exception as _exc:
            import traceback
            QMessageBox.critical(self, "Fehler in E-Mail-Dialog", traceback.format_exc())

    def _email_entwurf_dialog_intern(self):
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
            QTextEdit, QListWidget, QListWidgetItem, QDialogButtonBox,
            QAbstractItemView, QDateEdit,
        )
        from PySide6.QtCore import Qt, QDate
        import glob

        daten = self._sammle_daten()
        flug    = daten.get("flug", "").strip()
        datum   = daten.get("datum", "").strip()
        _ref_id = flug or daten.get("bereich", "")  # Fallback auf Bereich

        # ── Passende Word-Datei automatisch ermitteln oder erstellen ──────────
        vorselektierter_pfad = self._ermittle_word_pfad(flug, datum)
        if not vorselektierter_pfad and _ref_id:
            antwort = QMessageBox.question(
                self, "Kein Bericht gefunden",
                f"Für den Bericht »{_ref_id}« wurde keine Word-Datei gefunden.\n"
                "Soll die Datei jetzt automatisch erstellt und angehängt werden?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if antwort == QMessageBox.StandardButton.Yes:
                vorselektierter_pfad = self._exportiere_word_automatisch(daten)

        # Datum aus Formular parsen (dd.mm.yyyy), sonst heute
        try:
            tag, mon, jahr = datum.split(".")
            initial_date = QDate(int(jahr), int(mon), int(tag))
        except Exception:
            initial_date = QDate.currentDate()

        dlg = QDialog(self)
        dlg.setWindowTitle("E-Mail erstellen")
        dlg.setMinimumWidth(560)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 12)

        form = QFormLayout()
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        # Datum (auswählbar)
        datum_edit = QDateEdit()
        datum_edit.setCalendarPopup(True)
        datum_edit.setDate(initial_date)
        datum_edit.setDisplayFormat("dd.MM.yyyy")
        form.addRow("Datum:", datum_edit)

        # Absender-Name
        name_edit = QLineEdit()
        name_edit.setPlaceholderText("Vor- und Nachname …")
        form.addRow("Absender (Grüße):", name_edit)

        # Betreff
        betreff_edit = QLineEdit()
        form.addRow("Betreff:", betreff_edit)

        layout.addLayout(form)

        def _aktualisiere():
            d = datum_edit.date().toString("dd.MM.yyyy")
            _bereich = daten.get("bereich", "")
            if _bereich == "Flüge":
                _ref_text = f"Flug {flug}" if flug else ""
            elif _bereich == "Fahrzeuge":
                _ref_text = f"Fahrzeug {flug}" if flug else ""
            elif _bereich:
                _ref_text = f"{_bereich} – {flug}" if flug else _bereich
            else:
                _ref_text = f"Flug {flug}" if flug else ""
            betreff_edit.setText(
                f"Vorkommnisbericht – {_ref_text} – {d}" if _ref_text
                else f"Vorkommnisbericht – {d}"
            )
            name = name_edit.text().strip()
            gruss = f"\n{name}" if name else "\nDRK Erste-Hilfe-Station Flughafen Köln/Bonn"
            text_edit.setPlainText(
                "Hallo Lars,\n\n"
                "anbei der Vorkommnisbericht"
                + (f" zu {_ref_text}" if _ref_text else "")
                + f" vom {d}"
                + ".\n\nBei Fragen meld dich gerne.\n\n"
                "Viele Grüße"
                + gruss
            )

        datum_edit.dateChanged.connect(lambda _: _aktualisiere())
        name_edit.textChanged.connect(lambda _: _aktualisiere())

        # Mailtext
        layout.addWidget(QLabel("<b>Mailtext:</b>"))
        text_edit = QTextEdit()
        text_edit.setFixedHeight(140)
        layout.addWidget(text_edit)
        _aktualisiere()  # Initialtext setzen

        # Anhang-Auswahl aus Berichte-Ordner
        if vorselektierter_pfad:
            anhang_lbl_text = "<b>Word-Anhang</b> (automatisch ausgewählt, optional ändern):"
        else:
            anhang_lbl_text = "<b>Word-Anhang auswählen</b> (optional, aus Berichte-Ordner):"
        layout.addWidget(QLabel(anhang_lbl_text))
        anhang_liste = QListWidget()
        anhang_liste.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        anhang_liste.setFixedHeight(140)

        basis = self._berichte_basis_dir()
        docx_dateien = sorted(
            [p for p in glob.glob(str(basis / "**" / "*.docx"), recursive=True)
             if os.path.isfile(p)],  # OneDrive-Placeholder überspringen
            key=os.path.getmtime, reverse=True
        )
        vorsel_item = None
        for pfad in docx_dateien:
            item = QListWidgetItem(os.path.basename(pfad))
            item.setData(Qt.ItemDataRole.UserRole, pfad)
            anhang_liste.addItem(item)
            if vorselektierter_pfad and os.path.normcase(pfad) == os.path.normcase(vorselektierter_pfad):
                vorsel_item = item

        if not docx_dateien:
            anhang_liste.addItem("(Keine Berichte im Berichte-Ordner gefunden)")
            anhang_liste.setEnabled(False)
        elif vorsel_item:
            anhang_liste.setCurrentItem(vorsel_item)
            anhang_liste.scrollToItem(vorsel_item)

        layout.addWidget(anhang_liste)

        # Buttons
        btn_box = QDialogButtonBox()
        btn_senden = btn_box.addButton("✉️  E-Mail öffnen", QDialogButtonBox.ButtonRole.AcceptRole)
        btn_ab     = btn_box.addButton("Abbrechen",         QDialogButtonBox.ButtonRole.RejectRole)
        btn_box.rejected.connect(dlg.reject)
        layout.addWidget(btn_box)

        def _erstelle():
            betreff  = betreff_edit.text().strip()
            mailtext = text_edit.toPlainText()
            anhang_pfad = None
            sel = anhang_liste.selectedItems()
            if sel:
                pfad_data = sel[0].data(Qt.ItemDataRole.UserRole)
                if pfad_data and os.path.isfile(pfad_data):
                    anhang_pfad = pfad_data

            # Vorfall automatisch speichern, falls noch nicht gespeichert
            if self._current_id is None:
                _daten = self._sammle_daten()
                if _daten.get("flug") or _daten.get("bereich"):
                    try:
                        from functions.vorkommnisse_db import speichern as _vk_speichern
                        self._current_id = _vk_speichern(_daten)
                        self._aktualisiere_status()
                        self._aktualisiere_liste()
                    except Exception:
                        pass

            dlg.accept()
            self._erstelle_outlook_entwurf(betreff, mailtext, anhang_pfad)

        btn_senden.clicked.connect(_erstelle)
        dlg.exec()

    def _erstelle_outlook_entwurf(self, betreff: str, mailtext: str, anhang_pfad: str | None):
        """Öffnet eine neue E-Mail direkt in Outlook zur Bearbeitung und zum Versand."""
        try:
            import win32com.client
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Display()               # Outlook lädt Standardsignatur
            signature = mail.HTMLBody
            mail.Subject = betreff
            html_text = mailtext.replace("\n", "<br>")
            mail.HTMLBody = (
                f"<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"
                f"<div>{html_text}</div>"
                f"</body></html>"
            ) + signature
            if anhang_pfad and os.path.isfile(anhang_pfad):
                mail.Attachments.Add(os.path.abspath(anhang_pfad))
        except Exception as exc:
            QMessageBox.critical(self, "Outlook-Fehler", str(exc))

    def _export_word(self):
        daten = self._sammle_daten()
        if not daten["flug"] and not daten["bereich"]:
            QMessageBox.warning(self, "Fehler", "Bitte Bezeichnung / Flugnummer eingeben.")
            return

        try:
            from datetime import datetime as _dt
            datum_kurz = _dt.strptime(daten["datum"], "%d.%m.%Y").strftime("%d-%m-%Y")
        except Exception:
            datum_kurz = ""
        default_name = self._vk_dateiname(daten['flug'], datum_kurz)
        # Monatsordner basierend auf dem Datum des Vorkommnisses
        speicher_dir = self._monats_dir(daten["datum"])
        pfad, _ = QFileDialog.getSaveFileName(
            self, "Word-Dokument speichern",
            str(speicher_dir / default_name),
            "Word-Dokument (*.docx)",
        )
        if not pfad:
            return
        try:
            self._erstelle_word(pfad, daten)
            QMessageBox.information(
                self, "Exportiert",
                f"Vorkommnisbericht gespeichert:\n{pfad}",
            )
            # Word-Datei automatisch mit dem DB-Eintrag verknüpfen
            if self._current_id is not None:
                try:
                    basis = self._berichte_basis_dir()
                    pfad_obj = Path(pfad)
                    if pfad_obj.is_relative_to(basis):
                        rel = pfad_obj.relative_to(basis).as_posix()
                        from functions.vorkommnisse_db import setze_word_datei
                        setze_word_datei(self._current_id, rel)
                        self._aktualisiere_liste()
                except Exception:
                    pass
            os.startfile(pfad)
        except ImportError:
            QMessageBox.critical(
                self, "Fehler",
                "python-docx nicht installiert.\nBitte ausführen: pip install python-docx",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export-Fehler", str(exc))

    def _erstelle_word(self, pfad: str, d: dict):
        """Erstellt das Word-Dokument mit Kopf-/Fußzeile und allen Abschnitten."""
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        doc = Document()

        # ── Seitenränder ──────────────────────────────────────────────────────
        for section in doc.sections:
            section.top_margin    = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        # ── Kopfzeile ─────────────────────────────────────────────────────────
        self._baue_kopfzeile(doc)

        # ── Titel ─────────────────────────────────────────────────────────────
        # Bereichs-abhängiger Bezug (kein hartes "Flug" mehr)
        _bereich = d.get("bereich", "")
        _ref     = d.get("flug", "")
        if _bereich == "Flüge":
            _bezug = f"Flug {_ref}" if _ref else "Flug"
        elif _bereich == "Fahrzeuge":
            _bezug = f"Fahrzeug {_ref}" if _ref else "Fahrzeug"
        elif _bereich == "Personal":
            _bezug = f"Personal – {_ref}" if _ref else "Personal"
        elif _bereich == "Sachschäden":
            _bezug = f"Sachschaden – {_ref}" if _ref else "Sachschaden"
        elif _bereich == "Infrastruktur":
            _bezug = f"Infrastruktur – {_ref}" if _ref else "Infrastruktur"
        else:
            _bezug = _ref if _ref else (_bereich or "Vorkommnis")

        titel = doc.add_heading(
            f"Vorkommnis Bericht – {d['typ']} – {_bezug}", level=1
        )
        titel.runs[0].font.color.rgb = RGBColor(0x15, 0x65, 0xa8)

        # ── Grunddaten-Tabelle ────────────────────────────────────────────────
        _ref_label = _BEREICH_LABEL.get(_bereich, "Bezeichnung / Flugnummer:").rstrip(":")
        grund_daten = [
            ("Vorkommnis-Typ",  d["typ"]),
            ("Bereich",         _bereich or "–"),
            (_ref_label,        d["flug"]),
            ("Datum",           d["datum"]),
            ("Ort",             d["ort"]),
        ]
        if d.get("offblock_plan"):
            grund_daten.append(("Geplanter Offblock", d["offblock_plan"]))
        if d.get("offblock_ist"):
            grund_daten.append(("Tatsächlicher Offblock", d["offblock_ist"]))

        tbl_grund = doc.add_table(rows=len(grund_daten), cols=2)
        tbl_grund.style = "Table Grid"
        for i, (key, val) in enumerate(grund_daten):
            tbl_grund.rows[i].cells[0].text = key
            tbl_grund.rows[i].cells[1].text = val
            tbl_grund.rows[i].cells[0].paragraphs[0].runs[0].bold = True
        doc.add_paragraph()

        # ── 1. Betroffene Personen ────────────────────────────────────────────
        doc.add_heading("1. Betroffene Personen", level=2)
        if d["passagiere"]:
            tbl_pax = doc.add_table(rows=1, cols=4)
            tbl_pax.style = "Table Grid"
            for i, hdr in enumerate(["Person", "Typ", "Kategorie", "Anmerkung"]):
                cell = tbl_pax.rows[0].cells[i]
                cell.text = hdr
                cell.paragraphs[0].runs[0].bold = True
            for row_data in d["passagiere"]:
                row = tbl_pax.add_row()
                for i, val in enumerate(row_data[:4]):
                    row.cells[i].text = val
        else:
            doc.add_paragraph("(keine Personen erfasst)")
        doc.add_paragraph()

        # ── 2. Personal ───────────────────────────────────────────────────────
        doc.add_heading("2. Eingeteiltes Personal", level=2)
        if d["personal"]:
            tbl_pers = doc.add_table(rows=1, cols=3)
            tbl_pers.style = "Table Grid"
            for i, hdr in enumerate(["Mitarbeiter", "Funktion/Rolle", "Anmerkung"]):
                cell = tbl_pers.rows[0].cells[i]
                cell.text = hdr
                cell.paragraphs[0].runs[0].bold = True
            for row_data in d["personal"]:
                row = tbl_pers.add_row()
                for i, val in enumerate(row_data[:3]):
                    row.cells[i].text = val
        else:
            doc.add_paragraph("(kein Personal erfasst)")
        doc.add_paragraph()

        # ── 3. Chronologischer Ablauf ─────────────────────────────────────────
        doc.add_heading("3. Chronologischer Ablauf", level=2)
        if d["chronologie"]:
            tbl_chr = doc.add_table(rows=1, cols=2)
            tbl_chr.style = "Table Grid"
            for i, hdr in enumerate(["Uhrzeit", "Ereignis"]):
                cell = tbl_chr.rows[0].cells[i]
                cell.text = hdr
                cell.paragraphs[0].runs[0].bold = True
            for row_data in d["chronologie"]:
                row = tbl_chr.add_row()
                for i, val in enumerate(row_data[:2]):
                    row.cells[i].text = val
        else:
            doc.add_paragraph("(kein Ablauf erfasst)")
        doc.add_paragraph()

        # ── 4. Ursachenanalyse ────────────────────────────────────────────────
        doc.add_heading("4. Ursachenanalyse", level=2)
        doc.add_paragraph(d["ursache"] or "(keine Angabe)")
        doc.add_paragraph()

        # ── 5. Ergebnis ───────────────────────────────────────────────────────
        doc.add_heading("5. Ergebnis", level=2)
        doc.add_paragraph(d["ergebnis"] or "(keine Angabe)")
        doc.add_paragraph()

        # ── Unterschrift ──────────────────────────────────────────────────────
        doc.add_paragraph()
        p_sign = doc.add_paragraph(f"Erstellt von: {d['erstellt_von'] or '________________'}")
        p_sign.add_run(f"        Datum: {d['datum']}")

        # ── Fußzeile ──────────────────────────────────────────────────────────
        self._baue_fusszeile(doc)

        doc.save(pfad)

    def _baue_kopfzeile(self, doc):
        """Firmenlogo + Anschrift in der Kopfzeile."""
        from docx.shared import Pt, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        for section in doc.sections:
            header = section.header
            # Erste Zeile: Firmenname fett
            p1 = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            p1.clear()
            run = p1.add_run("Deutsches Rotes Kreuz – Erste-Hilfe-Station Flughafen Köln/Bonn")
            run.bold = True
            run.font.size = Pt(11)
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Zweite Zeile: Adresse
            p2 = header.add_paragraph()
            run2 = p2.add_run("Flughafen Köln/Bonn (CGN)  |  Kennedystraße  |  51147 Köln")
            run2.font.size = Pt(9)
            p2.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Logo versuchen einzubinden (optional, kein Fehler wenn nicht vorhanden)
            _base = Path(BASE_DIR) if not isinstance(BASE_DIR, Path) else BASE_DIR
            logo_pfad = _base / "json" / "logo.png"
            if not logo_pfad.exists():
                logo_pfad = _base / "gui" / "logo.png"
            if logo_pfad.exists():
                try:
                    p_logo = header.add_paragraph()
                    run_logo = p_logo.add_run()
                    run_logo.add_picture(str(logo_pfad), width=Cm(3))
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                except Exception:
                    pass  # Logo-Einbindung ist optional

    def _baue_fusszeile(self, doc):
        """Seitenzahl und Hinweis in der Fußzeile."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        for section in doc.sections:
            footer = section.footer
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p.clear()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Text links
            run_links = p.add_run("Vertraulich – DRK Erste-Hilfe-Station CGN    |    Seite ")
            run_links.font.size = Pt(8)

            # Automatische Seitenzahl (XML-Feld)
            fldChar1 = OxmlElement("w:fldChar")
            fldChar1.set(qn("w:fldCharType"), "begin")
            instrText = OxmlElement("w:instrText")
            instrText.text = "PAGE"
            fldChar2 = OxmlElement("w:fldChar")
            fldChar2.set(qn("w:fldCharType"), "end")
            run_page = p.add_run()
            run_page._r.append(fldChar1)
            run_page._r.append(instrText)
            run_page._r.append(fldChar2)
            run_page.font.size = Pt(8)
