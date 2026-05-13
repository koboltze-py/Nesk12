"""
Handys-Widget
Verwaltung der Diensthandys der DRK-EHS Flughafen Köln/Bonn
3 Tabs: Übersicht | Details / Bearbeiten | Historie
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QLineEdit, QComboBox, QDateEdit, QTextEdit,
    QFormLayout, QMessageBox, QSizePolicy, QAbstractItemView,
    QSplitter,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont, QColor

from config import FIORI_BLUE, FIORI_TEXT, FIORI_WHITE, FIORI_BORDER, FIORI_SIDEBAR_BG
from functions.handys_db import (
    lade_alle_handys, lade_handy, erstelle_handy, aktualisiere_handy,
    loesche_handy, lade_historie, ZUSTAND_OPTIONEN, ZUSTAND_MIT_DEFEKT_DETAIL,
)

ZUSTAND_FARBEN = {
    "Aktiv":         {"fg": "#276221", "bg": "#C6EFCE"},
    "Defekt":        {"fg": "#9C0006", "bg": "#FFC7CE"},
    "Außer Betrieb": {"fg": "#9C6500", "bg": "#FFEB9C"},
    "Reserve":       {"fg": "#404040", "bg": "#D9D9D9"},
    "Verloren":      {"fg": "#5c007a", "bg": "#e8c9f5"},
}

_UEBERSICHT_SPALTEN = [
    ("ID",                  "id",                   50,  False),
    ("Inventar-Nr.",        "inventarnummer",        110, True),
    ("Hersteller",          "hersteller",             90, True),
    ("Modell",              "modell",                110, True),
    ("Rufnummer",           "rufnummer",             110, True),
    ("Kartennummer",        "kartennummer",          120, True),
    ("PIN",                 "pin",                    60, True),
    ("PUK",                 "puk",                    90, True),
    ("PUK 2",               "puk2",                   90, True),
    ("Standort / Ausgabe",  "standort",              130, True),
    ("Zustand",             "zustand",                90, True),
    ("Defektbeschreibung",  "defekt_beschreibung",   180, True),
    ("Festgestellt am",     "defekt_datum",          100, True),
    ("Festgestellt von",    "defekt_gemeldet_von",   120, True),
    ("Anschaffung",         "anschaffungsdatum",      90, True),
    ("Zuletzt geändert",    "geaendert_am",          130, True),
]

_HISTORIE_SPALTEN = [
    ("Datum / Uhrzeit",  "geaendert_am",    150),
    ("Inventar-Nr.",     "inventarnummer",  110),
    ("Feld",             "feld",             120),
    ("Alter Wert",       "alter_wert",      150),
    ("Neuer Wert",       "neuer_wert",      150),
    ("Benutzer",         "benutzer",        100),
]


def _btn(text: str, color: str = FIORI_BLUE) -> QPushButton:
    b = QPushButton(text)
    b.setFixedHeight(32)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    b.setStyleSheet(f"""
        QPushButton {{
            background-color: {color};
            color: white;
            border: none;
            border-radius: 4px;
            padding: 0 12px;
            font-size: 11px;
            font-weight: bold;
        }}
        QPushButton:hover {{ background-color: #0856a8; }}
        QPushButton:disabled {{ background-color: #b0b8c1; }}
    """)
    return b


class HandysWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._aktueller_handy_id: int | None = None
        self._handys: list[dict] = []
        self._build_ui()
        self.refresh()

    # ─── UI-Aufbau ────────────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)

        # Seitentitel
        titel = QLabel("📱 Diensthandys – DRK EHS Flughafen Köln/Bonn")
        titel.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        titel.setStyleSheet(f"color: {FIORI_TEXT};")
        layout.addWidget(titel)

        # Tab-Widget
        self._tabs = QTabWidget()
        self._tabs.setDocumentMode(False)
        self._tabs.setStyleSheet("""
            QTabWidget::pane { border: none; background: #f8f9fa; }
            QTabBar::tab {
                padding: 8px 22px; font-size: 13px; font-family: 'Segoe UI';
                color: #666; background: #e8ecf0;
                border-bottom: 2px solid transparent;
                border-radius: 4px 4px 0 0; margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #f8f9fa; color: #1565a8;
                font-weight: bold; border-bottom: 2px solid #1565a8;
            }
            QTabBar::tab:hover:!selected { background: #dde4ec; color: #1565a8; }
        """)
        self._tabs.addTab(self._build_uebersicht_tab(), "📋  Übersicht")
        self._tabs.addTab(self._build_details_tab(),    "✏  Details / Bearbeiten")
        self._tabs.addTab(self._build_historie_tab(),   "🕐  Historie")
        self._tabs.addTab(self._build_berichte_tab(),   "📂  Berichte")
        layout.addWidget(self._tabs)

    # ── Tab 1: Übersicht ──────────────────────────────────────────────────────

    def _build_uebersicht_tab(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        # Button-Leiste
        btn_leiste = QHBoxLayout()
        self._btn_neu       = _btn("+ Neues Handy", "#107e3e")
        self._btn_bearbeiten = _btn("✎ Bearbeiten")
        self._btn_loeschen  = _btn("✕ Löschen", "#bb0000")
        self._btn_bericht   = _btn("📋 Schadens-/Verlustbericht", "#e07b00")
        self._btn_bericht.setEnabled(False)
        self._btn_refresh   = _btn("⟳ Aktualisieren", "#5c6bc0")
        self._btn_excel     = _btn("📤 Excel exportieren", "#0a6ed1")
        self._btn_email     = _btn("📧 Geräteübersicht per E-Mail", "#0a6ed1")

        for b in [self._btn_neu, self._btn_bearbeiten, self._btn_loeschen,
                  self._btn_bericht, self._btn_refresh, self._btn_excel, self._btn_email]:
            btn_leiste.addWidget(b)
        btn_leiste.addStretch()
        vl.addLayout(btn_leiste)

        # Tabelle
        self._tabelle = QTableWidget()
        self._tabelle.setColumnCount(len(_UEBERSICHT_SPALTEN))
        self._tabelle.setHorizontalHeaderLabels([s[0] for s in _UEBERSICHT_SPALTEN])
        self._tabelle.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._tabelle.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._tabelle.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tabelle.setAlternatingRowColors(True)
        self._tabelle.verticalHeader().setVisible(False)
        self._tabelle.setSortingEnabled(True)

        # Spaltenbreiten
        hh = self._tabelle.horizontalHeader()
        for col_idx, (_, _, breite, sichtbar) in enumerate(_UEBERSICHT_SPALTEN):
            if not sichtbar:
                self._tabelle.setColumnHidden(col_idx, True)
            else:
                self._tabelle.setColumnWidth(col_idx, breite)
        hh.setStretchLastSection(True)

        self._tabelle.setStyleSheet("""
            QTableWidget { border: 1px solid #d9d9d9; gridline-color: #ebebeb; }
            QTableWidget::item:selected { background-color: #0a6ed1; color: white; }
            QHeaderView::section { background-color: #354a5e; color: white;
                                   padding: 6px; font-weight: bold; font-size: 10px; }
        """)

        vl.addWidget(self._tabelle)

        # Verbindungen
        self._btn_neu.clicked.connect(self._neues_handy)
        self._btn_bearbeiten.clicked.connect(self._bearbeiten)
        self._btn_loeschen.clicked.connect(self._loeschen)
        self._btn_bericht.clicked.connect(self._schadensbericht_erstellen)
        self._btn_refresh.clicked.connect(self.refresh)
        self._btn_excel.clicked.connect(self._excel_export)
        self._btn_email.clicked.connect(self._email_senden)
        self._tabelle.doubleClicked.connect(self._bearbeiten)
        self._tabelle.selectionModel().selectionChanged.connect(self._auswahl_geaendert)

        return w

    # ── Tab 2: Details / Bearbeiten ───────────────────────────────────────────

    def _build_details_tab(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(16, 16, 16, 16)
        vl.setSpacing(12)

        # Formular
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setSpacing(8)

        self._f_inventar     = QLineEdit()
        self._f_hersteller   = QLineEdit()
        self._f_modell       = QLineEdit()
        self._f_rufnummer    = QLineEdit()
        self._f_sim          = QLineEdit()
        self._f_standort     = QLineEdit()
        self._f_zustand      = QComboBox()
        self._f_zustand.addItems(ZUSTAND_OPTIONEN)
        self._f_defekt       = QLineEdit()
        # Felder für Defekt-/Verloren-Details (nur bei relevantem Zustand sichtbar)
        self._f_defekt_datum         = QDateEdit()
        self._f_defekt_datum.setCalendarPopup(True)
        self._f_defekt_datum.setDisplayFormat("dd.MM.yyyy")
        self._f_defekt_datum.setDate(QDate.currentDate())
        self._f_defekt_gemeldet_von  = QLineEdit()
        self._f_anschaffung  = QDateEdit()
        self._f_anschaffung.setCalendarPopup(True)
        self._f_anschaffung.setSpecialValueText("–")
        self._f_anschaffung.setDate(QDate.currentDate())
        self._f_anschaffung.setDisplayFormat("dd.MM.yyyy")
        self._f_notizen      = QTextEdit()
        self._f_notizen.setFixedHeight(80)
        # SIM-Karte / PIN / PUK
        self._f_kartennummer = QLineEdit()
        self._f_pin          = QLineEdit()
        self._f_pin2         = QLineEdit()
        self._f_puk          = QLineEdit()
        self._f_puk2         = QLineEdit()

        _pflicht = " <span style='color:red'>*</span>"
        form.addRow(f"Inventarnummer{_pflicht}:", self._f_inventar)
        form.addRow("Hersteller:",               self._f_hersteller)
        form.addRow("Modell:",                   self._f_modell)
        form.addRow("Rufnummer:",                self._f_rufnummer)
        form.addRow("SIM-Nummer:",               self._f_sim)
        form.addRow("Standort / Ausgabe:",       self._f_standort)
        form.addRow(f"Zustand{_pflicht}:",       self._f_zustand)
        form.addRow("Defektbeschreibung:",       self._f_defekt)
        self._lbl_defekt_datum       = QLabel("Festgestellt am:")
        self._lbl_defekt_gemeldet    = QLabel("Festgestellt von:")
        form.addRow(self._lbl_defekt_datum,      self._f_defekt_datum)
        form.addRow(self._lbl_defekt_gemeldet,   self._f_defekt_gemeldet_von)
        form.addRow("Anschaffungsdatum:",        self._f_anschaffung)
        form.addRow("Notizen:",                  self._f_notizen)

        # Trennlinie SIM-Karte
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #c0c8d0; margin: 4px 0;")
        form.addRow(sep)
        lbl_sim_titel = QLabel("<b>SIM-Karte / PIN / PUK</b>")
        lbl_sim_titel.setStyleSheet("color: #1565a8; font-size: 11px;")
        form.addRow(lbl_sim_titel)
        form.addRow("Kartennummer:",             self._f_kartennummer)
        form.addRow("PIN:",                      self._f_pin)
        form.addRow("PIN 2:",                    self._f_pin2)
        form.addRow("PUK:",                      self._f_puk)
        form.addRow("PUK 2:",                    self._f_puk2)

        vl.addLayout(form)

        # Hinweis
        hinweis = QLabel("* = Pflichtfeld")
        hinweis.setStyleSheet("color: #888; font-size: 10px;")
        vl.addWidget(hinweis)

        # Buttons
        btn_row = QHBoxLayout()
        self._btn_speichern  = _btn("💾 Speichern", "#107e3e")
        self._btn_abbrechen  = _btn("✕ Abbrechen", "#888888")
        btn_row.addWidget(self._btn_speichern)
        btn_row.addWidget(self._btn_abbrechen)
        btn_row.addStretch()
        vl.addLayout(btn_row)
        vl.addStretch()

        # Zustand → Defektbeschreibung aktivieren/deaktivieren
        self._f_zustand.currentTextChanged.connect(self._zustand_geaendert)
        self._zustand_geaendert(self._f_zustand.currentText())

        self._btn_speichern.clicked.connect(self._speichern)
        self._btn_abbrechen.clicked.connect(self._abbrechen)

        # Formular-Stylesheet
        for feld in [self._f_inventar, self._f_hersteller, self._f_modell,
                     self._f_rufnummer, self._f_sim, self._f_standort,
                     self._f_defekt, self._f_defekt_gemeldet_von,
                     self._f_kartennummer, self._f_pin, self._f_pin2,
                     self._f_puk, self._f_puk2]:
            feld.setFixedHeight(30)
            feld.setStyleSheet(
                "border: 1px solid #d9d9d9; border-radius: 4px; padding: 0 6px;"
            )

        return w

    # ── Tab 3: Historie ───────────────────────────────────────────────────────

    def _build_historie_tab(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        # Filter-Leiste
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Gerät:"))
        self._hist_filter_geraet = QComboBox()
        self._hist_filter_geraet.setFixedWidth(160)
        self._hist_filter_geraet.addItem("Alle Geräte", None)
        filter_row.addWidget(self._hist_filter_geraet)
        filter_row.addWidget(QLabel("Zeitraum:"))
        self._hist_filter_tage = QComboBox()
        self._hist_filter_tage.addItems(["30 Tage", "90 Tage", "180 Tage", "Gesamt"])
        self._hist_filter_tage.setCurrentIndex(1)
        self._hist_filter_tage.setFixedWidth(100)
        filter_row.addWidget(self._hist_filter_tage)
        btn_reload = _btn("⟳ Laden", "#5c6bc0")
        btn_reload.setFixedWidth(80)
        filter_row.addWidget(btn_reload)
        filter_row.addStretch()
        vl.addLayout(filter_row)

        # Tabelle
        self._hist_tabelle = QTableWidget()
        self._hist_tabelle.setColumnCount(len(_HISTORIE_SPALTEN))
        self._hist_tabelle.setHorizontalHeaderLabels([s[0] for s in _HISTORIE_SPALTEN])
        self._hist_tabelle.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._hist_tabelle.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._hist_tabelle.setAlternatingRowColors(True)
        self._hist_tabelle.verticalHeader().setVisible(False)
        self._hist_tabelle.setSortingEnabled(True)
        for col_idx, (_, _, breite) in enumerate(_HISTORIE_SPALTEN):
            self._hist_tabelle.setColumnWidth(col_idx, breite)
        self._hist_tabelle.horizontalHeader().setStretchLastSection(True)
        self._hist_tabelle.setStyleSheet("""
            QTableWidget { border: 1px solid #d9d9d9; gridline-color: #ebebeb; }
            QTableWidget::item:selected { background-color: #0a6ed1; color: white; }
            QHeaderView::section { background-color: #354a5e; color: white;
                                   padding: 6px; font-weight: bold; font-size: 10px; }
        """)
        vl.addWidget(self._hist_tabelle)

        btn_reload.clicked.connect(self._lade_historie)
        self._hist_filter_geraet.currentIndexChanged.connect(self._lade_historie)
        self._hist_filter_tage.currentIndexChanged.connect(self._lade_historie)

        return w

    # ── Tab 4: Berichte ─────────────────────────────────────────────────

    def _build_berichte_tab(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(12, 12, 12, 12)
        vl.setSpacing(8)

        # Filter-Leiste
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Gerät:"))
        self._ber_filter_geraet = QComboBox()
        self._ber_filter_geraet.setFixedWidth(180)
        self._ber_filter_geraet.addItem("Alle Geräte", None)
        filter_row.addWidget(self._ber_filter_geraet)
        filter_row.addStretch()
        btn_ber_reload = _btn("⟳ Laden", "#5c6bc0")
        btn_ber_reload.setFixedWidth(80)
        btn_ber_oeffnen = _btn("📄 Öffnen", "#0a6ed1")
        btn_ber_oeffnen.setFixedWidth(100)
        self._btn_ber_email = _btn("📧 Per E-Mail senden", "#0a6ed1")
        self._btn_ber_email.setFixedWidth(160)
        btn_ber_ordner = _btn("📂 Ordner öffnen", "#5c6bc0")
        btn_ber_ordner.setFixedWidth(130)
        filter_row.addWidget(btn_ber_reload)
        filter_row.addWidget(btn_ber_oeffnen)
        filter_row.addWidget(self._btn_ber_email)
        filter_row.addWidget(btn_ber_ordner)
        vl.addLayout(filter_row)

        # Tabelle
        self._ber_tabelle = QTableWidget()
        self._ber_tabelle.setColumnCount(4)
        self._ber_tabelle.setHorizontalHeaderLabels(["Gerät", "Typ", "Erstellt am", "Dateiname"])
        self._ber_tabelle.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._ber_tabelle.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._ber_tabelle.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._ber_tabelle.setAlternatingRowColors(True)
        self._ber_tabelle.verticalHeader().setVisible(False)
        self._ber_tabelle.setSortingEnabled(True)
        self._ber_tabelle.setColumnWidth(0, 110)
        self._ber_tabelle.setColumnWidth(1, 120)
        self._ber_tabelle.setColumnWidth(2, 140)
        self._ber_tabelle.horizontalHeader().setStretchLastSection(True)
        self._ber_tabelle.setStyleSheet("""
            QTableWidget { border: 1px solid #d9d9d9; gridline-color: #ebebeb; }
            QTableWidget::item:selected { background-color: #0a6ed1; color: white; }
            QHeaderView::section { background-color: #354a5e; color: white;
                                   padding: 6px; font-weight: bold; font-size: 10px; }
        """)
        vl.addWidget(self._ber_tabelle)

        # Pfad-Info
        self._ber_pfad_lbl = QLabel()
        self._ber_pfad_lbl.setStyleSheet("color: #888; font-size: 10px;")
        vl.addWidget(self._ber_pfad_lbl)

        # Verbindungen
        btn_ber_reload.clicked.connect(self._lade_berichte)
        self._ber_filter_geraet.currentIndexChanged.connect(self._lade_berichte)
        btn_ber_oeffnen.clicked.connect(self._bericht_oeffnen)
        self._btn_ber_email.clicked.connect(self._bericht_per_email_senden)
        btn_ber_ordner.clicked.connect(self._berichte_ordner_oeffnen)
        self._ber_tabelle.doubleClicked.connect(self._bericht_oeffnen)
        self._ber_tabelle.selectionModel().selectionChanged.connect(
            lambda: self._btn_ber_email.setEnabled(
                len(self._ber_tabelle.selectedItems()) > 0
            )
        )
        self._btn_ber_email.setEnabled(False)

        return w

    # ─── Daten laden / anzeigen ───────────────────────────────────────────────

    def refresh(self):
        """Lädt alle Handys neu und füllt die Tabelle."""
        try:
            self._handys = lade_alle_handys()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Daten konnten nicht geladen werden:\n{e}")
            return

        self._tabelle.setSortingEnabled(False)
        self._tabelle.setRowCount(len(self._handys))

        for row_idx, handy in enumerate(self._handys):
            for col_idx, (_, feld, _, _) in enumerate(_UEBERSICHT_SPALTEN):
                wert = str(handy.get(feld, "") or "")
                item = QTableWidgetItem(wert)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

                if feld == "zustand":
                    farben = ZUSTAND_FARBEN.get(wert, {})
                    if farben:
                        item.setBackground(QColor(farben["bg"]))
                        item.setForeground(QColor(farben["fg"]))
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)

                self._tabelle.setItem(row_idx, col_idx, item)

        self._tabelle.setSortingEnabled(True)
        self._aktualisiere_geraet_filter()
        self._lade_historie()
        self._lade_berichte()

    def _aktualisiere_geraet_filter(self):
        # Historien-Filter
        self._hist_filter_geraet.blockSignals(True)
        aktuell = self._hist_filter_geraet.currentData()
        self._hist_filter_geraet.clear()
        self._hist_filter_geraet.addItem("Alle Geräte", None)
        for h in self._handys:
            self._hist_filter_geraet.addItem(h["inventarnummer"], h["id"])
        if aktuell is not None:
            for i in range(self._hist_filter_geraet.count()):
                if self._hist_filter_geraet.itemData(i) == aktuell:
                    self._hist_filter_geraet.setCurrentIndex(i)
                    break
        self._hist_filter_geraet.blockSignals(False)
        # Berichte-Filter
        self._ber_filter_geraet.blockSignals(True)
        aktuell_ber = self._ber_filter_geraet.currentData()
        self._ber_filter_geraet.clear()
        self._ber_filter_geraet.addItem("Alle Geräte", None)
        for h in self._handys:
            self._ber_filter_geraet.addItem(h["inventarnummer"], h["inventarnummer"])
        if aktuell_ber is not None:
            for i in range(self._ber_filter_geraet.count()):
                if self._ber_filter_geraet.itemData(i) == aktuell_ber:
                    self._ber_filter_geraet.setCurrentIndex(i)
                    break
        self._ber_filter_geraet.blockSignals(False)

    def _lade_historie(self):
        tage_map = {"30 Tage": 30, "90 Tage": 90, "180 Tage": 180, "Gesamt": 0}
        tage = tage_map.get(self._hist_filter_tage.currentText(), 90)
        handy_id = self._hist_filter_geraet.currentData()

        try:
            eintraege = lade_historie(handy_id=handy_id, tage=tage)
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Historie konnte nicht geladen werden:\n{e}")
            return

        self._hist_tabelle.setSortingEnabled(False)
        self._hist_tabelle.setRowCount(len(eintraege))
        for row_idx, eintrag in enumerate(eintraege):
            for col_idx, (_, feld, _) in enumerate(_HISTORIE_SPALTEN):
                wert = str(eintrag.get(feld, "") or "")
                item = QTableWidgetItem(wert)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._hist_tabelle.setItem(row_idx, col_idx, item)
        self._hist_tabelle.setSortingEnabled(True)
    def _lade_berichte(self):
        from functions.handys_bericht import lade_berichte_fuer_handy, berichte_basis_pfad
        inv_filter = self._ber_filter_geraet.currentData()  # str oder None

        if inv_filter:
            berichte = lade_berichte_fuer_handy(inv_filter)
        else:
            # Alle Geräte zusammenführen
            berichte = []
            for h in self._handys:
                berichte.extend(lade_berichte_fuer_handy(h["inventarnummer"]))
            # Nach Datum sortieren (neueste zuerst)
            berichte.sort(key=lambda x: x["dateiname"], reverse=True)

        self._ber_tabelle.setSortingEnabled(False)
        self._ber_tabelle.setRowCount(len(berichte))
        for row_idx, b in enumerate(berichte):
            inv_aus_name = b["dateiname"].split("_")[1] if "_" in b["dateiname"] else ""
            for col_idx, wert in enumerate([
                inv_aus_name,
                b["typ"],
                b["datum_str"],
                b["dateiname"],
            ]):
                item = QTableWidgetItem(wert)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setData(Qt.ItemDataRole.UserRole, b["pfad"])
                self._ber_tabelle.setItem(row_idx, col_idx, item)
        self._ber_tabelle.setSortingEnabled(True)

        basis = berichte_basis_pfad()
        self._ber_pfad_lbl.setText(f"Speicherort: {basis}")

    def _bericht_oeffnen(self):
        zeilen = self._ber_tabelle.selectedItems()
        if not zeilen:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst einen Bericht auswählen.")
            return
        pfad = zeilen[0].data(Qt.ItemDataRole.UserRole)
        if pfad and os.path.isfile(pfad):
            os.startfile(pfad)
        else:
            QMessageBox.warning(self, "Fehler", f"Datei nicht gefunden:\n{pfad}")

    def _berichte_ordner_oeffnen(self):
        from functions.handys_bericht import berichte_basis_pfad
        import subprocess
        basis = berichte_basis_pfad()
        os.makedirs(basis, exist_ok=True)
        subprocess.Popen(f'explorer "{basis}"', shell=True)

    def _bericht_per_email_senden(self):
        """Sendet den ausgewählten Schadens-/Verlustbericht als E-Mail-Anhang."""
        zeilen = self._ber_tabelle.selectedItems()
        if not zeilen:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst einen Bericht auswählen.")
            return
        pfad = zeilen[0].data(Qt.ItemDataRole.UserRole)
        if not pfad or not os.path.isfile(pfad):
            QMessageBox.warning(self, "Fehler", f"Datei nicht gefunden:\n{pfad}")
            return

        # Gerät und Typ aus Tabelle auslesen (Spalte 0 = Gerät, Spalte 1 = Typ)
        zeile_nr = self._ber_tabelle.currentRow()
        inventarnummer = (self._ber_tabelle.item(zeile_nr, 0) or QTableWidgetItem("")).text()
        bericht_typ    = (self._ber_tabelle.item(zeile_nr, 1) or QTableWidgetItem("")).text()

        try:
            from functions.handys_email import sende_bericht_email
            erfolg, meldung = sende_bericht_email(pfad, inventarnummer, bericht_typ)
            if erfolg:
                QMessageBox.information(self, "E-Mail", meldung)
            else:
                QMessageBox.warning(self, "E-Mail – Hinweis", meldung)
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"E-Mail konnte nicht erstellt werden:\n{e}")

    # ─── Formular befüllen / leeren ───────────────────────────────────────────

    def _formular_leeren(self):
        self._aktueller_handy_id = None
        self._f_inventar.clear()
        self._f_hersteller.clear()
        self._f_modell.clear()
        self._f_rufnummer.clear()
        self._f_sim.clear()
        self._f_standort.clear()
        self._f_zustand.setCurrentIndex(0)
        self._f_defekt.clear()
        self._f_defekt_datum.setDate(QDate.currentDate())
        self._f_defekt_gemeldet_von.clear()
        self._f_anschaffung.setDate(QDate.currentDate())
        self._f_notizen.clear()
        self._f_kartennummer.clear()
        self._f_pin.clear()
        self._f_pin2.clear()
        self._f_puk.clear()
        self._f_puk2.clear()

    def _formular_befuellen(self, handy: dict):
        self._aktueller_handy_id = handy["id"]
        self._f_inventar.setText(handy.get("inventarnummer", "") or "")
        self._f_hersteller.setText(handy.get("hersteller", "") or "")
        self._f_modell.setText(handy.get("modell", "") or "")
        self._f_rufnummer.setText(handy.get("rufnummer", "") or "")
        self._f_sim.setText(handy.get("sim_nummer", "") or "")
        self._f_standort.setText(handy.get("standort", "") or "")
        idx = self._f_zustand.findText(handy.get("zustand", "Aktiv"))
        self._f_zustand.setCurrentIndex(max(0, idx))
        self._f_defekt.setText(handy.get("defekt_beschreibung", "") or "")
        dd = handy.get("defekt_datum", "")
        if dd:
            try:
                d = QDate.fromString(dd, "yyyy-MM-dd")
                if d.isValid():
                    self._f_defekt_datum.setDate(d)
            except Exception:
                pass
        else:
            self._f_defekt_datum.setDate(QDate.currentDate())
        self._f_defekt_gemeldet_von.setText(handy.get("defekt_gemeldet_von", "") or "")
        ansch = handy.get("anschaffungsdatum", "")
        if ansch:
            try:
                d = QDate.fromString(ansch, "yyyy-MM-dd")
                if d.isValid():
                    self._f_anschaffung.setDate(d)
            except Exception:
                pass
        self._f_notizen.setPlainText(handy.get("notizen", "") or "")
        self._f_kartennummer.setText(handy.get("kartennummer", "") or "")
        self._f_pin.setText(handy.get("pin", "") or "")
        self._f_pin2.setText(handy.get("pin2", "") or "")
        self._f_puk.setText(handy.get("puk", "") or "")
        self._f_puk2.setText(handy.get("puk2", "") or "")

    def _formular_zu_dict(self) -> dict:
        zustand = self._f_zustand.currentText()
        mit_detail = zustand in ZUSTAND_MIT_DEFEKT_DETAIL
        return {
            "inventarnummer":      self._f_inventar.text().strip(),
            "hersteller":          self._f_hersteller.text().strip(),
            "modell":              self._f_modell.text().strip(),
            "rufnummer":           self._f_rufnummer.text().strip(),
            "sim_nummer":          self._f_sim.text().strip(),
            "standort":            self._f_standort.text().strip(),
            "zustand":             zustand,
            "defekt_beschreibung": self._f_defekt.text().strip() if mit_detail else "",
            "defekt_datum":        self._f_defekt_datum.date().toString("yyyy-MM-dd") if mit_detail else "",
            "defekt_gemeldet_von": self._f_defekt_gemeldet_von.text().strip() if mit_detail else "",
            "anschaffungsdatum":   self._f_anschaffung.date().toString("yyyy-MM-dd"),
            "notizen":             self._f_notizen.toPlainText().strip(),
            "kartennummer":        self._f_kartennummer.text().strip(),
            "pin":                 self._f_pin.text().strip(),
            "pin2":                self._f_pin2.text().strip(),
            "puk":                 self._f_puk.text().strip(),
            "puk2":                self._f_puk2.text().strip(),
        }

    # ─── Aktionen ─────────────────────────────────────────────────────────────

    def _ausgewaehlter_handy_id(self) -> int | None:
        zeilen = self._tabelle.selectedItems()
        if not zeilen:
            return None
        row = zeilen[0].row()
        id_item = self._tabelle.item(row, 0)  # versteckte ID-Spalte
        if id_item:
            try:
                return int(id_item.text())
            except ValueError:
                return None
        return None

    def _auswahl_geaendert(self):
        handy_id = self._ausgewaehlter_handy_id()
        hat_auswahl = handy_id is not None
        self._btn_bearbeiten.setEnabled(hat_auswahl)
        self._btn_loeschen.setEnabled(hat_auswahl)
        # Bericht-Button nur bei Defekt oder Verloren
        bericht_aktiv = False
        if hat_auswahl:
            zeilen = self._tabelle.selectedItems()
            if zeilen:
                row = zeilen[0].row()
                # Zustand-Spalte finden
                for col_idx, (_, feld, _, _) in enumerate(_UEBERSICHT_SPALTEN):
                    if feld == "zustand":
                        z_item = self._tabelle.item(row, col_idx)
                        if z_item and z_item.text() in {"Defekt", "Verloren"}:
                            bericht_aktiv = True
                        break
        self._btn_bericht.setEnabled(bericht_aktiv)

    def _zustand_geaendert(self, zustand: str):
        mit_detail = zustand in ZUSTAND_MIT_DEFEKT_DETAIL
        # Defektbeschreibung nur bei relevantem Zustand aktivieren
        self._f_defekt.setEnabled(mit_detail)
        stil_aktiv = "border: 1px solid #d9d9d9; border-radius: 4px; padding: 0 6px;"
        stil_inaktiv = (
            "border: 1px solid #d9d9d9; border-radius: 4px; padding: 0 6px;"
            "background-color: #f5f5f5; color: #888;"
        )
        self._f_defekt.setStyleSheet(stil_aktiv if mit_detail else stil_inaktiv)
        # Festgestellt-am / Festgestellt-von nur einblenden wenn relevant
        self._lbl_defekt_datum.setVisible(mit_detail)
        self._f_defekt_datum.setVisible(mit_detail)
        self._lbl_defekt_gemeldet.setVisible(mit_detail)
        self._f_defekt_gemeldet_von.setVisible(mit_detail)

    def _neues_handy(self):
        self._formular_leeren()
        self._tabs.setCurrentIndex(1)
        self._f_inventar.setFocus()

    def _bearbeiten(self):
        handy_id = self._ausgewaehlter_handy_id()
        if handy_id is None:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst ein Gerät auswählen.")
            return
        handy = lade_handy(handy_id)
        if not handy:
            QMessageBox.warning(self, "Fehler", "Gerät nicht gefunden.")
            return
        self._formular_befuellen(handy)
        self._tabs.setCurrentIndex(1)

    def _loeschen(self):
        handy_id = self._ausgewaehlter_handy_id()
        if handy_id is None:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst ein Gerät auswählen.")
            return
        handy = lade_handy(handy_id)
        if not handy:
            return
        inv_nr = handy['inventarnummer']
        antwort = QMessageBox.question(
            self,
            "Löschen bestätigen",
            f"Gerät '{inv_nr}' wirklich löschen?\n"
            "Alle Historieneinträge werden ebenfalls gelöscht.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if antwort != QMessageBox.StandardButton.Yes:
            return
        if loesche_handy(handy_id):
            self.refresh()
        else:
            QMessageBox.warning(self, "Fehler", "Gerät konnte nicht gelöscht werden.")

    def _speichern(self):
        daten = self._formular_zu_dict()

        # Pflichtfeldprüfung
        if not daten["inventarnummer"]:
            QMessageBox.warning(self, "Pflichtfeld", "Inventarnummer darf nicht leer sein.")
            self._f_inventar.setFocus()
            return

        try:
            if self._aktueller_handy_id is None:
                erstelle_handy(daten)
                inv = daten['inventarnummer']
                QMessageBox.information(self, "Gespeichert",
                                        f"Gerät '{inv}' wurde angelegt.")
            else:
                aktualisiere_handy(self._aktueller_handy_id, daten)
                inv = daten['inventarnummer']
                QMessageBox.information(self, "Gespeichert",
                                        f"Gerät '{inv}' wurde aktualisiert.")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Speichern fehlgeschlagen:\n{e}")
            return

        self._formular_leeren()
        self.refresh()
        self._tabs.setCurrentIndex(0)

    def _abbrechen(self):
        self._formular_leeren()
        self._tabs.setCurrentIndex(0)

    def _excel_export(self):
        try:
            from functions.handys_excel_export import export_handys_excel
            pfad = export_handys_excel(open_after=True)
            QMessageBox.information(self, "Export erfolgreich",
                                    f"Excel gespeichert und geöffnet:\n{pfad}")
        except Exception as e:
            QMessageBox.critical(self, "Export-Fehler", f"Excel-Export fehlgeschlagen:\n{e}")

    def _email_senden(self):
        try:
            from functions.handys_email import sende_handys_email
            erfolg, meldung = sende_handys_email()
            if erfolg:
                QMessageBox.information(self, "E-Mail", meldung)
            else:
                QMessageBox.warning(self, "E-Mail – Hinweis", meldung)
        except Exception as e:
            QMessageBox.critical(self, "E-Mail-Fehler", f"Fehler beim E-Mail-Versand:\n{e}")

    def _schadensbericht_erstellen(self):
        handy_id = self._ausgewaehlter_handy_id()
        if handy_id is None:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst ein Gerät auswählen.")
            return
        handy = lade_handy(handy_id)
        if not handy:
            QMessageBox.warning(self, "Fehler", "Gerät nicht gefunden.")
            return
        zustand = handy.get("zustand", "")
        if zustand not in {"Defekt", "Verloren"}:
            QMessageBox.information(
                self, "Hinweis",
                "Ein Bericht kann nur für Geräte mit Zustand 'Defekt' oder 'Verloren' erstellt werden."
            )
            return

        from gui.schadensbericht_dialog import SchadsberichtDialog
        dlg = SchadsberichtDialog(handy, parent=self)
        if dlg.exec() != SchadsberichtDialog.DialogCode.Accepted:
            return

        bericht_daten = dlg.get_bericht_daten()
        try:
            from functions.handys_bericht import erstelle_schadensbericht
            pfad = erstelle_schadensbericht(handy, bericht_daten)
            bericht_typ = "Verlustbericht" if zustand == "Verloren" else "Schadensbericht"
            antwort = QMessageBox.question(
                self,
                f"{bericht_typ} erstellt",
                f"Bericht gespeichert:\n{pfad}\n\nJetzt öffnen?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if antwort == QMessageBox.StandardButton.Yes:
                os.startfile(pfad)
            # Berichte-Tab neu laden
            self._lade_berichte()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Bericht konnte nicht erstellt werden:\n{e}")
