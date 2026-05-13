"""
Schadens-/Verlustbericht Dialog
Ermöglicht das Ausfüllen und Bestätigen eines Berichts vor der Word-Erstellung.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QDateEdit, QTextEdit, QCheckBox, QFrame,
    QFormLayout, QScrollArea, QWidget, QGroupBox,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont

from config import FIORI_BLUE, FIORI_TEXT


def _feld_stil() -> str:
    return "border: 1px solid #d9d9d9; border-radius: 4px; padding: 2px 6px; font-size: 11px;"

def _readonly_stil() -> str:
    return (
        "border: 1px solid #e0e0e0; border-radius: 4px; padding: 2px 6px;"
        "background-color: #f5f5f5; color: #555; font-size: 11px;"
    )

def _gruppe(titel: str) -> QGroupBox:
    gb = QGroupBox(titel)
    gb.setStyleSheet("""
        QGroupBox {
            font-weight: bold; font-size: 12px; color: #1565a8;
            border: 1px solid #c0cce0; border-radius: 6px;
            margin-top: 10px; padding-top: 6px;
        }
        QGroupBox::title {
            subcontrol-origin: margin; subcontrol-position: top left;
            padding: 0 6px; left: 10px;
        }
    """)
    return gb

def _btn(text: str, color: str = FIORI_BLUE) -> QPushButton:
    b = QPushButton(text)
    b.setFixedHeight(34)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    b.setStyleSheet(f"""
        QPushButton {{
            background-color: {color}; color: white; border: none;
            border-radius: 4px; padding: 0 16px;
            font-size: 12px; font-weight: bold;
        }}
        QPushButton:hover {{ background-color: #0856a8; }}
    """)
    return b


class SchadsberichtDialog(QDialog):
    """
    Dialog zum Ausfüllen eines Schadens- oder Verlustberichts.
    Felder werden aus dem handy-Dict vorbelegt.
    """

    _MASSNAHMEN_DEFEKT = [
        "Gerät aus dem Dienst genommen",
        "Reparatur veranlasst",
        "Ersatzgerät ausgegeben",
        "Entsorgung / Austausch veranlasst",
    ]
    _MASSNAHMEN_VERLOREN = [
        "SIM-Karte beim Anbieter gesperrt",
        "Gerät als verloren gemeldet (Polizei / Fundbüro)",
        "Vorgesetzter informiert",
        "Ersatzgerät ausgegeben",
    ]

    def __init__(self, handy: dict, parent=None):
        super().__init__(parent)
        self._handy = handy
        self._ist_verloren = handy.get("zustand", "") == "Verloren"
        bericht_typ = "Verlustbericht" if self._ist_verloren else "Schadensbericht"
        self.setWindowTitle(f"{bericht_typ} – {handy.get('inventarnummer', '')}")
        self.setMinimumWidth(620)
        self.setMinimumHeight(700)
        self._build_ui()
        self._prefill()

    # ── UI-Aufbau ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 12)
        outer.setSpacing(10)

        # Titel
        bericht_typ = "VERLUSTBERICHT" if self._ist_verloren else "SCHADENSBERICHT"
        titel = QLabel(f"📋  {bericht_typ}  –  {self._handy.get('inventarnummer', '')}")
        titel.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        farbe = "#5c007a" if self._ist_verloren else "#9C0006"
        titel.setStyleSheet(f"color: {farbe}; padding-bottom: 4px;")
        outer.addWidget(titel)

        # Scrollbarer Inhaltsbereich
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        vl = QVBoxLayout(inner)
        vl.setSpacing(10)
        vl.setContentsMargins(0, 0, 8, 0)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        # ── Gruppe 1: Gerätedaten (read-only) ────────────────────────────────
        gb1 = _gruppe("1. Gerätedaten")
        fl1 = QFormLayout(gb1)
        fl1.setSpacing(6)
        fl1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._ro_inv    = self._ro(self._handy.get("inventarnummer", ""))
        self._ro_hersteller = self._ro(
            f"{self._handy.get('hersteller', '')}  {self._handy.get('modell', '')}".strip()
        )
        self._ro_rufnr  = self._ro(self._handy.get("rufnummer", ""))
        self._ro_karte  = self._ro(self._handy.get("kartennummer", ""))
        self._ro_stand  = self._ro(self._handy.get("standort", ""))
        fl1.addRow("Inventarnummer:",    self._ro_inv)
        fl1.addRow("Hersteller / Modell:", self._ro_hersteller)
        fl1.addRow("Rufnummer:",         self._ro_rufnr)
        fl1.addRow("SIM-Kartennummer:",  self._ro_karte)
        fl1.addRow("Standort:",          self._ro_stand)
        vl.addWidget(gb1)

        # ── Gruppe 2: Meldung ─────────────────────────────────────────────────
        gb2 = _gruppe("2. Meldung")
        fl2 = QFormLayout(gb2)
        fl2.setSpacing(6)
        fl2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._f_datum = QDateEdit()
        self._f_datum.setCalendarPopup(True)
        self._f_datum.setDisplayFormat("dd.MM.yyyy")
        self._f_datum.setStyleSheet(_feld_stil())
        self._f_datum.setFixedHeight(28)
        self._f_von = QLineEdit()
        self._f_von.setStyleSheet(_feld_stil())
        self._f_von.setFixedHeight(28)
        self._f_von.setPlaceholderText("Name der meldenden Person")
        fl2.addRow("Festgestellt am:",   self._f_datum)
        fl2.addRow("Festgestellt von:",  self._f_von)
        vl.addWidget(gb2)

        # ── Gruppe 3: Beschreibung ────────────────────────────────────────────
        lbl3 = "3. Verlustbeschreibung" if self._ist_verloren else "3. Schadensbeschreibung"
        gb3 = _gruppe(lbl3)
        vl3 = QVBoxLayout(gb3)
        self._f_beschreibung = QTextEdit()
        self._f_beschreibung.setFixedHeight(100)
        self._f_beschreibung.setStyleSheet(_feld_stil())
        self._f_beschreibung.setPlaceholderText(
            "Beschreibung des Verlusts…" if self._ist_verloren
            else "Beschreibung des Schadens…"
        )
        vl3.addWidget(self._f_beschreibung)
        vl.addWidget(gb3)

        # ── Gruppe 4: Maßnahmen (Checkboxen) ─────────────────────────────────
        gb4 = _gruppe("4. Eingeleitete Maßnahmen")
        vl4 = QVBoxLayout(gb4)
        vl4.setSpacing(4)
        massnahmen_liste = (
            self._MASSNAHMEN_VERLOREN if self._ist_verloren else self._MASSNAHMEN_DEFEKT
        )
        self._cb_massnahmen: list[QCheckBox] = []
        for m in massnahmen_liste:
            cb = QCheckBox(m)
            cb.setStyleSheet("font-size: 11px;")
            self._cb_massnahmen.append(cb)
            vl4.addWidget(cb)
        # Sonstiges
        son_row = QHBoxLayout()
        self._cb_sonstiges = QCheckBox("Sonstiges:")
        self._cb_sonstiges.setStyleSheet("font-size: 11px;")
        self._f_sonstiges = QLineEdit()
        self._f_sonstiges.setStyleSheet(_feld_stil())
        self._f_sonstiges.setFixedHeight(26)
        self._f_sonstiges.setPlaceholderText("Freitext…")
        self._cb_sonstiges.toggled.connect(self._f_sonstiges.setEnabled)
        self._f_sonstiges.setEnabled(False)
        son_row.addWidget(self._cb_sonstiges)
        son_row.addWidget(self._f_sonstiges, 1)
        vl4.addLayout(son_row)
        vl.addWidget(gb4)

        # ── Gruppe 5: Notizen ─────────────────────────────────────────────────
        gb5 = _gruppe("5. Notizen / Weitere Anmerkungen")
        vl5 = QVBoxLayout(gb5)
        self._f_notizen = QTextEdit()
        self._f_notizen.setFixedHeight(80)
        self._f_notizen.setStyleSheet(_feld_stil())
        self._f_notizen.setPlaceholderText("Weitere Anmerkungen…")
        vl5.addWidget(self._f_notizen)
        vl.addWidget(gb5)

        # ── Gruppe 6: Ersteller ───────────────────────────────────────────────
        gb6 = _gruppe("6. Ersteller")
        fl6 = QFormLayout(gb6)
        fl6.setSpacing(6)
        fl6.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._f_ersteller = QLineEdit()
        self._f_ersteller.setStyleSheet(_feld_stil())
        self._f_ersteller.setFixedHeight(28)
        self._f_ersteller.setPlaceholderText("Name des Schichtleiters")
        fl6.addRow("Erstellt von:", self._f_ersteller)
        vl.addWidget(gb6)

        vl.addStretch()

        # ── Button-Leiste ─────────────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #d0d8e0;")
        outer.addWidget(sep)

        btn_row = QHBoxLayout()
        farbe_ok = "#5c007a" if self._ist_verloren else "#bb0000"
        bericht_typ_lbl = "Verlustbericht" if self._ist_verloren else "Schadensbericht"
        self._btn_ok     = _btn(f"📄 {bericht_typ_lbl} erstellen", farbe_ok)
        self._btn_cancel = _btn("Abbrechen", "#888888")
        btn_row.addStretch()
        btn_row.addWidget(self._btn_cancel)
        btn_row.addWidget(self._btn_ok)
        outer.addLayout(btn_row)

        self._btn_ok.clicked.connect(self.accept)
        self._btn_cancel.clicked.connect(self.reject)

    # ── Vorbelegen ────────────────────────────────────────────────────────────

    def _prefill(self):
        h = self._handy
        # Datum
        dd = h.get("defekt_datum", "")
        if dd:
            d = QDate.fromString(dd, "yyyy-MM-dd")
            if d.isValid():
                self._f_datum.setDate(d)
            else:
                self._f_datum.setDate(QDate.currentDate())
        else:
            self._f_datum.setDate(QDate.currentDate())

        self._f_von.setText(h.get("defekt_gemeldet_von", "") or "")
        self._f_beschreibung.setPlainText(h.get("defekt_beschreibung", "") or "")
        self._f_notizen.setPlainText(h.get("notizen", "") or "")

        # Ersteller aus Systemuser
        try:
            self._f_ersteller.setText(os.getlogin())
        except Exception:
            pass

    # ── Ergebnis auslesen ─────────────────────────────────────────────────────

    def get_bericht_daten(self) -> dict:
        """Gibt die Dialog-Eingaben als dict zurück."""
        massnahmen = [
            cb.text() for cb in self._cb_massnahmen if cb.isChecked()
        ]
        sonstiges = (
            self._f_sonstiges.text().strip()
            if self._cb_sonstiges.isChecked() else ""
        )
        return {
            "ersteller":    self._f_ersteller.text().strip(),
            "defekt_datum": self._f_datum.date().toString("yyyy-MM-dd"),
            "defekt_von":   self._f_von.text().strip(),
            "beschreibung": self._f_beschreibung.toPlainText().strip(),
            "massnahmen":   massnahmen,
            "sonstiges":    sonstiges,
            "notizen":      self._f_notizen.toPlainText().strip(),
        }

    # ── Hilfsmethode ──────────────────────────────────────────────────────────

    @staticmethod
    def _ro(text: str) -> QLineEdit:
        f = QLineEdit(text or "–")
        f.setReadOnly(True)
        f.setStyleSheet(_readonly_stil())
        f.setFixedHeight(26)
        return f
