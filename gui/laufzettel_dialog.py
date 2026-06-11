"""
Laufzettel-Dialog
Erstellt Laufzettel (Onboarding-Checklisten) für Mitarbeiter als Word-Dokument.
"""
from __future__ import annotations

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QRadioButton, QButtonGroup,
    QComboBox, QFormLayout, QLineEdit, QFrame, QScrollArea,
    QWidget, QMessageBox, QDateEdit, QSizePolicy, QDialogButtonBox,
    QStackedWidget, QSpacerItem,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont, QColor

from config import FIORI_BLUE, FIORI_TEXT, FIORI_WHITE, FIORI_BORDER

# ─── Schulungsauswahl-Konfiguration ──────────────────────────────────────────
_SCHULUNGEN = [
    "Vorfeldschulung",
    "PRM-Schulung",
    "SI-Schulung (Sicherheitsschulung)",
    "Arbeitsschutz",
    "Erste Hilfe",
]

# ─── Style-Helfer ─────────────────────────────────────────────────────────────

_GROUP_STYLE = (
    "QGroupBox { font-weight: bold; font-size: 12px; color: #1565a8;"
    " border: 1px solid #c5d8f0; border-radius: 6px;"
    " margin-top: 6px; padding-top: 10px; background: #f8fbff; }"
    " QGroupBox::title { subcontrol-origin: margin; left: 12px;"
    " padding: 0 4px; background: #f8fbff; }"
)

_FIELD_STYLE = (
    "QLineEdit, QComboBox, QDateEdit {"
    " border: 1px solid #ccc; border-radius: 4px;"
    " padding: 4px; font-size: 12px; background: white; }"
)


def _btn(text: str, color: str = FIORI_BLUE, hover: str = "#0057b8") -> QPushButton:
    btn = QPushButton(text)
    btn.setFixedHeight(34)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setStyleSheet(f"""
        QPushButton {{
            background: {color}; color: white; border: none;
            border-radius: 4px; padding: 4px 18px; font-size: 12px;
        }}
        QPushButton:hover {{ background: {hover}; }}
        QPushButton:disabled {{ background: #bbb; color: #888; }}
    """)
    return btn


# ══════════════════════════════════════════════════════════════════════════════
#  Dialog: Neuen Mitarbeiter anlegen (in beiden Datenbanken)
# ══════════════════════════════════════════════════════════════════════════════

class _NeuerMitarbeiterDialog(QDialog):
    """Legt einen neuen Mitarbeiter in mitarbeiter.db und schulungen.db an."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("👤 Neuer Mitarbeiter anlegen")
        self.setMinimumWidth(460)
        self.resize(480, 380)
        self._result_ma_id: int | None = None
        self._result_name: str = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(10)

        title = QLabel("Neuen Mitarbeiter anlegen")
        title.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {FIORI_BLUE};")
        layout.addWidget(title)

        info = QLabel(
            "Der Mitarbeiter wird in der Mitarbeiter-Datenbank und in der "
            "Schulungs-Datenbank angelegt und steht danach überall zur Verfügung."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #555; font-size: 11px;")
        layout.addWidget(info)

        grp = QGroupBox("Stammdaten")
        grp.setStyleSheet(_GROUP_STYLE)
        fl = QFormLayout(grp)
        fl.setSpacing(8)

        self._vorname = QLineEdit()
        self._vorname.setPlaceholderText("Vorname")
        self._vorname.setStyleSheet(_FIELD_STYLE)
        fl.addRow("Vorname *:", self._vorname)

        self._nachname = QLineEdit()
        self._nachname.setPlaceholderText("Nachname")
        self._nachname.setStyleSheet(_FIELD_STYLE)
        fl.addRow("Nachname *:", self._nachname)

        self._personalnr = QLineEdit()
        self._personalnr.setPlaceholderText("optional")
        self._personalnr.setStyleSheet(_FIELD_STYLE)
        fl.addRow("Personalnummer:", self._personalnr)

        self._funktion = QComboBox()
        self._funktion.setStyleSheet(_FIELD_STYLE)
        self._funktion.addItems(["stamm", "dispo", "aushilfe", "praktikant"])
        fl.addRow("Funktion:", self._funktion)

        self._eintrittsdatum = QDateEdit()
        self._eintrittsdatum.setCalendarPopup(True)
        self._eintrittsdatum.setDisplayFormat("dd.MM.yyyy")
        self._eintrittsdatum.setDate(QDate.currentDate())
        self._eintrittsdatum.setStyleSheet(_FIELD_STYLE)
        fl.addRow("Eintrittsdatum:", self._eintrittsdatum)

        layout.addWidget(grp)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Abbrechen")
        btn_cancel.setFixedHeight(32)
        btn_cancel.setStyleSheet(
            "QPushButton { background: #eee; color: #333; border: none;"
            " border-radius: 4px; padding: 4px 16px; font-size: 12px; }"
            " QPushButton:hover { background: #ddd; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        btn_ok = _btn("✅  Anlegen", "#107e3e", "#0a5c2e")
        btn_ok.clicked.connect(self._anlegen)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _anlegen(self):
        vorname  = self._vorname.text().strip()
        nachname = self._nachname.text().strip()

        if not vorname or not nachname:
            QMessageBox.warning(self, "Pflichtfelder", "Bitte Vor- und Nachname eingeben.")
            return

        eintritt = self._eintrittsdatum.date().toString("yyyy-MM-dd")

        # ── In mitarbeiter.db anlegen ─────────────────────────────────────────
        try:
            from database.models import Mitarbeiter
            from functions.mitarbeiter_functions import mitarbeiter_erstellen

            ma = Mitarbeiter(
                vorname=vorname,
                nachname=nachname,
                personalnummer=self._personalnr.text().strip(),
                funktion=self._funktion.currentText(),
                eintrittsdatum=None,
                status="aktiv",
            )
            ma_angelegt = mitarbeiter_erstellen(ma)
            self._result_ma_id = ma_angelegt.id
        except Exception as e:
            QMessageBox.critical(
                self, "Fehler (Mitarbeiter-DB)",
                f"Mitarbeiter konnte nicht in der Hauptdatenbank angelegt werden:\n{e}"
            )
            return

        # ── In schulungen.db anlegen ──────────────────────────────────────────
        try:
            from functions.schulungen_db import speichere_mitarbeiter
            speichere_mitarbeiter({
                "vorname":      vorname,
                "nachname":     nachname,
                "anstellung":   eintritt,
                "qualifikation": self._funktion.currentText(),
                "bemerkung":    f"Angelegt via Laufzettel am {QDate.currentDate().toString('dd.MM.yyyy')}",
            })
        except Exception as e:
            # Kein fataler Fehler – warnen, aber fortfahren
            QMessageBox.warning(
                self, "Hinweis (Schulungs-DB)",
                f"Mitarbeiter in Schulungs-Datenbank konnte nicht angelegt werden:\n{e}\n\n"
                "Der Mitarbeiter wurde dennoch in der Hauptdatenbank gespeichert."
            )

        self._result_name = f"{vorname} {nachname}"
        self.accept()

    def get_result(self) -> tuple[int | None, str]:
        """Gibt (ma_id, vollname) zurück."""
        return self._result_ma_id, self._result_name


# ══════════════════════════════════════════════════════════════════════════════
#  Haupt-Dialog: LaufzettelDialog
# ══════════════════════════════════════════════════════════════════════════════

class LaufzettelDialog(QDialog):
    """
    Erstellt einen Laufzettel für einen Mitarbeiter als Word-Dokument.
    Schritte:
      1) Mitarbeiter wählen (vorhanden) oder neu anlegen
      2) Vorlage / Inhalte konfigurieren
      3) Word-Datei erzeugen
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📋 Laufzettel erstellen")
        self.setMinimumWidth(620)
        self.setMinimumHeight(700)
        self.resize(660, 750)
        self._ma_name: str = ""
        self._build_ui()
        self._lade_mitarbeiter()

    # ── UI-Aufbau ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        # ── Titel-Leiste ──────────────────────────────────────────────────────
        header = QFrame()
        header.setFixedHeight(58)
        header.setStyleSheet(f"background: {FIORI_BLUE};")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(20, 0, 20, 0)

        lbl = QLabel("📋  Laufzettel erstellen")
        lbl.setFont(QFont("Arial", 15, QFont.Weight.Bold))
        lbl.setStyleSheet("color: white;")
        hl.addWidget(lbl)
        hl.addStretch()
        main.addWidget(header)

        # ── Scrollbarer Inhalt ────────────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; }")
        content = QWidget()
        content.setStyleSheet("background: #f5f7fa;")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(14)
        scroll.setWidget(content)
        main.addWidget(scroll, 1)

        # ── ABSCHNITT 1: Mitarbeiter ──────────────────────────────────────────
        grp_ma = QGroupBox("👤  Mitarbeiter")
        grp_ma.setStyleSheet(_GROUP_STYLE)
        ma_layout = QVBoxLayout(grp_ma)
        ma_layout.setSpacing(8)

        # Radio: vorhanden / neu
        radio_row = QHBoxLayout()
        self._radio_vorhanden = QRadioButton("Vorhandenen Mitarbeiter wählen")
        self._radio_neu       = QRadioButton("Neuen Mitarbeiter anlegen")
        self._radio_vorhanden.setChecked(True)
        self._radio_vorhanden.toggled.connect(self._ma_modus_geaendert)
        radio_row.addWidget(self._radio_vorhanden)
        radio_row.addWidget(self._radio_neu)
        radio_row.addStretch()
        ma_layout.addLayout(radio_row)

        # ── Vorhandener Mitarbeiter: Combobox ─────────────────────────────────
        self._ma_vorhanden_frame = QFrame()
        vf_layout = QHBoxLayout(self._ma_vorhanden_frame)
        vf_layout.setContentsMargins(0, 0, 0, 0)
        vf_layout.setSpacing(8)

        self._ma_combo = QComboBox()
        self._ma_combo.setMinimumWidth(300)
        self._ma_combo.setStyleSheet(_FIELD_STYLE)
        self._ma_combo.setEditable(True)
        self._ma_combo.currentIndexChanged.connect(self._ma_combo_geaendert)
        vf_layout.addWidget(QLabel("Mitarbeiter:"))
        vf_layout.addWidget(self._ma_combo, 1)

        btn_refresh = QPushButton("🔄")
        btn_refresh.setFixedSize(32, 28)
        btn_refresh.setToolTip("Liste neu laden")
        btn_refresh.setStyleSheet(
            "QPushButton { background: #e8f0fe; border: 1px solid #c5d8f0;"
            " border-radius: 4px; } QPushButton:hover { background: #c5d8f0; }"
        )
        btn_refresh.clicked.connect(self._lade_mitarbeiter)
        vf_layout.addWidget(btn_refresh)
        ma_layout.addWidget(self._ma_vorhanden_frame)

        # ── Neuer Mitarbeiter: Button ─────────────────────────────────────────
        self._ma_neu_frame = QFrame()
        self._ma_neu_frame.setVisible(False)
        nf_layout = QHBoxLayout(self._ma_neu_frame)
        nf_layout.setContentsMargins(0, 0, 0, 0)
        nf_layout.setSpacing(8)

        self._ma_neu_info = QLabel("Noch kein Mitarbeiter angelegt.")
        self._ma_neu_info.setStyleSheet("color: #666; font-style: italic; font-size: 11px;")
        nf_layout.addWidget(self._ma_neu_info, 1)

        btn_neu_ma = _btn("👤  Mitarbeiter anlegen", "#107e3e", "#0a5c2e")
        btn_neu_ma.clicked.connect(self._neuen_ma_anlegen)
        nf_layout.addWidget(btn_neu_ma)
        ma_layout.addWidget(self._ma_neu_frame)

        cl.addWidget(grp_ma)

        # ── ABSCHNITT 2: Vorlagen ─────────────────────────────────────────────
        grp_vorlagen = QGroupBox("📄  Vorgefertigte Abschnitte")
        grp_vorlagen.setStyleSheet(_GROUP_STYLE)
        vl_layout = QVBoxLayout(grp_vorlagen)
        vl_layout.setSpacing(6)

        info_vorlagen = QLabel(
            "Wähle vorgefertigte Abschnitte, die in den Laufzettel aufgenommen werden sollen:"
        )
        info_vorlagen.setWordWrap(True)
        info_vorlagen.setStyleSheet("font-size: 11px; color: #555;")
        vl_layout.addWidget(info_vorlagen)

        # Neu-Einstellung / ZÜP
        self._chk_zuep = QCheckBox(
            "🆕  Neu-Einstellung / ZÜP"
        )
        self._chk_zuep.setToolTip(
            "ZÜP-Unterlagen (Ausweiskopie, Wohnorte 10J, Arbeitgeber 5J, Zeugnisse)\n"
            "Ansprechpartner: Schichtleiter + Herr Peters"
        )
        self._chk_zuep.setStyleSheet("font-size: 12px; padding: 2px;")
        vl_layout.addWidget(self._chk_zuep)

        zuep_detail = QLabel(
            "    → Unterlagen: Ausweiskopie, Wohnorte letzte 10 Jahre, "
            "Arbeitgeber letzte 5 Jahre,\n"
            "       Arbeitszeugnisse, letzte Lohnabrechnung u. a.\n"
            "    → Ansprechpartner: Schichtleiter + Herr Peters"
        )
        zuep_detail.setStyleSheet("font-size: 10px; color: #666; padding-left: 20px;")
        vl_layout.addWidget(zuep_detail)

        # Dienstkleidung
        self._chk_kleidung = QCheckBox("👕  Dienstkleidung")
        self._chk_kleidung.setToolTip(
            "Dienstkleidungsausgabe\n"
            "Ansprechpartner: Herr Etz, Herr Kurthen (+ Schichtleiter)"
        )
        self._chk_kleidung.setStyleSheet("font-size: 12px; padding: 2px;")
        vl_layout.addWidget(self._chk_kleidung)

        kleidung_detail = QLabel(
            "    → Hauptansprechpartner: Herr Etz, Herr Kurthen\n"
            "    → Alternativ: Alle Schichtleiter können behilflich sein"
        )
        kleidung_detail.setStyleSheet("font-size: 10px; color: #666; padding-left: 20px;")
        vl_layout.addWidget(kleidung_detail)

        cl.addWidget(grp_vorlagen)

        # ── ABSCHNITT 3: Schulungen ───────────────────────────────────────────
        grp_sch = QGroupBox("🎓  Schulungen")
        grp_sch.setStyleSheet(_GROUP_STYLE)
        sch_layout = QVBoxLayout(grp_sch)
        sch_layout.setSpacing(6)

        info_sch = QLabel(
            "Wähle die Schulungen, für die Termine gemacht werden müssen:"
        )
        info_sch.setWordWrap(True)
        info_sch.setStyleSheet("font-size: 11px; color: #555;")
        sch_layout.addWidget(info_sch)

        self._sch_checkboxen: dict[str, QCheckBox] = {}
        for sch in _SCHULUNGEN:
            chk = QCheckBox(sch)
            chk.setStyleSheet("font-size: 12px; padding: 2px;")
            sch_layout.addWidget(chk)
            self._sch_checkboxen[sch] = chk

        # ── Einarbeitung / Refresher ──────────────────────────────────────────
        sep_line = QFrame()
        sep_line.setFrameShape(QFrame.Shape.HLine)
        sep_line.setStyleSheet("color: #ccc;")
        sch_layout.addWidget(sep_line)

        einarbeitung_label = QLabel("Einarbeitung / Refresher:")
        einarbeitung_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #1565a8;")
        sch_layout.addWidget(einarbeitung_label)

        einarbeitung_row = QHBoxLayout()
        self._radio_keine_einarbeitung = QRadioButton("Keins davon")
        self._radio_einarbeitung       = QRadioButton("Einarbeitung")
        self._radio_refresher          = QRadioButton("Refresher-Schulung")
        self._radio_keine_einarbeitung.setChecked(True)

        for rb in (self._radio_keine_einarbeitung, self._radio_einarbeitung, self._radio_refresher):
            rb.setStyleSheet("font-size: 12px;")
            einarbeitung_row.addWidget(rb)
        einarbeitung_row.addStretch()
        sch_layout.addLayout(einarbeitung_row)

        cl.addWidget(grp_sch)

        # ── ABSCHNITT 4: Schnellauswahl ───────────────────────────────────────
        grp_quick = QGroupBox("⚡  Schnellauswahl")
        grp_quick.setStyleSheet(_GROUP_STYLE)
        ql = QHBoxLayout(grp_quick)
        ql.setSpacing(8)

        lbl_quick = QLabel("Vorlage laden:")
        lbl_quick.setStyleSheet("font-size: 12px;")
        ql.addWidget(lbl_quick)

        btn_vorlage_neu = QPushButton("📋  Komplette Neu-Einstellung")
        btn_vorlage_neu.setFixedHeight(32)
        btn_vorlage_neu.setStyleSheet(
            "QPushButton { background: #e8f5e9; color: #1b5e20; border: 1px solid #a5d6a7;"
            " border-radius: 4px; padding: 4px 12px; font-size: 11px; }"
            " QPushButton:hover { background: #c8e6c9; }"
        )
        btn_vorlage_neu.setToolTip("Alle ZÜP, Dienstkleidung und alle Pflichtschulungen aktivieren")
        btn_vorlage_neu.clicked.connect(self._vorlage_neu_einstellung)
        ql.addWidget(btn_vorlage_neu)

        btn_vorlage_kleidung = QPushButton("👕  Nur Dienstkleidung")
        btn_vorlage_kleidung.setFixedHeight(32)
        btn_vorlage_kleidung.setStyleSheet(
            "QPushButton { background: #e3f2fd; color: #0d47a1; border: 1px solid #90caf9;"
            " border-radius: 4px; padding: 4px 12px; font-size: 11px; }"
            " QPushButton:hover { background: #bbdefb; }"
        )
        btn_vorlage_kleidung.setToolTip("Nur Dienstkleidungs-Abschnitt aktivieren")
        btn_vorlage_kleidung.clicked.connect(self._vorlage_kleidung)
        ql.addWidget(btn_vorlage_kleidung)

        btn_reset = QPushButton("✖  Alles zurücksetzen")
        btn_reset.setFixedHeight(32)
        btn_reset.setStyleSheet(
            "QPushButton { background: #fff3e0; color: #e65100; border: 1px solid #ffcc80;"
            " border-radius: 4px; padding: 4px 12px; font-size: 11px; }"
            " QPushButton:hover { background: #ffe0b2; }"
        )
        btn_reset.clicked.connect(self._alles_zuruecksetzen)
        ql.addWidget(btn_reset)

        ql.addStretch()
        cl.addWidget(grp_quick)
        cl.addStretch()

        # ── Footer-Leiste ─────────────────────────────────────────────────────
        footer = QFrame()
        footer.setFixedHeight(58)
        footer.setStyleSheet(
            "QFrame { background: #ffffff; border-top: 1px solid #ddd; }"
        )
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(20, 8, 20, 8)
        fl.setSpacing(10)

        self._status_lbl = QLabel("")
        self._status_lbl.setStyleSheet("font-size: 11px; color: #555;")
        fl.addWidget(self._status_lbl, 1)

        btn_cancel = QPushButton("Abbrechen")
        btn_cancel.setFixedHeight(34)
        btn_cancel.setStyleSheet(
            "QPushButton { background: #eee; color: #333; border: none;"
            " border-radius: 4px; padding: 4px 18px; font-size: 12px; }"
            " QPushButton:hover { background: #ddd; }"
        )
        btn_cancel.clicked.connect(self.reject)
        fl.addWidget(btn_cancel)

        self._btn_erstellen = _btn("📄  Word-Dokument erstellen", FIORI_BLUE)
        self._btn_erstellen.clicked.connect(self._erstellen)
        fl.addWidget(self._btn_erstellen)

        main.addWidget(footer)

    # ── Daten laden ───────────────────────────────────────────────────────────

    def _lade_mitarbeiter(self):
        """Lädt alle aktiven Mitarbeiter in die Combobox."""
        self._ma_combo.blockSignals(True)
        self._ma_combo.clear()
        self._ma_combo.addItem("— Mitarbeiter wählen —", None)
        try:
            from functions.mitarbeiter_functions import get_alle_mitarbeiter
            alle = get_alle_mitarbeiter(nur_aktive=True)
            for ma in alle:
                self._ma_combo.addItem(
                    f"{ma.nachname}, {ma.vorname}",
                    {"id": ma.id, "name": f"{ma.vorname} {ma.nachname}"}
                )
        except Exception:
            pass
        self._ma_combo.blockSignals(False)

    def _ma_combo_geaendert(self):
        data = self._ma_combo.currentData()
        if data:
            self._ma_name = data.get("name", "")
        else:
            self._ma_name = ""

    # ── Event-Handler ─────────────────────────────────────────────────────────

    def _ma_modus_geaendert(self, checked: bool):
        self._ma_vorhanden_frame.setVisible(self._radio_vorhanden.isChecked())
        self._ma_neu_frame.setVisible(self._radio_neu.isChecked())
        if self._radio_neu.isChecked():
            self._ma_name = ""
            self._ma_neu_info.setText("Noch kein Mitarbeiter angelegt.")

    def _neuen_ma_anlegen(self):
        dlg = _NeuerMitarbeiterDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            _, name = dlg.get_result()
            self._ma_name = name
            self._ma_neu_info.setText(f"✅  Angelegt: {name}")
            self._ma_neu_info.setStyleSheet("color: #1b5e20; font-weight: bold; font-size: 11px;")
            # Combobox aktualisieren für zukünftige Verwendung
            self._lade_mitarbeiter()

    def _vorlage_neu_einstellung(self):
        """Aktiviert alle Elemente für eine komplette Neu-Einstellung."""
        self._chk_zuep.setChecked(True)
        self._chk_kleidung.setChecked(True)
        for chk in self._sch_checkboxen.values():
            chk.setChecked(True)
        self._radio_einarbeitung.setChecked(True)

    def _vorlage_kleidung(self):
        """Aktiviert nur den Dienstkleidungs-Abschnitt."""
        self._alles_zuruecksetzen()
        self._chk_kleidung.setChecked(True)

    def _alles_zuruecksetzen(self):
        self._chk_zuep.setChecked(False)
        self._chk_kleidung.setChecked(False)
        for chk in self._sch_checkboxen.values():
            chk.setChecked(False)
        self._radio_keine_einarbeitung.setChecked(True)

    # ── Erstellen ─────────────────────────────────────────────────────────────

    def _erstellen(self):
        # ── Mitarbeiter ermitteln ─────────────────────────────────────────────
        ma_name = ""
        if self._radio_vorhanden.isChecked():
            data = self._ma_combo.currentData()
            if data:
                ma_name = data.get("name", "")
            else:
                # Freitext aus Combobox akzeptieren
                ma_name = self._ma_combo.currentText().strip()
                if ma_name == "— Mitarbeiter wählen —":
                    ma_name = ""
        else:
            ma_name = self._ma_name

        if not ma_name:
            QMessageBox.warning(
                self, "Mitarbeiter fehlt",
                "Bitte wähle einen Mitarbeiter aus oder lege einen neuen an."
            )
            return

        # ── Vorlagen ──────────────────────────────────────────────────────────
        vorlagen = []
        if self._chk_zuep.isChecked():
            vorlagen.append("Neu-Einstellung")
        if self._chk_kleidung.isChecked():
            vorlagen.append("Dienstkleidung")

        # ── Schulungen ────────────────────────────────────────────────────────
        schulungen = [
            sch for sch, chk in self._sch_checkboxen.items() if chk.isChecked()
        ]

        # ── Einarbeitung / Refresher ──────────────────────────────────────────
        if self._radio_einarbeitung.isChecked():
            einarbeitung_typ = "Einarbeitung"
        elif self._radio_refresher.isChecked():
            einarbeitung_typ = "Refresher-Schulung"
        else:
            einarbeitung_typ = None

        # Mindestens etwas ausgewählt?
        if not vorlagen and not schulungen and not einarbeitung_typ:
            QMessageBox.warning(
                self, "Nichts ausgewählt",
                "Bitte wähle mindestens einen Abschnitt oder eine Schulung aus."
            )
            return

        # ── Word-Dokument erstellen ───────────────────────────────────────────
        self._btn_erstellen.setEnabled(False)
        self._status_lbl.setText("⏳  Dokument wird erstellt …")

        try:
            from functions.laufzettel_functions import (
                erstelle_laufzettel_word, oeffne_laufzettel
            )
            from datetime import datetime

            daten = {
                "mitarbeiter_name": ma_name,
                "datum":            datetime.now().strftime("%d.%m.%Y"),
                "vorlagen":         vorlagen,
                "schulungen":       schulungen,
                "einarbeitung_typ": einarbeitung_typ,
            }

            pfad = erstelle_laufzettel_word(daten)
            self._status_lbl.setText(f"✅  Erstellt: {os.path.basename(pfad)}")

            antwort = QMessageBox.question(
                self,
                "Laufzettel erstellt",
                f"Der Laufzettel wurde erfolgreich erstellt:\n\n{pfad}\n\n"
                "Möchtest du die Datei jetzt öffnen?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if antwort == QMessageBox.StandardButton.Yes:
                oeffne_laufzettel(pfad)

            self.accept()

        except Exception as e:
            self._status_lbl.setText("❌  Fehler beim Erstellen.")
            QMessageBox.critical(
                self, "Fehler",
                f"Der Laufzettel konnte nicht erstellt werden:\n\n{e}"
            )
        finally:
            self._btn_erstellen.setEnabled(True)
