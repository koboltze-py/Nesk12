"""
Workflow-Widget – Abgleich Stärkemeldungen ↔ Tagesdienstpläne
==============================================================
Lädt beliebig viele Stärkemeldungen (Word .docx) und Tagesdienstpläne
(Excel .xlsx) und vergleicht die enthaltenen Personen / Dienste / Zeiten.
"""
from __future__ import annotations

import re
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QFileDialog, QFrame, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QMessageBox, QPushButton,
    QScrollArea, QSizePolicy, QSplitter, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QProgressBar,
)

from config import BASE_DIR, FIORI_BLUE, FIORI_TEXT

# ── Standardpfade ──────────────────────────────────────────────────────────────
_BASE_ONEDRIVE = Path(BASE_DIR).parent.parent  # …/!Gemeinsam.26/
_DEFAULT_STAERKE  = str(_BASE_ONEDRIVE / "06_Stärkemeldung")
_DEFAULT_DIENSTPLAN = str(_BASE_ONEDRIVE / "04_Tagesdienstpläne")


# ── Stärkemeldung-Parser (Word .docx) ─────────────────────────────────────────

def _parse_staerkemeldung(docx_path: str) -> dict:
    """
    Liest Name, Dienst, Beginn, Ende aus einer Stärkemeldungs-.docx.
    Gibt zurück:
      { 'datei': str, 'datum': str, 'personen': list[dict] }
    Jede Person: { 'vollname': str, 'dienst': str, 'beginn': str, 'ende': str }
    """
    try:
        from docx import Document
    except ImportError:
        return {"datei": docx_path, "datum": "", "personen": [], "fehler": "python-docx fehlt"}

    try:
        doc = Document(docx_path)
    except Exception as e:
        return {"datei": docx_path, "datum": "", "personen": [], "fehler": str(e)}

    personen: list[dict] = []
    datum_str = ""

    # Datum aus Dateiname extrahieren (z.B. Staerkemeldung_2026-06-01.docx)
    m_datum = re.search(r'(\d{4}[-_]\d{2}[-_]\d{2})', Path(docx_path).stem)
    if m_datum:
        datum_str = m_datum.group(1).replace("_", "-")

    # Alle Tabellen durchsuchen
    for tbl in doc.tables:
        # Kopfzeile finden (NAME, DIENST, BEGINN, ENDE)
        header_row_idx = None
        col_map: dict[str, int] = {}
        for ri, row in enumerate(tbl.rows):
            cells_text = [c.text.strip().upper() for c in row.cells]
            has_name   = any("NAME" in t for t in cells_text)
            has_dienst = any("DIENST" in t for t in cells_text)
            if has_name and has_dienst:
                header_row_idx = ri
                for ci, ct in enumerate(cells_text):
                    if "NAME" in ct:
                        col_map["name"]   = ci
                    elif "DIENST" in ct:
                        col_map["dienst"] = ci
                    elif "BEGINN" in ct or "START" in ct or "VON" in ct:
                        col_map["beginn"] = ci
                    elif "ENDE" in ct or "BIS" in ct:
                        col_map["ende"]   = ci
                break

        if header_row_idx is None:
            continue

        # Datenzeilen lesen
        for row in tbl.rows[header_row_idx + 1:]:
            cells = [c.text.strip() for c in row.cells]
            if not cells:
                continue
            name_idx   = col_map.get("name",   0)
            dienst_idx = col_map.get("dienst", 1)
            beginn_idx = col_map.get("beginn", 2)
            ende_idx   = col_map.get("ende",   3)

            name   = cells[name_idx]   if name_idx   < len(cells) else ""
            dienst = cells[dienst_idx] if dienst_idx < len(cells) else ""
            beginn = cells[beginn_idx] if beginn_idx < len(cells) else ""
            ende   = cells[ende_idx]   if ende_idx   < len(cells) else ""

            name = name.strip()
            if not name or name.upper() in ("NAME", ""):
                continue

            personen.append({
                "vollname": name,
                "dienst":   dienst.strip(),
                "beginn":   beginn.strip(),
                "ende":     ende.strip(),
            })

    # Falls kein Datum aus Dateiname, aus Dokumenttext versuchen
    if not datum_str:
        for para in doc.paragraphs:
            m = re.search(r'(\d{1,2}[./]\d{1,2}[./]\d{4})', para.text)
            if m:
                datum_str = m.group(1)
                break

    return {
        "datei":    Path(docx_path).name,
        "datum":    datum_str,
        "personen": personen,
    }


# ── Dienstplan-Parser (Excel .xlsx) ───────────────────────────────────────────

def _parse_dienstplan_fuer_abgleich(xlsx_path: str) -> dict:
    """
    Nutzt den vorhandenen DienstplanParser und gibt eine vereinfachte
    Personen-Liste zurück.
    """
    try:
        from functions.dienstplan_parser import DienstplanParser
        result = DienstplanParser(xlsx_path, alle_anzeigen=True, round_dispo=False).parse()
    except Exception as e:
        return {"datei": xlsx_path, "datum": "", "personen": [], "fehler": str(e)}

    if not result.get("success", True) and result.get("error"):
        return {
            "datei":    Path(xlsx_path).name,
            "datum":    "",
            "personen": [],
            "fehler":   result["error"],
        }

    personen: list[dict] = []
    for p in result.get("betreuer", []) + result.get("dispo", []):
        dienst = p.get("dienst_kuerzel") or p.get("dienst") or ""
        personen.append({
            "vollname": p.get("vollname", "").strip(),
            "dienst":   dienst.strip(),
            "beginn":   (p.get("start_zeit") or "").strip(),
            "ende":     (p.get("end_zeit")   or "").strip(),
            "ist_dispo": bool(p.get("ist_dispo")),
        })

    # Datum aus Dateiname
    datum_str = ""
    m = re.search(r'(\d{4}[-_]\d{2}[-_]\d{2})', Path(xlsx_path).stem)
    if m:
        datum_str = m.group(1).replace("_", "-")

    return {
        "datei":    Path(xlsx_path).name,
        "datum":    datum_str,
        "personen": personen,
    }


# ── Abgleich-Logik ─────────────────────────────────────────────────────────────

def _normiere_name(s: str) -> str:
    return " ".join(s.lower().split())

def _normiere_zeit(s: str) -> str:
    """Normiert Zeiten auf HH:MM (volle Stunde wenn nur Stunde angegeben)."""
    s = s.strip().replace(".", ":")
    if re.fullmatch(r'\d{1,2}', s):
        return f"{int(s):02d}:00"
    m = re.match(r'(\d{1,2}):(\d{2})', s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s

def _normiere_dienst(s: str) -> str:
    return s.strip().upper()


class AbgleichErgebnis:
    """Enthält alle Unterschiede zwischen Stärkemeldung und Dienstplan."""

    def __init__(self, datei_staerke: str, datei_dienstplan: str):
        self.datei_staerke    = datei_staerke
        self.datei_dienstplan = datei_dienstplan
        self.ok:         list[dict] = []   # identisch
        self.abweichung: list[dict] = []   # Unterschied im Dienst/Zeit
        self.nur_staerke: list[dict] = []  # in Stärke, nicht im Dienstplan
        self.nur_dienstplan: list[dict] = []  # im Dienstplan, nicht in Stärke

    @property
    def hat_fehler(self) -> bool:
        return bool(self.abweichung or self.nur_staerke or self.nur_dienstplan)


def _abgleichen(
    staerke_data: dict,
    dienstplan_data: dict,
) -> AbgleichErgebnis:
    erg = AbgleichErgebnis(
        staerke_data.get("datei", ""),
        dienstplan_data.get("datei", ""),
    )

    s_personen = {_normiere_name(p["vollname"]): p
                  for p in staerke_data.get("personen", [])
                  if p.get("vollname")}
    d_personen = {_normiere_name(p["vollname"]): p
                  for p in dienstplan_data.get("personen", [])
                  if p.get("vollname")}

    alle_namen = set(s_personen) | set(d_personen)

    for name in sorted(alle_namen):
        sp = s_personen.get(name)
        dp = d_personen.get(name)

        if sp and dp:
            # Beide vorhanden → Felder vergleichen
            s_dienst = _normiere_dienst(sp.get("dienst", ""))
            d_dienst  = _normiere_dienst(dp.get("dienst", ""))
            s_beginn  = _normiere_zeit(sp.get("beginn", ""))
            d_beginn  = _normiere_zeit(dp.get("beginn", ""))
            s_ende    = _normiere_zeit(sp.get("ende", ""))
            d_ende    = _normiere_zeit(dp.get("ende", ""))

            unterschiede: list[str] = []
            if s_dienst != d_dienst and s_dienst and d_dienst:
                unterschiede.append(f"Dienst: SM={s_dienst} / DP={d_dienst}")
            if s_beginn != d_beginn and s_beginn and d_beginn:
                unterschiede.append(f"Beginn: SM={s_beginn} / DP={d_beginn}")
            if s_ende != d_ende and s_ende and d_ende:
                unterschiede.append(f"Ende: SM={s_ende} / DP={d_ende}")

            eintrag = {
                "name": sp["vollname"],
                "sm_dienst": sp.get("dienst", ""),
                "sm_beginn": sp.get("beginn", ""),
                "sm_ende":   sp.get("ende", ""),
                "dp_dienst": dp.get("dienst", ""),
                "dp_beginn": dp.get("beginn", ""),
                "dp_ende":   dp.get("ende", ""),
                "unterschiede": unterschiede,
            }
            if unterschiede:
                erg.abweichung.append(eintrag)
            else:
                erg.ok.append(eintrag)
        elif sp:
            erg.nur_staerke.append({
                "name":      sp["vollname"],
                "sm_dienst": sp.get("dienst", ""),
                "sm_beginn": sp.get("beginn", ""),
                "sm_ende":   sp.get("ende", ""),
            })
        else:
            erg.nur_dienstplan.append({
                "name":      dp["vollname"],
                "dp_dienst": dp.get("dienst", ""),
                "dp_beginn": dp.get("beginn", ""),
                "dp_ende":   dp.get("ende", ""),
            })

    return erg


# ── Lade-Thread ────────────────────────────────────────────────────────────────

class _LadeThread(QThread):
    fortschritt = Signal(str)
    fertig      = Signal(list)   # list[AbgleichErgebnis]
    fehler      = Signal(str)

    def __init__(self, staerke_pfade: list[str], dienstplan_pfade: list[str]):
        super().__init__()
        self._staerke    = staerke_pfade
        self._dienstplan = dienstplan_pfade

    def run(self):
        try:
            ergebnisse: list[AbgleichErgebnis] = []

            # Daten laden
            self.fortschritt.emit("Stärkemeldungen werden geladen …")
            sm_data: list[dict] = []
            for pfad in self._staerke:
                self.fortschritt.emit(f"  Lese: {Path(pfad).name}")
                sm_data.append(_parse_staerkemeldung(pfad))

            self.fortschritt.emit("Dienstpläne werden geladen …")
            dp_data: list[dict] = []
            for pfad in self._dienstplan:
                self.fortschritt.emit(f"  Lese: {Path(pfad).name}")
                dp_data.append(_parse_dienstplan_fuer_abgleich(pfad))

            # Abgleich: versuche Datum-Matching
            self.fortschritt.emit("Abgleich wird durchgeführt …")

            # Datum-Index der Dienstpläne aufbauen
            dp_by_datum: dict[str, list[dict]] = {}
            for dp in dp_data:
                d = dp.get("datum", "")
                dp_by_datum.setdefault(d, []).append(dp)

            matched_dp: set[int] = set()

            for sm in sm_data:
                sm_datum = sm.get("datum", "")
                candidates = dp_by_datum.get(sm_datum, []) if sm_datum else []
                # noch nicht gematchte Kandidaten bevorzugen
                unmatch = [c for c in candidates if id(c) not in matched_dp]
                if unmatch:
                    dp_match = unmatch[0]
                    matched_dp.add(id(dp_match))
                elif dp_data:
                    # kein Datum-Match → einfach der Reihe nach
                    unmatched_all = [d for d in dp_data if id(d) not in matched_dp]
                    dp_match = unmatched_all[0] if unmatched_all else dp_data[0]
                    matched_dp.add(id(dp_match))
                else:
                    dp_match = {"datei": "–", "datum": "", "personen": []}

                ergebnisse.append(_abgleichen(sm, dp_match))

            # Übriggebliebene Dienstpläne ohne Stärkemeldung
            for dp in dp_data:
                if id(dp) not in matched_dp:
                    sm_leer = {"datei": "–", "datum": "", "personen": []}
                    ergebnisse.append(_abgleichen(sm_leer, dp))

            self.fertig.emit(ergebnisse)
        except Exception as exc:
            self.fehler.emit(str(exc))


# ── Detail-Dialog ──────────────────────────────────────────────────────────────

class _DetailDialog(QDialog):
    def __init__(self, erg: AbgleichErgebnis, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"🔍  Detail: {erg.datei_staerke} ↔ {erg.datei_dienstplan}")
        self.resize(900, 600)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)

        # Statistik-Zeile
        stats = QLabel(
            f"✅ Identisch: {len(erg.ok)}   "
            f"⚠️ Abweichungen: {len(erg.abweichung)}   "
            f"📄 Nur Stärke: {len(erg.nur_staerke)}   "
            f"📋 Nur Dienstplan: {len(erg.nur_dienstplan)}"
        )
        stats.setStyleSheet("font-size: 12px; font-weight: bold; color: #2c3e50; padding: 4px 0;")
        lay.addWidget(stats)

        tbl = QTableWidget()
        tbl.setColumnCount(8)
        tbl.setHorizontalHeaderLabels([
            "Status", "Name",
            "SM Dienst", "SM Beginn", "SM Ende",
            "DP Dienst", "DP Beginn", "DP Ende",
        ])
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tbl.setAlternatingRowColors(True)
        tbl.verticalHeader().setVisible(False)
        tbl.setStyleSheet("font-size: 11px;")

        rows: list[tuple[str, dict, QColor]] = []
        for e in erg.abweichung:
            rows.append(("⚠️ Abweichung", e, QColor("#fff3cd")))
        for e in erg.nur_staerke:
            rows.append(("📄 Nur Stärke", e, QColor("#fde8e8")))
        for e in erg.nur_dienstplan:
            rows.append(("📋 Nur Dienstplan", e, QColor("#e8f4fd")))
        for e in erg.ok:
            rows.append(("✅ Identisch", e, QColor("#f0fff0")))

        tbl.setRowCount(len(rows))
        for ri, (status, e, bg) in enumerate(rows):
            vals = [
                status,
                e.get("name", ""),
                e.get("sm_dienst", ""),
                e.get("sm_beginn", ""),
                e.get("sm_ende",   ""),
                e.get("dp_dienst", ""),
                e.get("dp_beginn", ""),
                e.get("dp_ende",   ""),
            ]
            for ci, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setBackground(bg)
                if ci == 0 and "Abweichung" in status:
                    item.setForeground(QColor("#856404"))
                elif ci == 0 and "Stärke" in status:
                    item.setForeground(QColor("#c0392b"))
                elif ci == 0 and "Dienstplan" in status:
                    item.setForeground(QColor("#1a6ea8"))
                tbl.setItem(ri, ci, item)

            # Unterschiede-Tooltip
            if e.get("unterschiede"):
                for ci in range(8):
                    if tbl.item(ri, ci):
                        tbl.item(ri, ci).setToolTip("\n".join(e["unterschiede"]))

        lay.addWidget(tbl)

        btn = QPushButton("Schließen")
        btn.clicked.connect(self.accept)
        btn.setStyleSheet(
            f"QPushButton{{background:{FIORI_BLUE};color:white;border:none;border-radius:4px;"
            f"padding:6px 18px;font-size:12px;font-weight:bold;}}"
            f"QPushButton:hover{{background:#1a5276;}}"
        )
        lay.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)


# ── Haupt-Widget ───────────────────────────────────────────────────────────────

class WorkflowWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._staerke_pfade:    list[str] = []
        self._dienstplan_pfade: list[str] = []
        self._ergebnisse:       list[AbgleichErgebnis] = []
        self._thread: _LadeThread | None = None
        self._setup_ui()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 10, 16, 10)
        root.setSpacing(10)

        # Titel
        title = QLabel("⚙️  Workflow – Stärkemeldung ↔ Dienstplan Abgleich")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Light))
        title.setStyleSheet(f"color: {FIORI_TEXT};")
        root.addWidget(title)

        hint = QLabel(
            "Lade Stärkemeldungen (Word) und Tagesdienstpläne (Excel) um sie automatisch abzugleichen. "
            "Es wird geprüft ob Namen, Dienste und Zeiten übereinstimmen."
        )
        hint.setStyleSheet("color: #666; font-size: 11px;")
        hint.setWordWrap(True)
        root.addWidget(hint)

        # Dateiauswahl-Bereich
        datei_row = QHBoxLayout()
        datei_row.setSpacing(12)

        datei_row.addWidget(self._build_datei_gruppe(
            "📄  Stärkemeldungen (Word .docx)",
            "staerke",
            _DEFAULT_STAERKE,
            ["Word-Dokumente (*.docx *.doc)"],
        ))
        datei_row.addWidget(self._build_datei_gruppe(
            "📋  Tagesdienstpläne (Excel .xlsx)",
            "dienstplan",
            _DEFAULT_DIENSTPLAN,
            ["Excel-Tabellen (*.xlsx *.xls)"],
        ))
        root.addLayout(datei_row)

        # Aktions-Zeile
        btn_row = QHBoxLayout()
        self._btn_start = QPushButton("🔄  Abgleich starten")
        self._btn_start.setMinimumHeight(40)
        self._btn_start.setEnabled(False)
        self._btn_start.setStyleSheet(self._btn_style(FIORI_BLUE, "#1a5276"))
        self._btn_start.clicked.connect(self._starte_abgleich)
        btn_row.addWidget(self._btn_start)

        self._btn_reset = QPushButton("🗑  Zurücksetzen")
        self._btn_reset.setMinimumHeight(40)
        self._btn_reset.setStyleSheet(self._btn_style("#7f8c8d", "#5d6d7e"))
        self._btn_reset.clicked.connect(self._reset)
        btn_row.addWidget(self._btn_reset)
        btn_row.addStretch()
        root.addLayout(btn_row)

        # Fortschrittsbalken
        self._progress = QProgressBar()
        self._progress.setVisible(False)
        self._progress.setRange(0, 0)
        self._progress.setFixedHeight(6)
        self._progress.setStyleSheet(
            f"QProgressBar{{border:none;border-radius:3px;background:#e0e0e0;}}"
            f"QProgressBar::chunk{{background:{FIORI_BLUE};border-radius:3px;}}"
        )
        root.addWidget(self._progress)

        self._status_lbl = QLabel("")
        self._status_lbl.setStyleSheet("color: #555; font-size: 11px;")
        root.addWidget(self._status_lbl)

        # Ergebnis-Tabelle
        grp_erg = QGroupBox("📊  Abgleich-Ergebnisse")
        grp_erg.setStyleSheet(
            f"QGroupBox{{border:1px solid #c8d0d8;border-radius:6px;margin-top:14px;"
            f"padding-top:6px;font-size:11px;font-weight:bold;color:{FIORI_TEXT};}}"
            f"QGroupBox::title{{subcontrol-origin:margin;subcontrol-position:top left;"
            f"left:10px;padding:0 6px;color:{FIORI_BLUE};}}"
        )
        erg_lay = QVBoxLayout(grp_erg)

        self._erg_table = QTableWidget()
        self._erg_table.setColumnCount(6)
        self._erg_table.setHorizontalHeaderLabels([
            "Stärkemeldung", "Dienstplan",
            "✅ Identisch", "⚠️ Abweichung",
            "📄 Nur Stärke", "📋 Nur Dienstplan",
        ])
        self._erg_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._erg_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for ci in range(2, 6):
            self._erg_table.horizontalHeader().setSectionResizeMode(
                ci, QHeaderView.ResizeMode.ResizeToContents
            )
        self._erg_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._erg_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._erg_table.setAlternatingRowColors(True)
        self._erg_table.verticalHeader().setVisible(False)
        self._erg_table.setStyleSheet("font-size: 11px;")
        self._erg_table.doubleClicked.connect(self._detail_zeigen)
        erg_lay.addWidget(self._erg_table)

        hint_detail = QLabel("💡  Doppelklick auf eine Zeile für Detailansicht")
        hint_detail.setStyleSheet("color: #888; font-size: 10px;")
        erg_lay.addWidget(hint_detail)

        root.addWidget(grp_erg, stretch=1)

    def _build_datei_gruppe(
        self,
        titel: str,
        key: str,
        default_dir: str,
        filter_list: list[str],
    ) -> QGroupBox:
        grp = QGroupBox(titel)
        grp.setStyleSheet(
            f"QGroupBox{{border:1px solid #c8d0d8;border-radius:6px;margin-top:14px;"
            f"padding-top:6px;font-size:11px;font-weight:bold;color:{FIORI_TEXT};}}"
            f"QGroupBox::title{{subcontrol-origin:margin;subcontrol-position:top left;"
            f"left:10px;padding:0 6px;color:{FIORI_BLUE};}}"
        )
        lay = QVBoxLayout(grp)
        lay.setSpacing(4)

        # Dateiliste
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setFixedHeight(110)
        scroll.setStyleSheet("background: #fafafa; border: 1px solid #e0e0e0; border-radius: 4px;")
        container = QWidget()
        container.setStyleSheet("background: transparent;")
        vlay = QVBoxLayout(container)
        vlay.setContentsMargins(4, 4, 4, 4)
        vlay.setSpacing(2)
        scroll.setWidget(container)
        lay.addWidget(scroll)

        lbl_leer = QLabel("Keine Dateien ausgewählt")
        lbl_leer.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic;")
        vlay.addWidget(lbl_leer)
        vlay.addStretch()

        # Buttons
        btn_lay = QHBoxLayout()
        btn_laden = QPushButton("📂  Dateien laden")
        btn_laden.setMinimumHeight(32)
        btn_laden.setStyleSheet(self._btn_style("#2980b9", "#1a6ea8"))
        btn_laden.clicked.connect(lambda: self._lade_dateien(key, default_dir, filter_list, vlay, lbl_leer))
        btn_lay.addWidget(btn_laden)

        btn_clear = QPushButton("✕")
        btn_clear.setFixedSize(32, 32)
        btn_clear.setToolTip("Auswahl leeren")
        btn_clear.setStyleSheet(self._btn_style("#e74c3c", "#c0392b"))
        btn_clear.clicked.connect(lambda: self._clear_dateien(key, vlay, lbl_leer))
        btn_lay.addWidget(btn_clear)
        lay.addLayout(btn_lay)

        return grp

    # ── Dateiauswahl ────────────────────────────────────────────────────────────

    def _lade_dateien(
        self,
        key: str,
        default_dir: str,
        filter_list: list[str],
        vlay: QVBoxLayout,
        lbl_leer: QLabel,
    ):
        # Standard-Dir anlegen falls nicht vorhanden
        import os
        start = default_dir if os.path.isdir(default_dir) else str(Path.home())

        dateien, _ = QFileDialog.getOpenFileNames(
            self,
            "Dateien auswählen",
            start,
            ";;".join(filter_list) + ";;Alle Dateien (*)",
        )
        if not dateien:
            return

        if key == "staerke":
            self._staerke_pfade = dateien
        else:
            self._dienstplan_pfade = dateien

        self._aktualisiere_dateiliste(vlay, lbl_leer, dateien)
        self._update_start_btn()

    def _clear_dateien(self, key: str, vlay: QVBoxLayout, lbl_leer: QLabel):
        if key == "staerke":
            self._staerke_pfade = []
        else:
            self._dienstplan_pfade = []
        self._aktualisiere_dateiliste(vlay, lbl_leer, [])
        self._update_start_btn()

    def _aktualisiere_dateiliste(self, vlay: QVBoxLayout, lbl_leer: QLabel, dateien: list[str]):
        # Alle alten Widgets entfernen
        while vlay.count():
            item = vlay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not dateien:
            lbl_leer2 = QLabel("Keine Dateien ausgewählt")
            lbl_leer2.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic;")
            vlay.addWidget(lbl_leer2)
        else:
            for pfad in dateien:
                lbl = QLabel(f"📄  {Path(pfad).name}")
                lbl.setStyleSheet("color: #2c3e50; font-size: 11px;")
                lbl.setToolTip(pfad)
                vlay.addWidget(lbl)
        vlay.addStretch()

    def _update_start_btn(self):
        self._btn_start.setEnabled(
            bool(self._staerke_pfade) and bool(self._dienstplan_pfade)
        )

    # ── Abgleich ────────────────────────────────────────────────────────────────

    def _starte_abgleich(self):
        if self._thread and self._thread.isRunning():
            return

        self._erg_table.setRowCount(0)
        self._progress.setVisible(True)
        self._status_lbl.setText("Wird geladen …")
        self._btn_start.setEnabled(False)

        self._thread = _LadeThread(self._staerke_pfade, self._dienstplan_pfade)
        self._thread.fortschritt.connect(self._status_lbl.setText)
        self._thread.fertig.connect(self._zeige_ergebnisse)
        self._thread.fehler.connect(self._zeige_fehler)
        self._thread.start()

    def _zeige_ergebnisse(self, ergebnisse: list):
        self._ergebnisse = ergebnisse
        self._progress.setVisible(False)
        self._btn_start.setEnabled(True)
        self._status_lbl.setText(
            f"Abgleich abgeschlossen – {len(ergebnisse)} Paar(e) verglichen."
        )

        self._erg_table.setRowCount(len(ergebnisse))
        for ri, erg in enumerate(ergebnisse):
            hat_fehler = erg.hat_fehler
            bg = QColor("#fff3cd") if hat_fehler else QColor("#f0fff0")

            vals = [
                erg.datei_staerke,
                erg.datei_dienstplan,
                str(len(erg.ok)),
                str(len(erg.abweichung)),
                str(len(erg.nur_staerke)),
                str(len(erg.nur_dienstplan)),
            ]
            for ci, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setBackground(bg)
                if ci == 3 and erg.abweichung:
                    item.setForeground(QColor("#856404"))
                    item.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
                if ci == 4 and erg.nur_staerke:
                    item.setForeground(QColor("#c0392b"))
                if ci == 5 and erg.nur_dienstplan:
                    item.setForeground(QColor("#1a6ea8"))
                self._erg_table.setItem(ri, ci, item)

        if ergebnisse:
            gesamt_ok   = sum(len(e.ok)          for e in ergebnisse)
            gesamt_abw  = sum(len(e.abweichung)  for e in ergebnisse)
            gesamt_ns   = sum(len(e.nur_staerke) for e in ergebnisse)
            gesamt_nd   = sum(len(e.nur_dienstplan) for e in ergebnisse)
            self._status_lbl.setText(
                f"Fertig – Gesamt: ✅ {gesamt_ok} identisch  "
                f"⚠️ {gesamt_abw} Abweichung(en)  "
                f"📄 {gesamt_ns} nur Stärke  "
                f"📋 {gesamt_nd} nur Dienstplan"
            )

    def _zeige_fehler(self, msg: str):
        self._progress.setVisible(False)
        self._btn_start.setEnabled(True)
        self._status_lbl.setText("Fehler aufgetreten.")
        QMessageBox.critical(self, "Fehler beim Abgleich", msg)

    def _detail_zeigen(self, index):
        ri = index.row()
        if 0 <= ri < len(self._ergebnisse):
            dlg = _DetailDialog(self._ergebnisse[ri], self)
            dlg.exec()

    def _reset(self):
        self._staerke_pfade = []
        self._dienstplan_pfade = []
        self._ergebnisse = []
        self._erg_table.setRowCount(0)
        self._status_lbl.setText("")
        self._update_start_btn()
        # Dateilisten-Labels zurücksetzen
        self._setup_ui.__func__  # keine neue UI, nur Status reset

    # ── Styles ─────────────────────────────────────────────────────────────────

    def _btn_style(self, bg: str, hover: str) -> str:
        return (
            f"QPushButton{{background:{bg};color:white;border:none;border-radius:4px;"
            f"padding:5px 14px;font-size:11px;font-weight:bold;}}"
            f"QPushButton:hover{{background:{hover};}}"
            f"QPushButton:disabled{{background:#ccc;color:#888;}}"
        )

    def refresh(self):
        pass
