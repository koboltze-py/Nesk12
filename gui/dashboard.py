"""
Dashboard-Widget
Zeigt Statistiken, Kalender und Fahrzeug-Termine
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QGridLayout,
    QMessageBox, QCalendarWidget, QScrollArea, QSizePolicy
)
from PySide6.QtCore import Qt, QTimer, QDate, QTime, QRect
from PySide6.QtGui import QFont, QPainter, QLinearGradient, QColor, QTextCharFormat, QBrush

from config import FIORI_BLUE, FIORI_TEXT, FIORI_WHITE, FIORI_SUCCESS, FIORI_WARNING


class _TerminKalender(QCalendarWidget):
    """QCalendarWidget mit kleinem farbigen Punkt für Tage mit Terminen/Notizen."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._termin_dates: set[str] = set()  # 'YYYY-MM-DD'  → blauer Punkt
        self._notiz_dates:  set[str] = set()  # 'YYYY-MM-DD'  → grüner Punkt

    def set_termin_dates(self, dates: set[str]):
        self._termin_dates = dates
        self.updateCells()

    def set_notiz_dates(self, dates: set[str]):
        self._notiz_dates = dates
        self.updateCells()

    def paintCell(self, painter: QPainter, rect: QRect, date: QDate):
        super().paintCell(painter, rect, date)
        datum_str = date.toString("yyyy-MM-dd")
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        dot_r = 3
        cy = rect.bottom() - dot_r - 2
        cx = rect.center().x()
        # Fahrzeug-Termin: blauer Punkt (links)
        if datum_str in self._termin_dates and datum_str in self._notiz_dates:
            painter.setBrush(QBrush(QColor("#1565c0")))
            painter.drawEllipse(cx - dot_r - 5, cy - dot_r, dot_r * 2, dot_r * 2)
            painter.setBrush(QBrush(QColor("#2e7d32")))
            painter.drawEllipse(cx + 1, cy - dot_r, dot_r * 2, dot_r * 2)
        elif datum_str in self._termin_dates:
            painter.setBrush(QBrush(QColor("#1565c0")))
            painter.drawEllipse(cx - dot_r, cy - dot_r, dot_r * 2, dot_r * 2)
        elif datum_str in self._notiz_dates:
            painter.setBrush(QBrush(QColor("#2e7d32")))
            painter.drawEllipse(cx - dot_r, cy - dot_r, dot_r * 2, dot_r * 2)
        painter.restore()


class StatCard(QFrame):
    """Eine Statistik-Karte im SAP Fiori-Stil."""
    def __init__(self, title: str, value: str, icon: str, color: str, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 4px solid {color};
            }}
        """)
        self.setMinimumHeight(110)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)

        top = QHBoxLayout()
        title_lbl = QLabel(title)
        title_lbl.setFont(QFont("Arial", 11))
        title_lbl.setStyleSheet("color: #666; border: none;")
        icon_lbl = QLabel(icon)
        icon_lbl.setFont(QFont("Arial", 20))
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        top.addWidget(title_lbl)
        top.addStretch()
        top.addWidget(icon_lbl)
        layout.addLayout(top)

        self._value_lbl = QLabel(value)
        self._value_lbl.setFont(QFont("Arial", 28, QFont.Weight.Bold))
        self._value_lbl.setStyleSheet(f"color: {color}; border: none;")
        layout.addWidget(self._value_lbl)

    def set_value(self, value: str):
        self._value_lbl.setText(value)


# ---------------------------------------------------------------------------
# Animierter Himmel (internes Widget für FlugzeugWidget)
# ---------------------------------------------------------------------------
class _SkyWidget(QWidget):
    """Himmel-Strip mit animiertem Flugzeug via QPainter + QTimer."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._x: float = -60.0
        self._speed: float = 1.8          # Pixel pro Frame (~30 fps)
        self.setFixedHeight(72)

        self._anim_timer = QTimer(self)
        self._anim_timer.timeout.connect(self._step)
        self._anim_timer.start(30)        # ~33 FPS

    def _step(self):
        self._x += self._speed
        if self._x > self.width() + 60:
            self._x = -60.0
        self.update()

    def paintEvent(self, event):  # noqa: N802
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Himmel-Verlauf
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0.0, QColor("#5BA3D0"))
        grad.setColorAt(1.0, QColor("#A8D8F0"))
        p.fillRect(self.rect(), grad)

        # Wolken (links)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(255, 255, 255, 190))
        p.drawEllipse(18,  6, 54, 26)
        p.drawEllipse(10, 15, 44, 20)
        p.drawEllipse(52,  8, 38, 22)

        # Wolken (rechts)
        w = self.width()
        p.drawEllipse(w - 130, 10, 58, 24)
        p.drawEllipse(w - 140, 18, 46, 18)
        p.drawEllipse(w - 100,  6, 42, 22)

        # Rollbahn unten
        p.setBrush(QColor(130, 130, 130, 120))
        p.drawRect(0, self.height() - 13, w, 13)
        p.setBrush(QColor(255, 255, 255, 210))
        for i in range(0, w, 32):
            p.drawRect(i + 4, self.height() - 9, 16, 4)

        # Flugzeug-Emoji
        font = QFont("Segoe UI Emoji", 22)
        p.setFont(font)
        p.setPen(QColor(30, 30, 30))
        p.drawText(int(self._x), 50, "✈")

        p.end()


# ---------------------------------------------------------------------------
# Flugzeug-Karte (klickbar)
# ---------------------------------------------------------------------------
class FlugzeugWidget(QFrame):
    """Animiertes Flugzeug mit Verspätungs-Uhr. Klickbar."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._delay_min = 0
        self._delay_sec = 0
        self._build()

    def _build(self):
        self.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 4px solid {FIORI_BLUE};
            }}
        """)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 10)
        layout.setSpacing(8)

        # Header-Zeile
        header = QHBoxLayout()
        title = QLabel("✈  Flughafen Köln/Bonn  –  Live Ansicht")
        title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        title.setStyleSheet("color: #333; border: none;")
        hint = QLabel("zum Klicken")
        hint.setFont(QFont("Segoe UI", 9))
        hint.setStyleSheet("color: #aaa; border: none;")
        header.addWidget(title)
        header.addStretch()
        header.addWidget(hint)
        layout.addLayout(header)

        # Animierter Himmel
        self._sky = _SkyWidget(self)
        layout.addWidget(self._sky)

        # Verspätungs-Anzeige
        bottom = QHBoxLayout()
        clock_icon = QLabel("🕐")
        clock_icon.setFont(QFont("Segoe UI Emoji", 16))
        clock_icon.setStyleSheet("border: none;")
        versp_lbl = QLabel("Aktuelle Verspätung:")
        versp_lbl.setFont(QFont("Segoe UI", 10))
        versp_lbl.setStyleSheet("color: #555; border: none;")
        self._delay_lbl = QLabel("00:00 min")
        self._delay_lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        self._delay_lbl.setStyleSheet("color: #bb0000; border: none;")
        bottom.addWidget(clock_icon)
        bottom.addSpacing(4)
        bottom.addWidget(versp_lbl)
        bottom.addStretch()
        bottom.addWidget(self._delay_lbl)
        layout.addLayout(bottom)

        # Uhr-Timer
        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._tick)
        self._clock_timer.start(1000)

    def _tick(self):
        self._delay_sec += 1
        if self._delay_sec >= 60:
            self._delay_sec = 0
            self._delay_min += 1
        self._delay_lbl.setText(f"{self._delay_min:02d}:{self._delay_sec:02d} min")

    def mousePressEvent(self, event):  # noqa: N802
        QMessageBox.information(
            self,
            "✈  Reisebüro Nesk3",
            f"Willkommen am Flughafen Köln/Bonn! ✈\n\n"
            f"Aktuelle Verspätung: {self._delay_min:02d}:{self._delay_sec:02d} min\n\n"
            f"Keine Sorge – das Flugzeug landet bestimmt irgendwann! 😄",
        )
        super().mousePressEvent(event)


# ---------------------------------------------------------------------------
class DashboardWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._termine: list[dict] = []
        self._build_ui()
        # Uhr-Timer
        self._uhr_timer = QTimer(self)
        self._uhr_timer.timeout.connect(self._uhr_tick)
        self._uhr_timer.start(1000)
        self._uhr_tick()

        # Dienstplan Auto-Refresh alle 5 Minuten
        self._dp_timer = QTimer(self)
        self._dp_timer.timeout.connect(self._aktualisiere_dp_panels)
        self._dp_timer.start(5 * 60 * 1000)

        # Krankmeldungen Auto-Refresh alle 10 Minuten
        self._km_timer = QTimer(self)
        self._km_timer.timeout.connect(self._lade_krankmeldungen)
        self._km_timer.start(10 * 60 * 1000)

    # ── UI-Aufbau ─────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(20)

        # ── Linke Seite: Kalender + Termine ───────────────────────────────
        linke = QVBoxLayout()
        linke.setSpacing(12)

        titel = QLabel("🏠  Dashboard")
        titel.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        titel.setStyleSheet(f"color: {FIORI_TEXT};")
        linke.addWidget(titel)

        sub = QLabel("Willkommen bei Nesk3 – DRK Flughafen Köln")
        sub.setFont(QFont("Arial", 11))
        sub.setStyleSheet("color: #888;")
        linke.addWidget(sub)

        # Kalender
        self._kalender = _TerminKalender()
        self._kalender.setGridVisible(True)
        self._kalender.setNavigationBarVisible(True)
        self._kalender.setMinimumHeight(280)
        self._kalender.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self._kalender.setStyleSheet("""
            QCalendarWidget {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 8px;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #C8102E;
                border-radius: 8px 8px 0 0;
            }
            QCalendarWidget QToolButton {
                color: white;
                background: transparent;
                border: none;
                font-size: 13px;
                font-weight: bold;
                padding: 4px 8px;
            }
            QCalendarWidget QToolButton:hover {
                background: rgba(255,255,255,51);
                border-radius: 4px;
            }
            QCalendarWidget QSpinBox {
                color: white;
                background: transparent;
                border: none;
                font-size: 13px;
                font-weight: bold;
            }
            QCalendarWidget QAbstractItemView {
                background-color: white;
                color: #333;
                selection-background-color: #0078D4;
                selection-color: white;
                font-size: 12px;
            }
            QCalendarWidget QAbstractItemView:disabled {
                color: #bbb;
            }
        """)
        self._kalender.activated.connect(self._kalender_tag_geklickt)
        linke.addWidget(self._kalender)

        # ── Termin-Kachel + Vorfeld-Kachel nebeneinander ─────────────────
        termin_vorfeld_row = QHBoxLayout()
        termin_vorfeld_row.setSpacing(10)

        # Linke Hälfte: Bevorstehende Fahrzeug-Termine
        termin_vbox = QVBoxLayout()
        termin_vbox.setSpacing(4)
        termin_hdr = QLabel("📌  Bevorstehende Fahrzeug-Termine")
        termin_hdr.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        termin_hdr.setStyleSheet(f"color: {FIORI_TEXT};")
        termin_vbox.addWidget(termin_hdr)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setMinimumHeight(140)
        scroll.setMaximumHeight(200)
        scroll.setStyleSheet("background: transparent;")
        self._termin_container = QWidget()
        self._termin_container.setStyleSheet("background: transparent;")
        self._termin_layout = QVBoxLayout(self._termin_container)
        self._termin_layout.setSpacing(5)
        self._termin_layout.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(self._termin_container)
        termin_vbox.addWidget(scroll)
        termin_vorfeld_row.addLayout(termin_vbox, 1)

        # Rechte Hälfte: Vorfeldmitarbeiter aus Sonderaufgaben-Historie
        vorfeld_vbox = QVBoxLayout()
        vorfeld_vbox.setSpacing(3)
        vorfeld_hdr = QLabel("👷  Vorfeldmitarbeiter")
        vorfeld_hdr.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        vorfeld_hdr.setStyleSheet(f"color: {FIORI_TEXT};")
        vorfeld_vbox.addWidget(vorfeld_hdr)

        self._vorfeld_kachel_datum_lbl = QLabel("")
        self._vorfeld_kachel_datum_lbl.setStyleSheet(
            "color: #888; font-size: 10px; font-style: italic;"
        )
        vorfeld_vbox.addWidget(self._vorfeld_kachel_datum_lbl)

        # Kein ScrollArea – direkt eingebettetes Widget ohne Scrollen
        self._vorfeld_kachel_container = QWidget()
        self._vorfeld_kachel_container.setStyleSheet("background: transparent;")
        self._vorfeld_kachel_layout = QVBoxLayout(self._vorfeld_kachel_container)
        self._vorfeld_kachel_layout.setSpacing(2)
        self._vorfeld_kachel_layout.setContentsMargins(0, 0, 0, 0)
        vorfeld_vbox.addWidget(self._vorfeld_kachel_container)
        vorfeld_vbox.addStretch()
        termin_vorfeld_row.addLayout(vorfeld_vbox, 1)

        linke.addLayout(termin_vorfeld_row)

        # Krankmeldungen aktueller Monat
        from PySide6.QtWidgets import QPushButton as _QPB2
        km_hdr_row = QHBoxLayout()
        km_hdr = QLabel("🤒  Krankmeldungen – aktueller Monat")
        km_hdr.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        km_hdr.setStyleSheet(f"color: {FIORI_TEXT};")
        km_hdr_row.addWidget(km_hdr)
        km_hdr_row.addStretch()
        self._km_ref_btn = _QPB2("🔄")
        self._km_ref_btn.setToolTip("Krankmeldungen aktualisieren")
        self._km_ref_btn.setFixedSize(26, 22)
        self._km_ref_btn.setStyleSheet(
            "QPushButton { background: #e8f0fe; border: 1px solid #c5d5f5; border-radius: 4px; font-size: 11px; }"
            " QPushButton:hover { background: #c5d5f5; }"
        )
        self._km_ref_btn.clicked.connect(self._lade_krankmeldungen)
        km_hdr_row.addWidget(self._km_ref_btn)
        linke.addLayout(km_hdr_row)

        self._km_status_lbl = QLabel("Wird geladen …")
        self._km_status_lbl.setStyleSheet("color: #888; font-size: 11px; font-style: italic;")
        linke.addWidget(self._km_status_lbl)

        from PySide6.QtWidgets import QTableWidget, QHeaderView as _QHV
        self._km_table = QTableWidget()
        self._km_table.setColumnCount(4)
        self._km_table.setHorizontalHeaderLabels(["Name", "Von", "Bis", "Bemerkungen"])
        km_hh = self._km_table.horizontalHeader()
        km_hh.setSectionResizeMode(0, _QHV.ResizeMode.ResizeToContents)
        km_hh.setSectionResizeMode(1, _QHV.ResizeMode.ResizeToContents)
        km_hh.setSectionResizeMode(2, _QHV.ResizeMode.ResizeToContents)
        km_hh.setSectionResizeMode(3, _QHV.ResizeMode.Stretch)
        self._km_table.verticalHeader().setVisible(False)
        self._km_table.verticalHeader().setDefaultSectionSize(18)
        self._km_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._km_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._km_table.setAlternatingRowColors(True)
        self._km_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border-radius: 6px;
                font-size: 12px;
                gridline-color: #e8ecf0;
            }
            QTableWidget::item { padding: 0px 4px; }
            QHeaderView::section {
                background-color: #f0f4f8;
                border: none;
                border-bottom: 1px solid #c8d4e0;
                padding: 2px 4px;
                font-size: 11px;
                font-weight: bold;
            }
        """)
        self._km_table.setMinimumHeight(120)
        self._km_table.setMaximumHeight(280)
        linke.addWidget(self._km_table)

        # ── Eigene Notizen ─────────────────────────────────────────────────
        notiz_hdr_row = QHBoxLayout()
        notiz_hdr = QLabel("📝  Eigene Notizen")
        notiz_hdr.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        notiz_hdr.setStyleSheet(f"color: {FIORI_TEXT};")
        notiz_hdr_row.addWidget(notiz_hdr)
        notiz_hdr_row.addStretch()
        from PySide6.QtWidgets import QPushButton as _QPB3
        self._notiz_neu_btn = _QPB3("➕  Neue Notiz")
        self._notiz_neu_btn.setToolTip("Neue persönliche Notiz erstellen")
        self._notiz_neu_btn.setFixedHeight(26)
        self._notiz_neu_btn.setStyleSheet(
            "QPushButton { font-size: 11px; padding: 2px 10px; "
            "background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; "
            "border-radius: 4px; } "
            "QPushButton:hover { background: #c8e6c9; }"
        )
        self._notiz_neu_btn.clicked.connect(self._neue_notiz_dialog)
        notiz_hdr_row.addWidget(self._notiz_neu_btn)

        self._notiz_archiv_btn = _QPB3("🗄️  Archiv")
        self._notiz_archiv_btn.setToolTip("Alle Notizen im Archiv anzeigen")
        self._notiz_archiv_btn.setFixedHeight(26)
        self._notiz_archiv_btn.setStyleSheet(
            "QPushButton { font-size: 11px; padding: 2px 10px; "
            "background: #ede7f6; color: #4527a0; border: 1px solid #b39ddb; "
            "border-radius: 4px; } "
            "QPushButton:hover { background: #d1c4e9; }"
        )
        self._notiz_archiv_btn.clicked.connect(self._notiz_archiv_dialog)
        notiz_hdr_row.addWidget(self._notiz_archiv_btn)
        linke.addLayout(notiz_hdr_row)

        self._notiz_scroll = QScrollArea()
        self._notiz_scroll.setWidgetResizable(True)
        self._notiz_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._notiz_scroll.setMinimumHeight(100)
        self._notiz_scroll.setMaximumHeight(240)
        self._notiz_scroll.setStyleSheet("background: transparent;")
        self._notiz_container = QWidget()
        self._notiz_container.setStyleSheet("background: transparent;")
        self._notiz_vlayout = QVBoxLayout(self._notiz_container)
        self._notiz_vlayout.setSpacing(5)
        self._notiz_vlayout.setContentsMargins(0, 0, 0, 0)
        self._notiz_scroll.setWidget(self._notiz_container)
        linke.addWidget(self._notiz_scroll)

        linke.addStretch()
        outer.addLayout(linke, 6)

        # ── Rechte Seite: Uhr + Statistiken + DB-Status ───────────────────
        rechte = QVBoxLayout()
        rechte.setSpacing(12)

        # Digitaluhr
        uhr_frame = QFrame()
        uhr_frame.setStyleSheet("""
            QFrame {
                background: #354a5e;
                border-radius: 10px;
            }
        """)
        uhr_vlayout = QVBoxLayout(uhr_frame)
        uhr_vlayout.setContentsMargins(16, 14, 16, 14)
        uhr_vlayout.setSpacing(2)
        self._uhr_lbl = QLabel("00:00:00")
        self._uhr_lbl.setFont(QFont("Arial", 36, QFont.Weight.Bold))
        self._uhr_lbl.setStyleSheet("color: white; border: none;")
        self._uhr_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        uhr_vlayout.addWidget(self._uhr_lbl)
        self._datum_lbl = QLabel()
        self._datum_lbl.setFont(QFont("Arial", 11))
        self._datum_lbl.setStyleSheet("color: #a0b4c8; border: none;")
        self._datum_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        uhr_vlayout.addWidget(self._datum_lbl)
        rechte.addWidget(uhr_frame)

        # DB-Statusanzeige
        self._db_status_lbl = QLabel("🔄 Datenbankverbindung wird geprüft...")
        self._db_status_lbl.setFont(QFont("Arial", 10))
        self._db_status_lbl.setStyleSheet(
            "background-color: white; border-radius: 6px; padding: 8px 12px;"
        )
        rechte.addWidget(self._db_status_lbl)

        # PAX-Statistik-Karten
        pax_row = QHBoxLayout()
        pax_row.setSpacing(10)
        self._card_pax_jahr = StatCard(
            f"PAX {QDate.currentDate().year()}", "–", "✈", "#0078D4"
        )
        self._card_pax_gestern = StatCard("PAX Vortag", "–", "📊", "#107c10")
        pax_row.addWidget(self._card_pax_jahr)
        pax_row.addWidget(self._card_pax_gestern)
        rechte.addLayout(pax_row)

        # Dienstplan-Panels (heute, im Nachtdienst auch gestern)
        from PySide6.QtWidgets import QPushButton, QTableWidget, QHeaderView

        _TBL_STYLE = """
            QTableWidget {
                background-color: white;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
                gridline-color: #e8ecf0;
            }
            QTableWidget::item { padding: 0px 4px; }
            QHeaderView::section {
                background-color: #f0f4f8;
                border: none;
                border-bottom: 1px solid #c8d4e0;
                padding: 2px 4px;
                font-size: 12px;
                font-weight: bold;
            }
        """

        def _make_dp_panel(titel: str) -> QWidget:
            panel = QWidget()
            pv = QVBoxLayout(panel)
            pv.setContentsMargins(0, 0, 0, 0)
            pv.setSpacing(4)
            hdr_row = QHBoxLayout()
            hdr_row.setContentsMargins(0, 0, 0, 0)
            hdr_lbl = QLabel(titel)
            hdr_lbl.setFont(QFont("Arial", 11, QFont.Weight.Bold))
            hdr_lbl.setStyleSheet(f"color: {FIORI_TEXT};")
            hdr_row.addWidget(hdr_lbl)
            hdr_row.addStretch()
            ref_btn = QPushButton("🔄 Aktualisieren")
            ref_btn.setFixedHeight(26)
            ref_btn.setStyleSheet(
                "QPushButton { font-size: 11px; padding: 2px 10px; "
                "background: #e8f0fa; color: #1565a8; border: 1px solid #b3c8e8; "
                "border-radius: 4px; } "
                "QPushButton:hover { background: #c8daf5; } "
                "QPushButton:pressed { background: #a8c0e8; }"
            )
            hdr_row.addWidget(ref_btn)
            pv.addLayout(hdr_row)
            status_lbl = QLabel("Wird geladen …")
            status_lbl.setStyleSheet("color: #888; font-size: 11px; font-style: italic;")
            pv.addWidget(status_lbl)
            tbl = QTableWidget()
            tbl.setColumnCount(5)
            tbl.setHorizontalHeaderLabels(["Kategorie", "Name", "Dienst", "Von", "Bis"])
            hh_tbl = tbl.horizontalHeader()
            hh_tbl.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            hh_tbl.setStretchLastSection(False)
            hh_tbl.setMinimumSectionSize(40)
            tbl.verticalHeader().setVisible(False)
            tbl.verticalHeader().setDefaultSectionSize(18)
            tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            tbl.setAlternatingRowColors(True)
            tbl.setStyleSheet(_TBL_STYLE)
            pv.addWidget(tbl, 1)
            count_lbl = QLabel("")
            count_lbl.setFont(QFont("Arial", 9))
            count_lbl.setWordWrap(True)
            count_lbl.setStyleSheet("color: #555; font-weight: bold; padding: 2px 0;")
            pv.addWidget(count_lbl)
            panel._ref_btn    = ref_btn
            panel._status_lbl = status_lbl
            panel._table      = tbl
            panel._count_lbl  = count_lbl
            return panel

        dp_area = QHBoxLayout()
        dp_area.setContentsMargins(0, 0, 0, 0)
        dp_area.setSpacing(12)

        # Gestern links (älteres Datum), Heute rechts — gestern im Normalfall versteckt
        self._dp_panel_gestern = _make_dp_panel("📅  Gestriger Dienstplan")
        self._dp_panel_gestern._ref_btn.clicked.connect(self._aktualisiere_dp_panels)
        self._dp_panel_gestern.setVisible(False)
        dp_area.addWidget(self._dp_panel_gestern, 1)

        self._dp_panel_heute = _make_dp_panel("📅  Heutiger Dienstplan")
        self._dp_panel_heute._ref_btn.clicked.connect(self._aktualisiere_dp_panels)
        dp_area.addWidget(self._dp_panel_heute, 1)

        rechte.addLayout(dp_area, 1)

        outer.addLayout(rechte, 4)

    # ── Uhr ───────────────────────────────────────────────────────────────

    def _uhr_tick(self):
        now = QTime.currentTime()
        self._uhr_lbl.setText(now.toString("HH:mm:ss"))
        today = QDate.currentDate()
        _WOCHENTAGE = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"]
        _MONATE = ["","Januar","Februar","März","April","Mai","Juni",
                   "Juli","August","September","Oktober","November","Dezember"]
        wd = _WOCHENTAGE[today.dayOfWeek() - 1]
        mo = _MONATE[today.month()]
        self._datum_lbl.setText(f"{wd}, {today.day()}. {mo} {today.year()}")

    # ── Fahrzeug-Termine laden ────────────────────────────────────────────

    def _lade_fahrzeug_termine(self) -> list[dict]:
        try:
            from database.connection import db_cursor
            with db_cursor() as cur:
                cur.execute("""
                    SELECT ft.id, ft.datum, ft.uhrzeit, ft.typ, ft.titel,
                           ft.beschreibung, f.kennzeichen, f.typ AS fzg_typ
                    FROM fahrzeug_termine ft
                    JOIN fahrzeuge f ON f.id = ft.fahrzeug_id
                    WHERE ft.datum >= date('now') AND ft.erledigt = 0
                    ORDER BY ft.datum, ft.uhrzeit
                    LIMIT 30
                """)
                return cur.fetchall()
        except Exception:
            return []

    # ── Kalender-Klick ────────────────────────────────────────────────────

    def _kalender_tag_geklickt(self, datum: QDate):
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
            QScrollArea, QPushButton as _QB, QSizePolicy,
        )
        datum_str    = datum.toString("yyyy-MM-dd")
        datum_de_str = datum.toString("dd.MM.yyyy")
        treffer = [t for t in self._termine if t.get("datum") == datum_str]

        _WOCHENTAGE = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"]
        wd = _WOCHENTAGE[datum.dayOfWeek() - 1]
        datum_de = f"{wd}, {datum.day():02d}.{datum.month():02d}.{datum.year()}"

        try:
            from functions.notizen_db import lade_fuer_datum
            notizen_tag = lade_fuer_datum(datum_de_str)
        except Exception:
            notizen_tag = []

        # ── Dialog aufbauen ───────────────────────────────────────────
        dlg = QDialog(self)
        dlg.setWindowTitle(f"📅  {datum_de}")
        dlg.setMinimumWidth(440)
        dlg.setStyleSheet("QDialog { background: #f4f6f8; }")

        main_v = QVBoxLayout(dlg)
        main_v.setContentsMargins(16, 14, 16, 14)
        main_v.setSpacing(12)

        def _sep_lbl(text, bg, fg="#fff"):
            lbl = QLabel(text)
            lbl.setStyleSheet(
                f"background: {bg}; color: {fg}; font-size: 11px; font-weight: bold; "
                "padding: 3px 8px; border-radius: 3px;"
            )
            return lbl

        def _karte(text, bg, border):
            f = QFrame()
            f.setStyleSheet(
                f"QFrame {{ background: {bg}; border-left: 3px solid {border}; "
                f"border-radius: 4px; }}"
            )
            lbl = QLabel(text)
            lbl.setWordWrap(True)
            lbl.setStyleSheet("border: none; font-size: 12px; padding: 2px 0;")
            fl = QVBoxLayout(f)
            fl.setContentsMargins(8, 6, 8, 6)
            fl.addWidget(lbl)
            return f

        # ── Fahrzeug-Termine ──────────────────────────────────────
        main_v.addWidget(_sep_lbl("🚗  Fahrzeug-Termine", "#1565a8"))
        if treffer:
            for t in treffer:
                kz    = t.get("kennzeichen", "?")
                ftyp  = t.get("fzg_typ", "") or ""
                titel = t.get("titel", "") or t.get("typ", "")
                uhr   = t.get("uhrzeit", "") or ""
                beschr = t.get("beschreibung", "") or ""
                zeile = f"<b>[{kz}]</b>"
                if ftyp:  zeile += f"  {ftyp}"
                if uhr:   zeile += f"  –  {uhr} Uhr"
                zeile += f"<br><span style='color:#555;'>{titel}</span>"
                if beschr: zeile += f"<br><span style='color:#888; font-size:11px;'>{beschr}</span>"
                main_v.addWidget(_karte(zeile, "#e8f0fb", "#1565a8"))
        else:
            leer = QLabel("✔  Keine Fahrzeug-Termine")
            leer.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic; padding: 2px 4px;")
            main_v.addWidget(leer)

        # ── Eigene Notizen ───────────────────────────────────────
        main_v.addWidget(_sep_lbl("📝  Eigene Notizen", "#2e7d32"))
        if notizen_tag:
            _STATUS_ICON = {"offen": "🟢", "gelesen": "🔵", "erledigt": "✅"}
            for n in notizen_tag:
                si = _STATUS_ICON.get(n["status"], "📝")
                txt = f"{si}  <b>{n['titel']}</b>"
                if n["text"]:
                    txt += f"<br><span style='color:#555; font-size:11px;'>{n['text']}</span>"
                txt += f"<br><span style='color:#999; font-size:10px;'>{n['status'].capitalize()}</span>"
                main_v.addWidget(_karte(txt, "#e8f5e9", "#2e7d32"))
        else:
            leer2 = QLabel("✔  Keine Notizen für diesen Tag")
            leer2.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic; padding: 2px 4px;")
            main_v.addWidget(leer2)

        # ── Buttons ───────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_neu = _QB("➕  Neue Notiz")
        btn_neu.setStyleSheet(
            "QPushButton { background: #2e7d32; color: white; border-radius: 4px; "
            "padding: 6px 16px; font-weight: bold; } "
            "QPushButton:hover { background: #1b5e20; }"
        )
        btn_close = _QB("Schließen")
        btn_close.setStyleSheet(
            "QPushButton { background: #e0e0e0; color: #333; border-radius: 4px; padding: 6px 14px; } "
            "QPushButton:hover { background: #bdbdbd; }"
        )

        def _neue_notiz_und_refresh():
            dlg.accept()
            self._neue_notiz_dialog(datum_de_str)

        btn_neu.clicked.connect(_neue_notiz_und_refresh)
        btn_close.clicked.connect(dlg.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        btn_row.addWidget(btn_neu)
        main_v.addLayout(btn_row)

        dlg.exec()

    # ── Kalender-Markierungen ─────────────────────────────────────────────

    def _markiere_termine(self):
        # Alle alten Markierungen zurücksetzen (letzter Monat ± 6 Monate)
        today = QDate.currentDate()
        leer = QTextCharFormat()
        for offset in range(-6 * 30, 6 * 30):
            d = today.addDays(offset)
            self._kalender.setDateTextFormat(d, leer)

        # ── Formate ───────────────────────────────────────────────────────
        # Heute: kräftiges DRK-Rot, weiße Schrift, unterstrichen
        heute_fmt = QTextCharFormat()
        heute_fmt.setBackground(QColor("#C8102E"))
        heute_fmt.setForeground(QColor("#ffffff"))
        heute_fmt.setFontWeight(800)
        heute_fmt.setFontUnderline(True)

        # Morgen: warmes Orange, dunkle Schrift
        morgen_fmt = QTextCharFormat()
        morgen_fmt.setBackground(QColor("#e65100"))
        morgen_fmt.setForeground(QColor("#ffffff"))
        morgen_fmt.setFontWeight(700)
        morgen_fmt.setFontItalic(True)

        # Diese Woche (2–6 Tage): helles Gelb-Orange
        soon_fmt = QTextCharFormat()
        soon_fmt.setBackground(QColor("#fff3e0"))
        soon_fmt.setForeground(QColor("#bf360c"))
        soon_fmt.setFontWeight(700)

        # Weiter in der Zukunft: kräftiges Grün
        termin_fmt = QTextCharFormat()
        termin_fmt.setBackground(QColor("#e8f5e9"))
        termin_fmt.setForeground(QColor("#1b5e20"))
        termin_fmt.setFontWeight(700)

        morgen = today.addDays(1)
        in6    = today.addDays(6)

        # Termine nach Datum gruppieren für Tooltip
        by_date: dict[str, list] = {}
        for t in self._termine:
            ds = t.get("datum", "")
            if ds:
                by_date.setdefault(ds, []).append(t)

        self._kalender.set_termin_dates(set(by_date.keys()))

        # Notiz-Dots im Kalender (grün) — aus aktiven Notizen der letzten 5 Tage
        try:
            from functions.notizen_db import lade_alle as _lade_alle_notizen
            from datetime import datetime as _dt2
            _notiz_iso: set[str] = set()
            for _n in _lade_alle_notizen():
                try:
                    _nd = _dt2.strptime(_n["datum"], "%d.%m.%Y")
                    _notiz_iso.add(_nd.strftime("%Y-%m-%d"))
                except ValueError:
                    pass
            self._kalender.set_notiz_dates(_notiz_iso)
        except Exception:
            pass

        for datum_str, tage_termine in by_date.items():
            parts = datum_str.split("-")
            if len(parts) != 3:
                continue
            d = QDate(int(parts[0]), int(parts[1]), int(parts[2]))

            # Tooltip: alle Termine des Tages in Kurzform
            tooltip_zeilen = []
            for t in tage_termine:
                kz    = t.get("kennzeichen", "?")
                titel = t.get("titel", "") or t.get("typ", "")
                uhr   = t.get("uhrzeit", "") or ""
                uhr_txt = f" {uhr}" if uhr else ""
                tooltip_zeilen.append(f"• [{kz}]{uhr_txt}  {titel}" if titel else f"• [{kz}]{uhr_txt}")
            tooltip = "\n".join(tooltip_zeilen)

            if d == today:
                fmt = heute_fmt
            elif d == morgen:
                fmt = morgen_fmt
            elif today < d <= in6:
                fmt = soon_fmt
            else:
                fmt = termin_fmt

            fmt2 = QTextCharFormat(fmt)
            fmt2.setToolTip(tooltip)
            self._kalender.setDateTextFormat(d, fmt2)

    # ── Termin-Liste aktualisieren ────────────────────────────────────────

    def _zeige_termine_liste(self):
        # Alte Einträge entfernen
        while self._termin_layout.count():
            item = self._termin_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        today = QDate.currentDate()
        morgen = today.addDays(1)

        if not self._termine:
            leer = QLabel("✅  Keine bevorstehenden Fahrzeug-Termine")
            leer.setStyleSheet("color: #888; font-size: 11px; padding: 4px 0;")
            self._termin_layout.addWidget(leer)
            return

        for t in self._termine[:10]:
            datum_str = t.get("datum", "")
            if datum_str:
                parts = datum_str.split("-")
                if len(parts) == 3:
                    d = QDate(int(parts[0]), int(parts[1]), int(parts[2]))
                    datum_de = f"{d.day():02d}.{d.month():02d}.{d.year()}"
                else:
                    datum_de = datum_str
            else:
                datum_de = ""

            kz = t.get("kennzeichen", "?")
            titel = t.get("titel", "") or t.get("typ", "")
            uhrzeit = t.get("uhrzeit", "") or ""

            if datum_str:
                parts = datum_str.split("-")
                if len(parts) == 3:
                    d_check = QDate(int(parts[0]), int(parts[1]), int(parts[2]))
                    if d_check == today:
                        farbe = "#C8102E"
                        badge = " 🔴 HEUTE"
                    elif d_check == morgen:
                        farbe = "#e53935"
                        badge = " 🟠 Morgen"
                    else:
                        farbe = "#1565a8"
                        badge = ""
                else:
                    farbe = "#555"
                    badge = ""
            else:
                farbe = "#555"
                badge = ""

            uhr_txt = f"  {uhrzeit}" if uhrzeit else ""
            text = f"{datum_de}{uhr_txt}  [{kz}]  {titel}{badge}"

            lbl = QLabel(text)
            lbl.setWordWrap(True)
            lbl.setStyleSheet(
                f"background: white; color: {farbe}; border-left: 3px solid {farbe};"
                "border-radius: 4px; padding: 5px 8px; font-size: 12px;"
            )
            self._termin_layout.addWidget(lbl)

        if len(self._termine) > 10:
            mehr = QLabel(f"… und {len(self._termine) - 10} weitere Termine")
            mehr.setStyleSheet("color: #888; font-size: 11px; padding: 2px 4px;")
            self._termin_layout.addWidget(mehr)

    # ── Vorfeld-Kachel ───────────────────────────────────────────────────

    def _lade_vorfeld_kachel(self):
        """Lädt den aktuellsten Sonderaufgaben-Snapshot und zeigt die Vorfeldmitarbeiter."""
        # Container leeren
        while self._vorfeld_kachel_layout.count():
            item = self._vorfeld_kachel_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        try:
            from database import sonderaufgaben_db
            snapshots = sonderaufgaben_db.get_snapshots(limit=1)
            if not snapshots:
                leer = QLabel("Keine Einträge vorhanden")
                leer.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic; padding: 4px 0;")
                self._vorfeld_kachel_layout.addWidget(leer)
                self._vorfeld_kachel_datum_lbl.setText("")
                return
            snap_meta = snapshots[0]
            snap = sonderaufgaben_db.get_snapshot_by_id(snap_meta["id"])
            if not snap:
                return
        except Exception:
            leer = QLabel("Fehler beim Laden")
            leer.setStyleSheet("color: #bb0000; font-size: 11px; padding: 4px 0;")
            self._vorfeld_kachel_layout.addWidget(leer)
            self._vorfeld_kachel_datum_lbl.setText("")
            return

        datum = snap.get("datum") or snap_meta.get("datum") or ""
        gespeichert = snap.get("gespeichert_am", "")
        self._vorfeld_kachel_datum_lbl.setText(
            f"vom {datum}  (gespeichert: {gespeichert[:16]})" if gespeichert else f"vom {datum}"
        )

        vorfeld: dict = snap.get("vorfeld") or {}

        _GRUPPEN = [
            ("gr1", "Gr. 1  09–14 Uhr", "#1B5E20"),
            ("gr2", "Gr. 2  14–19 Uhr", "#E65100"),
            ("gr3", "Gr. 3  19–00 Uhr", "#4527A0"),
        ]

        hat_eintraege = False
        for group_key, group_label, group_color in _GRUPPEN:
            namen = [
                vorfeld.get(f"vorfeld_{group_key}_{slot}", "").strip()
                for slot in range(1, 4)
            ]
            namen = [n for n in namen if n]
            if not namen:
                continue
            hat_eintraege = True
            # Kompakte Einzeilen-Darstellung: Farbiger Badge + Namen kommasepariert
            namen_str = ",  ".join(namen)
            grp_lbl = QLabel(
                f"<span style='font-weight:bold; color:{group_color};'>{group_label}:</span>"
                f"&nbsp;&nbsp;<span style='color:#333;'>{namen_str}</span>"
            )
            grp_lbl.setWordWrap(True)
            grp_lbl.setStyleSheet(
                f"background: white; border-left: 3px solid {group_color};"
                "border-radius: 3px; padding: 3px 8px; font-size: 12px;"
            )
            self._vorfeld_kachel_layout.addWidget(grp_lbl)

        if not hat_eintraege:
            leer = QLabel("Keine Vorfeldmitarbeiter eingetragen")
            leer.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic; padding: 4px 0;")
            self._vorfeld_kachel_layout.addWidget(leer)

    # ── Heutiger Dienstplan ─────────────────────────────────────────────────

    _MONATSORDNER = {
        1: "01_Januar",  2: "02_Februar",  3: "03_März",
        4: "04_April",   5: "05_Mai",      6: "06_Juni",
        7: "07_July",    8: "08_August",   9: "09_September",
        10: "10_Oktober", 11: "11_November", 12: "12_Dezember",
    }

    def _aktualisiere_dp_panels(self):
        """Lädt den heutigen Plan (immer) und den gestrigen Plan (00:01–07:00 Uhr)."""
        from datetime import date as _date, timedelta
        jetzt = QTime.currentTime()
        ms = jetzt.msecsSinceStartOfDay()
        nachts = QTime(0, 1).msecsSinceStartOfDay() <= ms <= QTime(7, 0).msecsSinceStartOfDay()
        self._dp_panel_gestern.setVisible(nachts)
        heute = _date.today()
        self._lade_dienstplan_fuer_datum(heute, self._dp_panel_heute)
        if nachts:
            self._lade_dienstplan_fuer_datum(heute - timedelta(days=1), self._dp_panel_gestern)

    def _lade_dienstplan_fuer_datum(self, datum, panel):
        from PySide6.QtWidgets import QTableWidgetItem
        import os

        _TAG_DIENSTE   = frozenset({'T', 'T10', 'T8', 'DT', 'DT3'})
        _NACHT_DIENSTE = frozenset({'N', 'N10', 'NF', 'DN', 'DN3'})
        STATIONSLEITUNG = {'lars peters'}

        tbl        = panel._table
        status_lbl = panel._status_lbl
        count_lbl  = panel._count_lbl

        tbl.clearSpans()
        tbl.setRowCount(0)
        try:
            from functions.settings_functions import get_setting
            basis = get_setting("dienstplan_ordner")
        except Exception:
            basis = ""

        if not basis or not os.path.isdir(basis):
            status_lbl.setText("⚠️  Dienstplan-Ordner nicht gefunden.")
            count_lbl.setText("")
            return

        monatsordner = self._MONATSORDNER.get(datum.month, "")
        dateiname = datum.strftime("%d.%m.%Y") + ".xlsx"
        pfad = os.path.join(basis, monatsordner, dateiname)

        if not os.path.isfile(pfad):
            status_lbl.setText(f"📂  Keine Datei für {dateiname} vorhanden.")
            count_lbl.setText("")
            return

        try:
            from functions.dienstplan_parser import DienstplanParser
            data = DienstplanParser(pfad, alle_anzeigen=True).parse()
        except Exception as exc:
            status_lbl.setText(f"⚠️  Fehler beim Lesen: {exc}")
            count_lbl.setText("")
            return

        if not data.get("success"):
            status_lbl.setText("⚠️  Dienstplan konnte nicht verarbeitet werden.")
            count_lbl.setText("")
            return

        datum_str = data.get("datum") or datum.strftime("%d.%m.%Y")
        status_lbl.setText(f"Dienstplan vom {datum_str}")

        # ── Personen nach Schichttyp aufteilen ──────────────────────────────
        tag_personen   = []
        nacht_personen = []
        sonst_personen = []

        for kat, liste in (('Dispo', data.get('dispo', [])),
                           ('Betreuer', data.get('betreuer', []))):
            for p in liste:
                name_lower = p.get('vollname', p.get('anzeigename', '')).strip().lower()
                effekt_kat = 'Stationsleitung' if name_lower in STATIONSLEITUNG else kat
                dk = (p.get('dienst_kategorie') or '').upper()
                if dk in _TAG_DIENSTE:
                    tag_personen.append((effekt_kat, p))
                elif dk in _NACHT_DIENSTE:
                    nacht_personen.append((effekt_kat, p))
                else:
                    sonst_personen.append((effekt_kat, p))

        # ── Kranke nach Schichttyp + Dispo/Betreuer ──────────────────────────
        krank_tag_dispo   = []
        krank_tag_betr    = []
        krank_nacht_dispo = []
        krank_nacht_betr  = []
        krank_sonder      = []

        for p in data.get('kranke', []):
            stype = p.get('krank_schicht_typ') or 'sonderdienst'
            is_d  = p.get('krank_ist_dispo', False)
            if stype == 'tagdienst':
                (krank_tag_dispo if is_d else krank_tag_betr).append(p)
            elif stype == 'nachtdienst':
                (krank_nacht_dispo if is_d else krank_nacht_betr).append(p)
            else:
                krank_sonder.append(p)

        krank_tag_personen    = ([('KrankDispo', p) for p in krank_tag_dispo] +
                                 [('Krank',      p) for p in krank_tag_betr])
        krank_nacht_personen  = ([('KrankDispo', p) for p in krank_nacht_dispo] +
                                 [('Krank',      p) for p in krank_nacht_betr])
        krank_sonder_personen = [('Krank', p) for p in krank_sonder]

        abschnitte = []
        if tag_personen:
            abschnitte.append(('☀️ Tagdienst',           '#1565a8', tag_personen))
        if nacht_personen:
            abschnitte.append(('🌙 Nachtdienst',          '#0d2b4a', nacht_personen))
        if sonst_personen:
            abschnitte.append(('📋 Sonstige',             '#555555', sonst_personen))
        if krank_tag_personen:
            abschnitte.append(('🤒 Krank – Tagdienst',   '#8b0000', krank_tag_personen))
        if krank_nacht_personen:
            abschnitte.append(('🤒 Krank – Nachtdienst', '#5a0000', krank_nacht_personen))
        if krank_sonder_personen:
            abschnitte.append(('🤒 Krank – Sonder',      '#6b3300', krank_sonder_personen))

        total_rows = sum(1 + len(personen) for _, _, personen in abschnitte)
        tbl.clearSpans()
        tbl.setRowCount(total_rows)

        farben = {
            'Dispo':           QColor('#dce8f5'),
            'Betreuer':        QColor('#ffffff'),
            'Stationsleitung': QColor('#fff8e1'),
            'Krank':           QColor('#fce8e8'),
            'KrankDispo':      QColor('#f0d0d0'),
        }
        text_farben = {
            'Dispo':           QColor('#0a5ba4'),
            'Betreuer':        QColor('#1a1a1a'),
            'Stationsleitung': QColor('#7a5000'),
            'Krank':           QColor('#bb0000'),
            'KrankDispo':      QColor('#7a0000'),
        }

        row = 0
        sep_font = QFont('Arial', 10, QFont.Weight.Bold)
        for label_text, hdr_bg, personen in abschnitte:
            tbl.setSpan(row, 0, 1, 5)
            sep_item = QTableWidgetItem(label_text)
            sep_item.setBackground(QColor(hdr_bg))
            sep_item.setForeground(QColor('#ffffff'))
            sep_item.setFont(sep_font)
            sep_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            tbl.setItem(row, 0, sep_item)
            tbl.verticalHeader().resizeSection(row, 24)
            row += 1

            for kategorie, p in personen:
                bg = farben.get(kategorie, QColor('#ffffff'))
                fg = text_farben.get(kategorie, QColor('#1a1a1a'))
                kat_anzeige = ('Dispo' if kategorie == 'KrankDispo'
                               else ('Betreuer' if kategorie == 'Krank' else kategorie))
                dienst_anzeige = (
                    p.get('krank_abgeleiteter_dienst') or p.get('dienst_kategorie') or ''
                    if p.get('ist_krank') else p.get('dienst_kategorie') or ''
                )
                vals = [
                    kat_anzeige,
                    p.get('anzeigename', ''),
                    dienst_anzeige,
                    p.get('start_zeit', '') or '',
                    p.get('end_zeit',   '') or '',
                ]
                for col, val in enumerate(vals):
                    item = QTableWidgetItem(str(val))
                    item.setBackground(bg)
                    item.setForeground(fg)
                    tbl.setItem(row, col, item)
                if p.get('ist_bulmorfahrer'):
                    for col in range(5):
                        it = tbl.item(row, col)
                        if it:
                            it.setBackground(QColor('#fff3b0'))
                row += 1

        # ── Statuszeile ──────────────────────────────────────────────────────
        tag_n   = len(tag_personen)
        nacht_n = len(nacht_personen)
        sonst_n = len(sonst_personen)
        tag_dispo_n   = sum(1 for k, _ in tag_personen   if k == 'Dispo')
        nacht_dispo_n = sum(1 for k, _ in nacht_personen if k == 'Dispo')
        tag_betr_n    = tag_n - tag_dispo_n
        nacht_betr_n  = nacht_n - nacht_dispo_n
        n_kd_tag   = len(krank_tag_dispo)
        n_kb_tag   = len(krank_tag_betr)
        n_kd_nacht = len(krank_nacht_dispo)
        n_kb_nacht = len(krank_nacht_betr)
        n_k_sonder = len(krank_sonder)
        krank_gesamt = n_kd_tag + n_kb_tag + n_kd_nacht + n_kb_nacht + n_k_sonder

        teile = []
        if tag_n:
            tag_sub = []
            if tag_betr_n:  tag_sub.append(f'Betreuer {tag_betr_n}')
            if tag_dispo_n: tag_sub.append(f'Dispo {tag_dispo_n}')
            teile.append(
                f'{tag_n} Tagdienst ({", ".join(tag_sub)})' if tag_sub else f'{tag_n} Tagdienst'
            )
        if nacht_n:
            nacht_sub = []
            if nacht_betr_n:  nacht_sub.append(f'Betreuer {nacht_betr_n}')
            if nacht_dispo_n: nacht_sub.append(f'Dispo {nacht_dispo_n}')
            teile.append(
                f'{nacht_n} Nachtdienst ({", ".join(nacht_sub)})' if nacht_sub
                else f'{nacht_n} Nachtdienst'
            )
        if sonst_n:
            teile.append(f'{sonst_n} Sonstige')
        if krank_gesamt:
            betr_teile = []
            if n_kb_tag:   betr_teile.append(f'{n_kb_tag} Tag')
            if n_kb_nacht: betr_teile.append(f'{n_kb_nacht} Nacht')
            if n_k_sonder: betr_teile.append(f'{n_k_sonder} Sonder')
            betr_gesamt = n_kb_tag + n_kb_nacht + n_k_sonder
            dispo_teile = []
            if n_kd_tag:   dispo_teile.append(f'{n_kd_tag} Tag')
            if n_kd_nacht: dispo_teile.append(f'{n_kd_nacht} Nacht')
            dispo_gesamt = n_kd_tag + n_kd_nacht
            krank_blocks = []
            if betr_gesamt:
                krank_blocks.append(f'Betreuer {betr_gesamt} ({" / ".join(betr_teile)})')
            if dispo_gesamt:
                krank_blocks.append(f'Dispo {dispo_gesamt} ({" / ".join(dispo_teile)})')
            teile.append(f'{krank_gesamt} Krank  –  {" | ".join(krank_blocks)}')

        count_lbl.setText('  |  '.join(teile) if teile else '0 Einträge')

    # ── Refresh ───────────────────────────────────────────────────────────

    def refresh(self):
        """Aktualisiert alle Dashboard-Daten."""
        # DB-Verbindung testen
        try:
            from database.connection import test_connection
            ok, info = test_connection()
            if ok:
                self._db_status_lbl.setText(f"✅ Datenbank verbunden  |  {info[:60]}")
                self._db_status_lbl.setStyleSheet(
                    "background-color: #e8f5e8; border-radius: 6px; "
                    "border-left: 4px solid #107e3e; padding: 8px 12px; color: #107e3e;"
                )
            else:
                self._db_status_lbl.setText(f"❌ Keine Datenbankverbindung: {info[:80]}")
                self._db_status_lbl.setStyleSheet(
                    "background-color: #fce8e8; border-radius: 6px; "
                    "border-left: 4px solid #bb0000; padding: 8px 12px; color: #bb0000;"
                )
        except Exception as e:
            self._db_status_lbl.setText(f"❌ Fehler: {e}")

        # Fahrzeug-Termine laden
        self._termine = self._lade_fahrzeug_termine()
        self._markiere_termine()
        self._zeige_termine_liste()

        # Vorfeld-Kachel (letzter Sonderaufgaben-Snapshot)
        self._lade_vorfeld_kachel()

        # Krankmeldungen aktueller Monat
        self._lade_krankmeldungen()

        # Dienstplan
        self._aktualisiere_dp_panels()

        # PAX-Statistiken
        self._lade_pax_stats()

        # Notizen laden + anzeigen
        self._zeige_notizen()

    # ── PAX-Statistiken ───────────────────────────────────────────────────

    def _lade_pax_stats(self):
        """Lädt PAX-Jahressumme und Vortag-PAX aus der DB und aktualisiert die StatCards."""
        try:
            from database.pax_db import lade_jahres_pax, lade_tages_pax
            from datetime import date as _date, timedelta
            heute = _date.today()
            gestern = heute - timedelta(days=1)
            jahr = heute.year
            pax_jahr = lade_jahres_pax(jahr)
            pax_gestern = lade_tages_pax(gestern.isoformat())
            self._card_pax_jahr.set_value(f"{pax_jahr:,}".replace(",", "."))
            self._card_pax_gestern.set_value(f"{pax_gestern:,}".replace(",", ".") if pax_gestern else "–")
        except Exception:
            pass  # Tabelle existiert noch nicht → einfach ignorieren

    # ── Krankmeldungen ────────────────────────────────────────────────────

    _KM_BASIS = (
        r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V"
        r"\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26"
        r"\03_Krankmeldungen"
    )

    def _lade_krankmeldungen(self):
        from datetime import date as _date
        from PySide6.QtWidgets import QTableWidgetItem
        import glob, os

        heute = _date.today()
        jahr_ordner = os.path.join(self._KM_BASIS, str(heute.year))
        if not os.path.isdir(jahr_ordner):
            self._km_status_lbl.setText(f"⚠️  Ordner {heute.year} nicht gefunden.")
            self._km_table.setRowCount(0)
            return

        muster = os.path.join(jahr_ordner, f"{heute.month:02d}_*.xlsm")
        treffer = glob.glob(muster)
        if not treffer:
            self._km_status_lbl.setText(
                f"📂  Keine Datei für {heute.month:02d}/{heute.year} gefunden."
            )
            self._km_table.setRowCount(0)
            return

        pfad = treffer[0]
        try:
            import openpyxl
            wb = openpyxl.load_workbook(pfad, read_only=True, data_only=True, keep_vba=False)
            ws = wb.active
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            wb.close()
        except Exception as exc:
            self._km_status_lbl.setText(f"⚠️  Fehler beim Lesen: {exc}")
            self._km_table.setRowCount(0)
            return

        # Nur Zeilen mit mindestens einem Namen
        eintraege = [r for r in rows if r and r[1]]

        if not eintraege:
            self._km_status_lbl.setText("Keine Einträge vorhanden.")
            self._km_table.setRowCount(0)
            return

        dateiname = os.path.basename(pfad)
        self._km_status_lbl.setText(
            f"Aus: {dateiname}  –  {len(eintraege)} Eintrag/Einträge"
        )

        self._km_table.setRowCount(len(eintraege))
        for row_i, r in enumerate(eintraege):
            # r: (Datum, Meldender, Krank von, Krank bis, Anruf um, Angenommen von, Bemerkungen, ...)
            name    = str(r[1]) if r[1] is not None else ""
            von_raw = r[2]
            bis_raw = r[3]
            bem     = str(r[6]) if len(r) > 6 and r[6] is not None else ""

            def _fmt_date(v):
                if v is None:
                    return ""
                if hasattr(v, "strftime"):
                    return v.strftime("%d.%m.")
                return str(v)

            von_str = _fmt_date(von_raw)
            bis_str = _fmt_date(bis_raw)

            # Farbe: Krankmeldung die noch läuft (bis >= heute) → leicht rot
            laeuft_noch = False
            if hasattr(bis_raw, "date"):
                laeuft_noch = bis_raw.date() >= heute
            elif isinstance(bis_raw, _date):
                laeuft_noch = bis_raw >= heute

            bg = QColor("#fce8e8") if laeuft_noch else QColor("#ffffff")
            fg = QColor("#bb0000") if laeuft_noch else QColor("#1a1a1a")

            self._km_table.setRowHeight(row_i, 18)
            for col_i, text in enumerate([name, von_str, bis_str, bem]):
                item = QTableWidgetItem(text)
                item.setBackground(bg)
                item.setForeground(fg)
                self._km_table.setItem(row_i, col_i, item)

    def get_termine(self) -> list[dict]:
        """Gibt die zuletzt geladenen Fahrzeug-Termine zurück (für badge/popup)."""
        return self._termine

    # ── Notizen ───────────────────────────────────────────────────────────────

    def _zeige_notizen(self):
        """Notizen im Fenster -5 bis +10 Tage (nach Datum-Feld), nach Datum sortiert."""
        while self._notiz_vlayout.count():
            item = self._notiz_vlayout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        try:
            from functions.notizen_db import lade_fenster
            notizen = lade_fenster()
        except Exception:
            notizen = []

        if not notizen:
            leer = QLabel("Keine Notizen (letzte 5 / nächste 10 Tage)")
            leer.setStyleSheet("color: #aaa; font-size: 11px; font-style: italic; padding: 4px 2px;")
            self._notiz_vlayout.addWidget(leer)
            return

        from PySide6.QtWidgets import QPushButton as _QPB4
        from datetime import datetime as _dtt, date as _ddate

        _WOCHENTAGE = ["Mo","Di","Mi","Do","Fr","Sa","So"]
        _MONATE_K   = ["","Jan","Feb","Mrz","Apr","Mai","Jun",
                       "Jul","Aug","Sep","Okt","Nov","Dez"]
        _BTN_MINI = (
            "QPushButton {{ font-size: 10px; padding: 1px 8px; border-radius: 3px; "
            "background: {bg}; color: {fg}; border: 1px solid {bd}; }} "
            "QPushButton:hover {{ background: {hv}; }}"
        )
        heute = _ddate.today()

        # Nach Datum gruppieren (aufsteigend – schon sortiert aus DB)
        gruppen: dict[str, list] = {}
        for n in notizen:
            gruppen.setdefault(n["datum"], []).append(n)

        for datum_str, gruppe in gruppen.items():
            try:
                d = _dtt.strptime(datum_str, "%d.%m.%Y").date()
                wd = _WOCHENTAGE[d.weekday()]
                mo = _MONATE_K[d.month]
                delta = (d - heute).days
                if d == heute:
                    tag_label = f"🟢  Heute – {wd}, {d.day}. {mo} {d.year}"
                    hdr_bg    = "#2e7d32"
                elif delta == -1:
                    tag_label = f"🔵  Gestern – {wd}, {d.day}. {mo} {d.year}"
                    hdr_bg    = "#1565a8"
                elif delta == 1:
                    tag_label = f"📅  Morgen – {wd}, {d.day}. {mo} {d.year}"
                    hdr_bg    = "#4a6fa5"
                elif delta > 1:
                    tag_label = f"📅  In {delta} Tagen – {wd}, {d.day}. {mo} {d.year}"
                    hdr_bg    = "#4a6fa5"
                else:
                    tag_label = f"📍  {wd}, {d.day}. {mo} {d.year}"
                    hdr_bg    = "#546e7a"
            except ValueError:
                tag_label = f"📍  {datum_str}"
                hdr_bg    = "#546e7a"
                d         = None

            hdr = QLabel(tag_label)
            hdr.setStyleSheet(
                f"background: {hdr_bg}; color: white; font-size: 11px; font-weight: bold; "
                "padding: 3px 8px; border-radius: 3px; margin-top: 4px;"
            )
            self._notiz_vlayout.addWidget(hdr)

            for n in gruppe:
                nid    = n["id"]
                titel  = n["titel"]
                text   = n["text"]
                status = n["status"]

                if status == "erledigt":
                    border_color = "#9e9e9e"
                    bg_color     = "#f5f5f5"
                    txt_color    = "#9e9e9e"
                elif status == "gelesen":
                    border_color = "#1976d2"
                    bg_color     = "#e3f2fd"
                    txt_color    = "#1565c0"
                else:
                    border_color = "#2e7d32"
                    bg_color     = "#e8f5e9"
                    txt_color    = "#1b5e20"

                card = QFrame()
                card.setStyleSheet(
                    f"QFrame {{ background: {bg_color}; border-left: 3px solid {border_color}; "
                    f"border-radius: 4px; margin-left: 8px; }}"
                )
                card_v = QVBoxLayout(card)
                card_v.setContentsMargins(8, 5, 8, 5)
                card_v.setSpacing(2)

                status_icon = {"offen": "🟢", "gelesen": "🔵", "erledigt": "✅"}.get(status, "📝")
                titel_lbl = QLabel(f"{status_icon}  <b>{titel}</b>")
                titel_lbl.setStyleSheet(f"color: {txt_color}; font-size: 12px; border: none;")
                card_v.addWidget(titel_lbl)

                if text:
                    txt_lbl = QLabel(text)
                    txt_lbl.setWordWrap(True)
                    txt_lbl.setStyleSheet(f"color: {txt_color}; font-size: 11px; border: none;")
                    card_v.addWidget(txt_lbl)

                btn_row = QHBoxLayout()
                btn_row.setContentsMargins(0, 3, 0, 0)
                btn_row.setSpacing(6)

                if status == "offen":
                    btn_gelesen = _QPB4("👁  Gelesen")
                    btn_gelesen.setFixedHeight(20)
                    btn_gelesen.setStyleSheet(_BTN_MINI.format(
                        bg="#bbdefb", fg="#0d47a1", bd="#90caf9", hv="#90caf9"
                    ))
                    btn_gelesen.clicked.connect(lambda _, _nid=nid: self._notiz_als_gelesen(_nid))
                    btn_row.addWidget(btn_gelesen)

                if status == "erledigt":
                    btn_rueck = _QPB4("↩️  Rükg.")
                    btn_rueck.setFixedHeight(20)
                    btn_rueck.setToolTip("Erledigt rükgängig machen")
                    btn_rueck.setStyleSheet(_BTN_MINI.format(
                        bg="#fff9c4", fg="#f57f17", bd="#fff176", hv="#fff176"
                    ))
                    btn_rueck.clicked.connect(lambda _, _nid=nid: self._notiz_als_offen(_nid))
                    btn_row.addWidget(btn_rueck)

                if status != "erledigt":
                    btn_erledigt = _QPB4("✅  Erledigt")
                    btn_erledigt.setFixedHeight(20)
                    btn_erledigt.setStyleSheet(_BTN_MINI.format(
                        bg="#c8e6c9", fg="#1b5e20", bd="#a5d6a7", hv="#a5d6a7"
                    ))
                    btn_erledigt.clicked.connect(lambda _, _nid=nid: self._notiz_als_erledigt(_nid))
                    btn_row.addWidget(btn_erledigt)

                btn_del = _QPB4("🗑")
                btn_del.setFixedHeight(20)
                btn_del.setFixedWidth(28)
                btn_del.setToolTip("Notiz löschen")
                btn_del.setStyleSheet(_BTN_MINI.format(
                    bg="#fce8e8", fg="#b71c1c", bd="#ef9a9a", hv="#ef9a9a"
                ))
                btn_del.clicked.connect(lambda _, _nid=nid: self._notiz_loeschen(_nid))
                btn_row.addWidget(btn_del)
                btn_row.addStretch()
                card_v.addLayout(btn_row)

                self._notiz_vlayout.addWidget(card)

    def _notiz_als_gelesen(self, nid: int):
        try:
            from functions.notizen_db import als_gelesen
            als_gelesen(nid)
        except Exception:
            pass
        self._zeige_notizen()

    def _notiz_als_offen(self, nid: int):
        try:
            from functions.notizen_db import als_offen
            als_offen(nid)
        except Exception:
            pass
        self._zeige_notizen()

    def _notiz_als_erledigt(self, nid: int):
        try:
            from functions.notizen_db import als_erledigt
            als_erledigt(nid)
        except Exception:
            pass
        self._zeige_notizen()
        self._markiere_termine()  # Kalender-Dots aktualisieren

    def _notiz_loeschen(self, nid: int):
        from PySide6.QtWidgets import QMessageBox as _QMB
        if _QMB.question(
            self, "Notiz löschen",
            "Notiz wirklich dauerhaft löschen?",
            _QMB.StandardButton.Yes | _QMB.StandardButton.No,
            _QMB.StandardButton.No,
        ) != _QMB.StandardButton.Yes:
            return
        try:
            from functions.notizen_db import loeschen
            loeschen(nid)
        except Exception:
            pass
        self._zeige_notizen()
        self._markiere_termine()

    def _notiz_archiv_dialog(self):
        """Dialog: alle Notizen aus der Datenbank anzeigen (Archiv)."""
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
            QPushButton, QScrollArea, QWidget, QFrame,
        )
        from PySide6.QtCore import Qt

        try:
            from functions.notizen_db import lade_alle, loeschen as _notiz_loeschen_db
        except Exception:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("🗄️  Notizen-Archiv")
        dlg.setMinimumSize(560, 520)
        dlg.setStyleSheet("QDialog { background: #f5f5f5; }")

        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(8)

        # Suchzeile
        such_row = QHBoxLayout()
        such_lbl = QLabel("🔍  Suche:")
        such_lbl.setFixedWidth(60)
        such_edit = QLineEdit()
        such_edit.setPlaceholderText("Titel oder Text filtern …")
        such_edit.setStyleSheet(
            "QLineEdit { border: 1px solid #ccc; border-radius: 4px; padding: 4px 8px; background: white; }"
        )
        such_row.addWidget(such_lbl)
        such_row.addWidget(such_edit)
        outer.addLayout(such_row)

        # Scroll-Bereich
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        container = QWidget()
        container.setStyleSheet("background: transparent;")
        liste_layout = QVBoxLayout(container)
        liste_layout.setSpacing(4)
        liste_layout.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(container)
        outer.addWidget(scroll)

        _WOCHENTAGE = ["Mo","Di","Mi","Do","Fr","Sa","So"]
        _MONATE_K   = ["","Jan","Feb","Mrz","Apr","Mai","Jun",
                       "Jul","Aug","Sep","Okt","Nov","Dez"]
        _STATUS_FARBEN = {
            "offen":    ("#2e7d32", "#e8f5e9", "#1b5e20"),
            "gelesen":  ("#1976d2", "#e3f2fd", "#1565c0"),
            "erledigt": ("#9e9e9e", "#f5f5f5", "#616161"),
        }

        def _render(filter_text: str = ""):
            # Widgets leeren
            while liste_layout.count():
                item = liste_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            alle = lade_alle()
            if filter_text:
                ft = filter_text.lower()
                alle = [n for n in alle if ft in n["titel"].lower() or ft in n["text"].lower()]

            if not alle:
                leer = QLabel("Keine Notizen gefunden.")
                leer.setStyleSheet("color: #aaa; font-style: italic; padding: 8px;")
                liste_layout.addWidget(leer)
                liste_layout.addStretch()
                return

            from datetime import datetime as _dtt
            for n in alle:
                nid    = n["id"]
                titel  = n["titel"]
                text   = n["text"]
                datum  = n["datum"]
                status = n["status"]

                try:
                    d  = _dtt.strptime(datum, "%d.%m.%Y").date()
                    wd = _WOCHENTAGE[d.weekday()]
                    mo = _MONATE_K[d.month]
                    datum_txt = f"{wd}, {d.day}. {mo} {d.year}"
                except ValueError:
                    datum_txt = datum

                bd, bg, fg = _STATUS_FARBEN.get(status, ("#9e9e9e", "#f5f5f5", "#616161"))
                status_icon = {"offen": "🟢", "gelesen": "🔵", "erledigt": "✅"}.get(status, "📝")

                card = QFrame()
                card.setStyleSheet(
                    f"QFrame {{ background: {bg}; border-left: 3px solid {bd}; "
                    "border-radius: 4px; margin: 2px 0; }}"
                )
                cv = QVBoxLayout(card)
                cv.setContentsMargins(10, 6, 8, 6)
                cv.setSpacing(2)

                hdr_row = QHBoxLayout()
                titel_lbl = QLabel(f"{status_icon}  <b>{titel}</b>")
                titel_lbl.setStyleSheet(f"color: {fg}; font-size: 12px; border: none;")
                datum_lbl = QLabel(datum_txt)
                datum_lbl.setStyleSheet("color: #888; font-size: 10px; border: none;")
                datum_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                hdr_row.addWidget(titel_lbl)
                hdr_row.addStretch()
                hdr_row.addWidget(datum_lbl)
                cv.addLayout(hdr_row)

                if text:
                    txt_lbl = QLabel(text)
                    txt_lbl.setWordWrap(True)
                    txt_lbl.setStyleSheet(f"color: {fg}; font-size: 11px; border: none;")
                    cv.addWidget(txt_lbl)

                btn_del = QPushButton("🗑")
                btn_del.setFixedSize(26, 20)
                btn_del.setToolTip("Notiz löschen")
                btn_del.setStyleSheet(
                    "QPushButton { font-size: 10px; background: #fce8e8; color: #b71c1c; "
                    "border: 1px solid #ef9a9a; border-radius: 3px; } "
                    "QPushButton:hover { background: #ef9a9a; }"
                )

                def _loesche(checked=False, _nid=nid):
                    from PySide6.QtWidgets import QMessageBox as _QMB
                    if _QMB.question(
                        dlg, "Löschen", "Notiz wirklich löschen?",
                        _QMB.StandardButton.Yes | _QMB.StandardButton.No,
                        _QMB.StandardButton.No,
                    ) == _QMB.StandardButton.Yes:
                        _notiz_loeschen_db(_nid)
                        _render(such_edit.text().strip())
                        self._zeige_notizen()

                btn_del.clicked.connect(_loesche)
                btn_row = QHBoxLayout()
                btn_row.addStretch()
                btn_row.addWidget(btn_del)
                cv.addLayout(btn_row)

                liste_layout.addWidget(card)

            liste_layout.addStretch()

        _render()
        such_edit.textChanged.connect(lambda t: _render(t.strip()))

        # Schließen-Button
        btn_close = QPushButton("Schließen")
        btn_close.clicked.connect(dlg.accept)
        btn_close.setStyleSheet(
            "QPushButton { padding: 5px 18px; background: #546e7a; color: white; "
            "border-radius: 4px; font-size: 12px; } "
            "QPushButton:hover { background: #37474f; }"
        )
        outer.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

        dlg.exec()

    def _neue_notiz_dialog(self, vorgabe_datum: str = ""):
        """Dialog zum Erstellen einer neuen Notiz."""
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
            QTextEdit, QPushButton, QDateEdit,
        )
        from PySide6.QtCore import QDate

        dlg = QDialog(self)
        dlg.setWindowTitle("📝  Neue Notiz")
        dlg.setMinimumWidth(400)
        dlg.setStyleSheet("QDialog { background: #f9f9f9; }")

        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)
        layout.setContentsMargins(18, 16, 18, 16)

        # Titel
        layout.addWidget(QLabel("<b>Titel:</b>"))
        titel_edit = QLineEdit()
        titel_edit.setPlaceholderText("Kurzer Betreff …")
        titel_edit.setStyleSheet(
            "QLineEdit { border: 1px solid #ccc; border-radius: 4px; padding: 4px 8px; background: white; }"
        )
        layout.addWidget(titel_edit)

        # Datum
        layout.addWidget(QLabel("<b>Datum:</b>"))
        datum_edit = QDateEdit()
        datum_edit.setCalendarPopup(True)
        datum_edit.setDisplayFormat("dd.MM.yyyy")
        datum_edit.setStyleSheet(
            "QDateEdit { border: 1px solid #ccc; border-radius: 4px; padding: 4px 8px; background: white; }"
        )
        if vorgabe_datum:
            try:
                from datetime import datetime as _dt3
                _d = _dt3.strptime(vorgabe_datum, "%d.%m.%Y")
                datum_edit.setDate(QDate(_d.year, _d.month, _d.day))
            except ValueError:
                datum_edit.setDate(QDate.currentDate())
        else:
            datum_edit.setDate(QDate.currentDate())
        layout.addWidget(datum_edit)

        # Text
        layout.addWidget(QLabel("<b>Notiz</b> (optional):"))
        text_edit = QTextEdit()
        text_edit.setPlaceholderText("Weitere Details …")
        text_edit.setFixedHeight(90)
        text_edit.setStyleSheet(
            "QTextEdit { border: 1px solid #ccc; border-radius: 4px; padding: 4px 8px; background: white; }"
        )
        layout.addWidget(text_edit)

        # Buttons
        btn_row = QHBoxLayout()
        btn_ok = QPushButton("💾  Speichern")
        btn_ok.setStyleSheet(
            "QPushButton { background: #2e7d32; color: white; border-radius: 4px; "
            "padding: 6px 18px; font-weight: bold; } "
            "QPushButton:hover { background: #1b5e20; }"
        )
        btn_ab = QPushButton("Abbrechen")
        btn_ab.setStyleSheet(
            "QPushButton { background: #e0e0e0; color: #333; border-radius: 4px; padding: 6px 14px; } "
            "QPushButton:hover { background: #bdbdbd; }"
        )
        btn_ok.clicked.connect(dlg.accept)
        btn_ab.clicked.connect(dlg.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_ab)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        titel_edit.setFocus()

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        titel = titel_edit.text().strip()
        if not titel:
            from PySide6.QtWidgets import QMessageBox as _QMB
            _QMB.warning(self, "Titel fehlt", "Bitte einen Titel eingeben.")
            return

        datum_str = datum_edit.date().toString("dd.MM.yyyy")
        text_str  = text_edit.toPlainText().strip()

        try:
            from functions.notizen_db import speichern
            speichern(titel, text_str, datum_str)
        except Exception as exc:
            from PySide6.QtWidgets import QMessageBox as _QMB
            _QMB.critical(self, "Fehler", f"Notiz konnte nicht gespeichert werden:\n{exc}")
            return

        self._zeige_notizen()
        self._markiere_termine()  # Kalender-Dot aktualisieren
