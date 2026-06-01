"""
Haupt-Fenster (MainWindow)
SAP Fiori-Design mit Sidebar-Navigation
"""
import sys
import os
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import math, time

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QSizePolicy, QMessageBox,
    QScrollArea, QGraphicsOpacityEffect,
)
from PySide6.QtCore import Qt, QSize, QTimer, QRectF, QPointF, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import (
    QFont, QColor, QPixmap, QPainter, QPen, QBrush,
    QRadialGradient, QLinearGradient, QFontMetricsF,
)

from config import (
    APP_NAME, APP_VERSION, BASE_DIR,
    FIORI_SIDEBAR_BG, FIORI_BLUE, FIORI_WHITE, FIORI_LIGHT_BLUE, FIORI_TEXT
)


# ---------------------------------------------------------------------------
# Animiertes NeSk-Logo-Widget (identisch mit SplashScreen, skaliert auf Sidebar)
# ---------------------------------------------------------------------------
class _NeskLogoWidget(QWidget):
    """Rotierender Doppelring mit NeSk-Schriftzug und Shimmer — wie der Splash Screen."""

    W, H = 200, 170

    _RING1 = QColor("#5B8AAA")   # Teal
    _RING2 = QColor("#C0944A")   # Gold
    _ACCENT = QColor("#5B8AAA")
    _GRAY   = QColor(74, 104, 128)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._t0 = time.monotonic()
        self.setFixedSize(self.W, self.H)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("Doppelklick für 🎩 Alice's Wunderrad")
        self._timer = QTimer(self)
        self._timer.timeout.connect(self.update)
        self._timer.start(30)  # ~33 FPS
        self._slot_dlg = None

    def paintEvent(self, event):
        t  = time.monotonic() - self._t0
        W, H = float(self.W), float(self.H)

        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        # Hintergrund — Sidebar-Farbe exakt
        p.fillRect(0, 0, int(W), int(H), QBrush(QColor("#354a5e")))

        # Ring-Mittelpunkt
        cx, cy = W / 2, H * 0.42

        # Winkel
        ang1 = (t * 120.0) % 360.0
        ang2 = (180.0 - t * 75.0) % 360.0

        # Pulsierender Glow
        glow_alpha = int(18 + 14 * math.sin(t * 2.2))
        glow_r = 34.0
        glow_grad = QRadialGradient(cx, cy, glow_r + 14)
        glow_grad.setColorAt(0.0, QColor(91, 138, 170, glow_alpha * 2))
        glow_grad.setColorAt(0.6, QColor(91, 138, 170, glow_alpha))
        glow_grad.setColorAt(1.0, QColor(91, 138, 170, 0))
        p.setBrush(QBrush(glow_grad))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - glow_r - 14, cy - glow_r - 14,
                             (glow_r + 14) * 2, (glow_r + 14) * 2))

        # Innerer dunkler Kreis
        inner_r = 24.0
        p.setBrush(QBrush(QColor("#2d4155")))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - inner_r, cy - inner_r,
                             inner_r * 2, inner_r * 2))

        # "N" in der Mitte
        font_n = QFont("Segoe UI", 15, QFont.Weight.Light)
        p.setFont(font_n)
        fm_n = QFontMetricsF(font_n)
        nw   = fm_n.horizontalAdvance("N")
        p.setPen(QPen(self._RING1))
        p.drawText(QPointF(cx - nw / 2, cy + fm_n.ascent() / 2 - 1), "N")

        # Ring 1 (Teal, vorwärts)
        r1_r = 31.0
        p.setPen(QPen(self._RING1, 2.2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawArc(QRectF(cx - r1_r, cy - r1_r, r1_r * 2, r1_r * 2),
                  int(-ang1 * 16), 230 * 16)

        # Ring 2 (Gold, rückwärts)
        r2_r = 37.0
        p.setPen(QPen(self._RING2, 1.4, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.drawArc(QRectF(cx - r2_r, cy - r2_r, r2_r * 2, r2_r * 2),
                  int(-ang2 * 16), 110 * 16)

        # "NeSk" Schriftzug
        ty = cy + 50.0
        font_title = QFont("Segoe UI", 20, QFont.Weight.Light)
        p.setFont(font_title)
        fm_t = QFontMetricsF(font_title)
        text = "NeSk"
        tw   = fm_t.horizontalAdvance(text)
        tx   = cx - tw / 2

        # Basis
        p.setPen(QPen(QColor(180, 205, 220, 130)))
        p.drawText(QPointF(tx, ty), text)

        # Shimmer
        sh_prog = ((t * 0.55) % 1.9) - 0.4
        sh_cx   = tx + sh_prog * (tw + 60) - 15
        sh_w    = tw * 0.30
        sh_grad = QLinearGradient(sh_cx - sh_w, ty - 28, sh_cx + sh_w, ty)
        sh_grad.setColorAt(0.0, QColor(255, 255, 255, 0))
        sh_grad.setColorAt(0.5, QColor(255, 255, 255, 210))
        sh_grad.setColorAt(1.0, QColor(255, 255, 255, 0))
        p.setPen(QPen(QBrush(sh_grad), 0))
        p.drawText(QPointF(tx, ty), text)

        # Untertitel
        font_sub = QFont("Segoe UI", 7)
        p.setFont(font_sub)
        fm_s = QFontMetricsF(font_sub)
        sub  = "DRK  ·  Flughafen Köln / Bonn"
        sw   = fm_s.horizontalAdvance(sub)
        p.setPen(QPen(self._ACCENT))
        p.drawText(QPointF(cx - sw / 2, ty + 17), sub)

        p.end()

    def mouseDoubleClickEvent(self, event):
        from gui.slot_machine import SlotMachineDialog
        if self._slot_dlg is None or not self._slot_dlg.isVisible():
            self._slot_dlg = SlotMachineDialog()
            # Fenster zentriert über dem Haupt-Fenster positionieren
            mw = self.window()
            if mw:
                center = mw.geometry().center()
                self._slot_dlg.move(
                    center.x() - self._slot_dlg.width() // 2,
                    center.y() - self._slot_dlg.height() // 2,
                )
            self._slot_dlg.show()
        else:
            self._slot_dlg.raise_()
            self._slot_dlg.activateWindow()

from gui.dashboard        import DashboardWidget
from gui.aufgaben_haupt   import AufgabenHauptWidget
from gui.dienstplan       import DienstplanWidget
from gui.uebergabe        import UebergabeWidget
from gui.fahrzeuge        import FahrzeugeWidget
from gui.einstellungen    import EinstellungenWidget
from gui.code19           import Code19Widget
from gui.bericht          import BerichtWidget
from gui.mitarbeiter            import MitarbeiterHauptWidget
from gui.hilfe_dialog           import HilfeDialog
from gui.dienstliches           import DienstlichesWidget
from gui.telefonnummern         import TelefonnummernWidget
from gui.handys_widget          import HandysWidget
from gui.backup_widget          import BackupWidget
from gui.passagiere             import PassagiereWidget
from gui.vorkommnisse           import VorkommnisseWidget
from gui.workflow               import WorkflowWidget


NAV_ITEMS = [
    ("🏠", "Dashboard",          0),
    ("👥", "Mitarbeiter",         1),
    ("☕️", "Dienstliches",       2),
    ("📝", "Aufgaben",            3),
    ("📅", "Dienstplan",          4),
    ("📋", "Übergabe",            5),
    ("🚗", "Fahrzeuge",           6),
    ("🕐", "Code 19",            7),
    ("📊", "Bericht",             8),
    ("✈️",  "Passagiere",        9),
    ("📞", "Telefonnummern",     10),
    ("📱", "Handys",             11),
    ("💾", "Backup",             12),
    ("⚙️",  "Einstellungen",    13),
    ("⚠️",  "Vorkommnisse",     14),
    ("🔄",  "Workflow",          15),
]

NAV_TOOLTIPS = [
    "Startseite – Statistiken und Übersicht",
    "Mitarbeiter-Übersicht + Dokumente + Ausdrucke + Krankmeldungen",
    "Dienstliche Protokolle: Einsätze und Berichte",
    "Aufgaben Tag (Tagesaufgaben, Checklisten) und Aufgaben Nacht (Checklisten, Sonderaufgaben, AOCC)",
    "Dienstplan laden, anzeigen und Hausverwaltung exportieren",
    "Schichtprotokoll erstellen, ausfüllen und abschließen",
    "Fahrzeugstatus, Schäden und Wartungstermine verwalten",
    "Code-19-Protokoll führen und Uhrzeigen-Animation",
    "Kombinierten Bericht aus Verspätungen, Schulungen, Einsätzen und Patienten als Excel exportieren oder per E-Mail senden",
    "Passagieranfragen verarbeiten und Beschwerden erfassen und verwalten",
    "Telefonnummern-Verzeichnis: FKB Gate-/Check-In-Nummern und DRK-Kontakte",
    "Diensthandys verwalten: Status, Zustand, Historie, Excel-Export und E-Mail",
    "Datensicherung erstellen und wiederherstellen",
    "App-Einstellungen, Pfade und E-Mobby-Fahrerliste",
    "Vorkommnisse erfassen und verwalten",
    "Stärkemeldungen mit Tagesdienstplänen abgleichen – Abweichungen und fehlende Einträge erkennen",
]


class SidebarButton(QPushButton):
    def __init__(self, icon: str, text: str, parent=None):
        super().__init__(f"  {icon}  {text}", parent)
        self.setCheckable(True)
        self.setMinimumHeight(48)
        self.setFont(QFont("Arial", 12))
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._apply_style(False)

    def _apply_style(self, active: bool):
        _tt = "QToolTip { color: #1b3a5c; background-color: #fffde7; border: 1px solid #90a4ae; padding: 4px 8px; }"
        if active:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {FIORI_BLUE};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    text-align: left;
                    font-weight: bold;
                }}
                {_tt}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: #cdd5e0;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    text-align: left;
                }}
                QPushButton:hover {{
                    background-color: rgba(255,255,255,25);
                    color: white;
                }}
                {_tt}
            """)

    def setActive(self, active: bool):
        self._apply_style(active)
        self.setChecked(active)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1280, 800)
        self.setMinimumSize(900, 600)
        self._nav_buttons: list[SidebarButton] = []
        self._build_ui()
        self._navigate(0)
        QTimer.singleShot(800, self._check_termine_startup)

    # ── UI aufbauen ────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())
        root.addWidget(self._build_content(), 1)

    def _build_sidebar(self) -> QWidget:
        # äußerer Container mit fester Breite und Sidebar-Farbe
        outer = QWidget()
        outer.setFixedWidth(220)
        outer.setStyleSheet(f"background-color: {FIORI_SIDEBAR_BG};")
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        # ScrollArea damit die Sidebar bei kleinem Fenster scrollbar ist
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical {
                background: transparent; width: 4px; margin: 0;
            }
            QScrollBar::handle:vertical {
                background: rgba(255,255,255,64); border-radius: 2px; min-height: 20px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)

        # Logo randlos direkt in outer (außerhalb der ScrollArea)
        logo_widget = _NeskLogoWidget()
        outer_layout.addWidget(logo_widget)

        # Trennlinie (ebenfalls außerhalb, direkt unter Logo)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #4a6480;")
        outer_layout.addWidget(line)

        sidebar = QWidget()
        sidebar.setStyleSheet(f"background-color: {FIORI_SIDEBAR_BG};")

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(8, 8, 8, 16)
        layout.setSpacing(4)

        # Navigations-Buttons
        for (icon, label, idx), tooltip in zip(NAV_ITEMS, NAV_TOOLTIPS):
            btn = SidebarButton(icon, label)
            btn.setToolTip(tooltip)
            btn.clicked.connect(lambda _, i=idx: self._navigate(i))
            self._nav_buttons.append(btn)
            layout.addWidget(btn)

        layout.addStretch()

        # Termin-Badge (über Hilfe-Button)  
        self._termin_badge = QLabel("🔔  Terminhinweis")
        self._termin_badge.setWordWrap(True)
        self._termin_badge.setStyleSheet("""
            QLabel {
                background-color: #e67e22;
                color: white;
                border-radius: 4px;
                padding: 5px 8px;
                font-size: 11px;
                font-weight: bold;
            }
        """)
        self._termin_badge.setVisible(False)
        layout.addWidget(self._termin_badge)

        # Hilfe-Button
        self._hilfe_btn = QPushButton("❓  Hilfe")
        self._hilfe_btn.setMinimumHeight(36)
        self._hilfe_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._hilfe_btn.setToolTip("Bedienungsanleitung und Übersicht aller Funktionen öffnen")
        self._hilfe_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255,255,255,26);
                color: #cdd5e0;
                border: 1px solid rgba(255,255,255,46);
                border-radius: 4px;
                padding: 6px 12px;
                text-align: left;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: rgba(255,255,255,51);
                color: white;
            }
            QToolTip { color: #1b3a5c; background-color: #fffde7; border: 1px solid #90a4ae; padding: 4px 8px; }
        """)
        self._hilfe_btn.clicked.connect(lambda: HilfeDialog(self).exec())
        layout.addWidget(self._hilfe_btn)

        # Version unten
        ver_lbl = QLabel(f"v{APP_VERSION}")
        ver_lbl.setStyleSheet("color: #4a6480; font-size: 10px; padding: 0 8px;")
        layout.addWidget(ver_lbl)

        # Scroll-Widget zusammenbauen und outer zurückgeben
        scroll.setWidget(sidebar)
        outer_layout.addWidget(scroll)
        return outer

    def _build_content(self) -> QWidget:
        frame = QWidget()
        frame.setStyleSheet(f"background-color: #f5f6f7;")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)

        self._stack = QStackedWidget()

        # Pages
        self._dashboard_page         = DashboardWidget()
        self._mitarbeiter_page       = MitarbeiterHauptWidget()
        self._dienstliches_page      = DienstlichesWidget()
        self._aufgaben_haupt_page    = AufgabenHauptWidget()
        self._dienstplan_page        = DienstplanWidget()
        self._uebergabe_page         = UebergabeWidget()
        self._fahrzeuge_page         = FahrzeugeWidget()
        self._code19_page            = Code19Widget()
        self._bericht_page           = BerichtWidget()
        self._telefonnummern_page    = TelefonnummernWidget()
        self._handys_page            = HandysWidget()
        self._backup_page            = BackupWidget()
        self._settings_page          = EinstellungenWidget()
        self._passagiere_page        = PassagiereWidget()
        self._vorkommnisse_page      = VorkommnisseWidget()
        self._workflow_page          = WorkflowWidget()
        for page in [self._dashboard_page, self._mitarbeiter_page,
                     self._dienstliches_page,
                     self._aufgaben_haupt_page,
                     self._dienstplan_page, self._uebergabe_page,
                     self._fahrzeuge_page, self._code19_page,
                     self._bericht_page,
                     self._passagiere_page,
                     self._telefonnummern_page,
                     self._handys_page,
                     self._backup_page, self._settings_page,
                     self._vorkommnisse_page,
                     self._workflow_page]:
            self._stack.addWidget(page)

        layout.addWidget(self._stack)
        return frame

    def _placeholder_page(self, title: str, msg: str) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_title = QLabel(title)
        lbl_title.setFont(QFont("Arial", 22, QFont.Weight.Bold))
        lbl_title.setStyleSheet(f"color: {FIORI_TEXT};")
        lbl_msg = QLabel(msg)
        lbl_msg.setFont(QFont("Arial", 13))
        lbl_msg.setStyleSheet("color: #999;")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_msg)
        return w

    # ── Fahrzeug-Termin-Check ──────────────────────────────────────────────────

    def _check_termine_startup(self):
        """
        Prüft Fahrzeug-Termine für heute und morgen.
        Zeigt Badge an und öffnet Popup wenn Termine morgen anstehen.
        """
        try:
            termine = self._dashboard_page.get_termine()
        except Exception:
            return

        from PySide6.QtCore import QDate
        today = QDate.currentDate()
        morgen = today.addDays(1)
        today_str  = today.toString("yyyy-MM-dd")
        morgen_str = morgen.toString("yyyy-MM-dd")

        heute_termine  = [t for t in termine if t.get("datum") == today_str]
        morgen_termine = [t for t in termine if t.get("datum") == morgen_str]
        badge_termine  = heute_termine + morgen_termine

        # Badge anzeigen/verstecken
        if badge_termine:
            teile = []
            if heute_termine:
                teile.append(f"{len(heute_termine)} Termin(e) heute")
            if morgen_termine:
                teile.append(f"{len(morgen_termine)} Termin(e) morgen")
            self._termin_badge.setText("🔔  " + "  |  ".join(teile))
            self._termin_badge.setVisible(True)
        else:
            self._termin_badge.setVisible(False)

        # Popup wenn Termine morgen
        if morgen_termine:
            zeilen = []
            for t in morgen_termine:
                kz = t.get("kennzeichen", "?")
                titel = t.get("titel", "") or t.get("typ", "")
                uhr = t.get("uhrzeit", "")
                uhr_txt = f"  {uhr}" if uhr else ""
                zeilen.append(f"• [{kz}]{uhr_txt}  {titel}")
            msg = (
                f"Morgen ({morgen.toString('dd.MM.yyyy')}) stehen "
                f"{len(morgen_termine)} Fahrzeug-Termin(e) an:\n\n"
                + "\n".join(zeilen)
            )
            QMessageBox.information(
                self,
                "🚗  Fahrzeug-Termine morgen",
                msg,
            )

    # ── Navigation ─────────────────────────────────────────────────────────────
    def _navigate(self, index: int):
        for i, btn in enumerate(self._nav_buttons):
            btn.setActive(i == index)
        self._stack.setCurrentIndex(index)

        # Fade-In Animation der neuen Seite
        page = self._stack.currentWidget()
        if page:
            effect = QGraphicsOpacityEffect(page)
            page.setGraphicsEffect(effect)
            self._fade_anim = QPropertyAnimation(effect, b"opacity")
            self._fade_anim.setDuration(180)
            self._fade_anim.setStartValue(0.0)
            self._fade_anim.setEndValue(1.0)
            self._fade_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            self._fade_anim.finished.connect(lambda: page.setGraphicsEffect(None))
            self._fade_anim.start()

        # Refresh nach dem Seitenumbruch aufrufen (UI reagiert sofort)
        page_map = {
            0: self._dashboard_page.refresh,
            1: self._mitarbeiter_page.refresh,
            2: self._dienstliches_page.refresh,
            3: self._aufgaben_haupt_page.refresh,
            4: self._dienstplan_page.reload_tree,
            5: self._uebergabe_page.refresh,
            6: self._fahrzeuge_page.refresh,
            7: self._code19_page.refresh,
            8: self._passagiere_page.refresh,
            9: self._telefonnummern_page.refresh,
            10: self._handys_page.refresh,
        }
        if index in page_map:
            QTimer.singleShot(0, page_map[index])

    # ── Screenshot-Erstellung ──────────────────────────────────────────────────
    def grab_all_screenshots(self, callback=None):
        """
        Erstellt PNG-Screenshots aller App-Seiten und speichert sie in
        Daten/Hilfe/screenshots/{idx:02d}.png.
        Ruft am Ende callback(list[str]) auf.
        """
        ss_dir = Path(BASE_DIR) / "Daten" / "Hilfe" / "screenshots"
        ss_dir.mkdir(parents=True, exist_ok=True)

        self._ss_paths: list[str] = []
        self._ss_idx: int = 0
        self._ss_dir = ss_dir
        self._ss_callback = callback

        def _grab_next():
            if self._ss_idx >= len(NAV_ITEMS):
                self._navigate(0)
                if self._ss_callback:
                    self._ss_callback(self._ss_paths)
                return
            _icon, _label, page_idx = NAV_ITEMS[self._ss_idx]
            self._navigate(page_idx)
            QTimer.singleShot(300, _do_grab)

        def _do_grab():
            _icon, _label, page_idx = NAV_ITEMS[self._ss_idx]
            pixmap = self._stack.grab()
            fpath = str(self._ss_dir / f"{page_idx:02d}.png")
            pixmap.save(fpath, "PNG")
            self._ss_paths.append(fpath)
            self._ss_idx += 1
            QTimer.singleShot(50, _grab_next)

        _grab_next()
