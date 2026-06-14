"""
Ladebildschirm (Splash Screen) - Moderne Variante
- rotierender Teal-Ring + gegenlaeuftiger Gold-Ring (doppelter Spinner)
- pulsierender Glow in der Mitte
- "NeSk" Schriftzug mit wanderndem Shimmer
- Animation laeuft via time.monotonic() - kein Event-Loop noetig
"""
import os, sys, math, time

from PySide6.QtWidgets import QWidget, QApplication, QDialog
from PySide6.QtCore    import Qt, QRectF, QPointF, QTimer, QThread, Signal
from PySide6.QtGui     import (
    QColor, QPainter, QPen, QPixmap, QBrush, QFont, QFontMetricsF,
    QLinearGradient, QRadialGradient,
)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# ---------------------------------------------------------------------------
# Farbpalette
# ---------------------------------------------------------------------------
_BG      = QColor("#0F1D28")
_RING1   = QColor("#5B8AAA")     # Teal  - groesserer Ring
_RING2   = QColor("#C0944A")     # Gold  - kleinerer Gegenring
_GLOW    = QColor(91, 138, 170)  # Teal fuer Glow
_WHITE   = QColor("#E8EFF5")
_ACCENT  = QColor("#5B8AAA")
_GRAY    = QColor(74, 104, 128)


class SplashScreen(QWidget):
    """
    Frameless moderner Splash Screen - alles via QPainter.
    Animation laeuft durch time.monotonic(); set_status() pumpt Events.
    """

    W, H = 480, 340

    def __init__(self, version=""):
        super().__init__()
        self._version    = version
        self._status     = "Wird gestartet \u2026"
        self._t0         = time.monotonic()

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.SplashScreen
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setFixedSize(self.W, self.H)

        # Zentrieren
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            self.move(
                geo.center().x() - self.W // 2,
                geo.center().y() - self.H // 2,
            )

    # ------------------------------------------------------------------
    def paintEvent(self, event):
        t  = time.monotonic() - self._t0    # Sekunden seit Start
        W, H = float(self.W), float(self.H)

        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        # ── Hintergrund (radialer Gradient) ───────────────────────────────
        bg = QRadialGradient(W / 2, H * 0.44, W * 0.58)
        bg.setColorAt(0.0, QColor("#182634"))
        bg.setColorAt(1.0, QColor("#0A1520"))
        p.fillRect(0, 0, int(W), int(H), QBrush(bg))

        # Goldstreifen oben
        p.fillRect(0, 0, int(W), 3, QBrush(QColor("#C0944A")))

        # ── Spinner / Ring-Zone ───────────────────────────────────────────
        cx, cy = W / 2, H * 0.40

        # Winkel aus Echtzeit (unabhaengig vom Event-Loop)
        ang1 = (t * 120.0) % 360.0           # 120 Grad/s vorwaerts
        ang2 = (180.0 - t * 75.0) % 360.0    # 75 Grad/s rueckwaerts

        # Pulsierender Glow
        glow_alpha = int(18 + 14 * math.sin(t * 2.2))
        glow_r     = 48.0
        glow_grad  = QRadialGradient(cx, cy, glow_r + 20)
        glow_grad.setColorAt(0.0, QColor(91, 138, 170, glow_alpha * 2))
        glow_grad.setColorAt(0.6, QColor(91, 138, 170, glow_alpha))
        glow_grad.setColorAt(1.0, QColor(91, 138, 170, 0))
        p.setBrush(QBrush(glow_grad))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - glow_r - 20, cy - glow_r - 20,
                             (glow_r + 20) * 2, (glow_r + 20) * 2))

        # Innerer Kreis (dunkel)
        inner_r = 34.0
        p.setBrush(QBrush(QColor("#111E2A")))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - inner_r, cy - inner_r,
                             inner_r * 2, inner_r * 2))

        # "N" in der Mitte des Kreises (statt Logo)
        font_n = QFont("Segoe UI", 22, QFont.Weight.Light)
        p.setFont(font_n)
        fm_n = QFontMetricsF(font_n)
        nw   = fm_n.horizontalAdvance("N")
        p.setPen(QPen(_RING1))
        p.drawText(QPointF(cx - nw / 2, cy + fm_n.ascent() / 2 - 2), "N")

        # Ring 1 (Teal, gross, vorwaerts)
        r1_r   = 44.0
        r1_pen = QPen(_RING1, 2.8, Qt.PenStyle.SolidLine,
                      Qt.PenCapStyle.RoundCap)
        p.setPen(r1_pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        r1_rect = QRectF(cx - r1_r, cy - r1_r, r1_r * 2, r1_r * 2)
        p.drawArc(r1_rect, int(-ang1 * 16), 230 * 16)

        # Ring 2 (Gold, kleiner, rueckwaerts)
        r2_r   = 52.0
        r2_pen = QPen(_RING2, 1.5, Qt.PenStyle.SolidLine,
                      Qt.PenCapStyle.RoundCap)
        p.setPen(r2_pen)
        r2_rect = QRectF(cx - r2_r, cy - r2_r, r2_r * 2, r2_r * 2)
        p.drawArc(r2_rect, int(-ang2 * 16), 110 * 16)

        # ── "NeSk" Schriftzug ─────────────────────────────────────────────
        ty  = cy + 68.0    # Y-Position des Texts (Baseline)
        font_title = QFont("Segoe UI", 32, QFont.Weight.Light)
        p.setFont(font_title)
        fm_t = QFontMetricsF(font_title)
        text = "NeSk"
        tw   = fm_t.horizontalAdvance(text)
        tx   = cx - tw / 2

        # Basis (gedaempft weiss)
        p.setPen(QPen(QColor(180, 205, 220, 130)))
        p.drawText(QPointF(tx, ty), text)

        # Shimmer-Streifen durch den Text
        sh_prog  = ((t * 0.55) % 1.9) - 0.4     # -0.4 .. 1.5
        sh_cx    = tx + sh_prog * (tw + 60) - 15
        sh_w     = tw * 0.30
        sh_grad  = QLinearGradient(sh_cx - sh_w, ty - 38,
                                   sh_cx + sh_w, ty)
        sh_grad.setColorAt(0.0, QColor(255, 255, 255, 0))
        sh_grad.setColorAt(0.5, QColor(255, 255, 255, 210))
        sh_grad.setColorAt(1.0, QColor(255, 255, 255, 0))
        p.setPen(QPen(QBrush(sh_grad), 0))
        p.drawText(QPointF(tx, ty), text)

        # ── "NewSkies" – Klein-Schriftzug unter NeSk ─────────────────────
        font_ns = QFont("Segoe UI", 9, QFont.Weight.Light)
        font_ns.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.5)
        p.setFont(font_ns)
        fm_ns = QFontMetricsF(font_ns)
        ns_text = "NewSkies"
        ns_w    = fm_ns.horizontalAdvance(ns_text)
        p.setPen(QPen(QColor(130, 165, 190, 170)))
        p.drawText(QPointF(cx - ns_w / 2, ty + 17), ns_text)

        # ── Untertitel ────────────────────────────────────────────────────
        font_sub = QFont("Segoe UI", 9)
        p.setFont(font_sub)
        fm_s = QFontMetricsF(font_sub)
        sub  = "DRK  \u00b7  Flughafen K\u00f6ln / Bonn"
        sw   = fm_s.horizontalAdvance(sub)
        p.setPen(QPen(_ACCENT))
        p.drawText(QPointF(cx - sw / 2, ty + 36), sub)

        if self._version:
            fv = QFont("Segoe UI", 8)
            p.setFont(fv)
            fmv = QFontMetricsF(fv)
            vt  = f"v{self._version}"
            vw  = fmv.horizontalAdvance(vt)
            p.setPen(QPen(_GRAY))
            p.drawText(QPointF(cx - vw / 2, ty + 52), vt)

        # ── Trennlinie + Status ───────────────────────────────────────────
        sep_y = H - 40
        p.setPen(QPen(QColor("#1E3040")))
        p.drawLine(QPointF(28, sep_y), QPointF(W - 28, sep_y))

        font_st = QFont("Segoe UI", 8)
        p.setFont(font_st)
        p.setPen(QPen(QColor(80, 112, 138)))
        p.drawText(QPointF(28, H - 20), self._status)

        # Teal-Streifen unten
        p.fillRect(0, int(H) - 2, int(W), 2, QBrush(_RING1))

        p.end()

    # ------------------------------------------------------------------
    def set_status(self, message: str):
        """Setzt den Status-Text (thread-sicher: nur Stringzuweisung)."""
        self._status = message

    def finish(self, main_window=None):
        self.close()


# ---------------------------------------------------------------------------
# Wiederverwendbarer Lade-Dialog (animierter Doppelring-Spinner)
# ---------------------------------------------------------------------------
class LoadingDialog(QDialog):
    """
    Modaler Lade-Dialog mit animiertem Doppelring-Spinner (wie Splash Screen).
    Ersatz fuer QProgressDialog bei unbestimmter Ladedauer.

    Verwendung::
        dlg = LoadingDialog("Wird geladen …", parent)
        thread.finished.connect(dlg.accept)
        thread.start()
        dlg.exec()   # blockiert + animiert
    """
    W, H = 300, 200

    def __init__(self, status: str = "Wird geladen \u2026", parent=None):
        super().__init__(parent)
        self._status = status
        self._t0 = time.monotonic()

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Dialog
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setFixedSize(self.W, self.H)

        # In der Mitte des Hauptfensters (oder des Bildschirms) zentrieren
        ref = None
        if parent is not None:
            # Aufwärts zum obersten Fenster traversieren
            top = parent
            while top.parent() is not None:
                top = top.parent()
            ref = top
        if ref is not None:
            geo = ref.geometry()
            self.move(
                geo.x() + (geo.width()  - self.W) // 2,
                geo.y() + (geo.height() - self.H) // 2,
            )
        else:
            screen = QApplication.primaryScreen()
            if screen:
                sg = screen.availableGeometry()
                self.move(sg.center().x() - self.W // 2,
                          sg.center().y() - self.H // 2)

        self._timer = QTimer(self)
        self._timer.timeout.connect(self.repaint)
        self._timer.start(30)   # ~33 FPS

    # ------------------------------------------------------------------
    def set_status(self, text: str):
        self._status = text

    # ------------------------------------------------------------------
    def paintEvent(self, event):          # noqa: N802
        t = time.monotonic() - self._t0
        W, H = float(self.W), float(self.H)

        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        # Hintergrund (radialer Gradient, wie Splash)
        from PySide6.QtGui import QRadialGradient
        bg = QRadialGradient(W / 2, H * 0.44, W * 0.58)
        bg.setColorAt(0.0, QColor("#182634"))
        bg.setColorAt(1.0, QColor("#0A1520"))
        p.fillRect(0, 0, int(W), int(H), QBrush(bg))

        # Goldstreifen oben
        p.fillRect(0, 0, int(W), 3, QBrush(QColor("#C0944A")))

        # Ring-Mittelpunkt
        cx, cy = W / 2, H * 0.40

        ang1 = (t * 120.0) % 360.0          # Teal-Ring vorwaerts
        ang2 = (180.0 - t * 75.0) % 360.0  # Gold-Ring rueckwaerts

        # Pulsierender Glow
        import math as _math
        glow_alpha = int(18 + 14 * _math.sin(t * 2.2))
        glow_r = 38.0
        glow_grad = QRadialGradient(cx, cy, glow_r + 16)
        glow_grad.setColorAt(0.0, QColor(91, 138, 170, glow_alpha * 2))
        glow_grad.setColorAt(0.6, QColor(91, 138, 170, glow_alpha))
        glow_grad.setColorAt(1.0, QColor(91, 138, 170, 0))
        p.setBrush(QBrush(glow_grad))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - glow_r - 16, cy - glow_r - 16,
                             (glow_r + 16) * 2, (glow_r + 16) * 2))

        # Innerer dunkler Kreis
        inner_r = 26.0
        p.setBrush(QBrush(QColor("#111E2A")))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - inner_r, cy - inner_r,
                             inner_r * 2, inner_r * 2))

        # "N" in der Mitte
        from PySide6.QtGui import QFontMetricsF as _FMF
        font_n = QFont("Segoe UI", 16, QFont.Weight.Light)
        p.setFont(font_n)
        fm_n = _FMF(font_n)
        nw = fm_n.horizontalAdvance("N")
        p.setPen(QPen(_RING1))
        p.drawText(QPointF(cx - nw / 2, cy + fm_n.ascent() / 2 - 1), "N")

        # Ring 1 (Teal, gross, vorwaerts)
        r1_r = 34.0
        p.setPen(QPen(_RING1, 2.5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawArc(QRectF(cx - r1_r, cy - r1_r, r1_r * 2, r1_r * 2),
                  int(-ang1 * 16), 230 * 16)

        # Ring 2 (Gold, kleiner, rueckwaerts)
        r2_r = 42.0
        p.setPen(QPen(_RING2, 1.5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.drawArc(QRectF(cx - r2_r, cy - r2_r, r2_r * 2, r2_r * 2),
                  int(-ang2 * 16), 110 * 16)

        # Trennlinie + Status-Text
        sep_y = H - 45.0
        p.setPen(QPen(QColor("#1E3040")))
        p.drawLine(QPointF(20, sep_y), QPointF(W - 20, sep_y))

        font_st = QFont("Segoe UI", 9)
        p.setFont(font_st)
        fm_st = _FMF(font_st)
        sw = fm_st.horizontalAdvance(self._status)
        p.setPen(QPen(QColor(80, 112, 138)))
        p.drawText(QPointF(cx - sw / 2, H - 20), self._status)

        # Teal-Streifen unten
        p.fillRect(0, int(H) - 2, int(W), 2, QBrush(_RING1))
        p.end()

    def closeEvent(self, event):  # noqa: N802
        self._timer.stop()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Generischer Hintergrundthread + Hilfsfunktion für DB-Ladevorgänge
# ---------------------------------------------------------------------------
class _LadeThread(QThread):
    """Übergibt fn(*args, **kwargs) an einen Hintergrundthread und emittiert das Ergebnis."""
    fertig = Signal(object)
    fehler = Signal(Exception)

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self._fn     = fn
        self._args   = args
        self._kwargs = kwargs

    def run(self):
        try:
            self.fertig.emit(self._fn(*self._args, **self._kwargs))
        except Exception as exc:
            self.fehler.emit(exc)


def _mit_ladeanimation(parent, status: str, fn, *args, **kwargs):
    """
    Führt fn(*args, **kwargs) im Hintergrundthread aus und zeigt dabei
    den animierten NeSk-Spinner-Dialog.  Gibt (ergebnis, fehler) zurück.

    Beispiel::
        daten, exc = _mit_ladeanimation(self, "Daten laden …", db_funktion, param=wert)
        if exc:
            QMessageBox.critical(self, "Fehler", str(exc))
            return
    """
    ergebnis   = [None]
    fehler_ref = [None]

    dlg    = LoadingDialog(status, parent)
    thread = _LadeThread(fn, *args, **kwargs)
    thread.fertig.connect(lambda r: ergebnis.__setitem__(0, r))
    thread.fehler.connect(lambda e: fehler_ref.__setitem__(0, e))
    thread.finished.connect(dlg.accept)
    thread.start()
    dlg.exec()
    thread.wait()

    return ergebnis[0], fehler_ref[0]