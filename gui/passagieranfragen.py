"""
Passagieranfragen-Widget
Passagier-E-Mails verarbeiten, Daten extrahieren und
Antworten per Outlook (win32com) versenden.
"""
from __future__ import annotations

import os
import re
from pathlib import Path

from PySide6.QtCore import Qt, QDateTime, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton,
    QScrollArea, QSizePolicy, QSplitter, QTableWidget, QTableWidgetItem, QTextEdit,
    QVBoxLayout, QWidget,
)

from config import BASE_DIR, FIORI_BLUE, FIORI_TEXT

_LOGO_PATH = str(Path(BASE_DIR) / "Daten" / "Email" / "Logo.jpg")


# ── Outlook-Posteingang-Dialog ──────────────────────────────────────────────────────────

class OutlookInboxDialog(QDialog):
    """Zeigt die letzten Posteingang-E-Mails und lädt die gewählte zurück."""

    def __init__(self, parent=None, max_items: int = 75):
        super().__init__(parent)
        self.setWindowTitle("📬  Outlook-Posteingang")
        self.resize(860, 480)
        self.selected_body: str = ""
        self.selected_sender_email: str = ""
        self.selected_sender_name: str = ""
        self._items: list = []   # (datum_str, absender_name, absender_email, betreff, body)

        self._load_mails(max_items)
        self._setup_ui()

    def _load_mails(self, max_items: int):
        try:
            import win32com.client
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            inbox = ns.GetDefaultFolder(6)   # 6 = olFolderInbox
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)   # neueste zuerst

            count = 0
            for msg in messages:
                if count >= max_items:
                    break
                try:
                    datum = str(msg.ReceivedTime)[:16] if msg.ReceivedTime else ""
                    sender_name  = (msg.SenderName or "").strip()
                    sender_email = (msg.SenderEmailAddress or "").strip()
                    # Exchange-interne Adressen (EX:/CN=...) überspringen – SMTP bevorzugen
                    if sender_email.upper().startswith(("EX:", "/O=", "/CN=")):
                        sender_email = ""
                    # Versuche Antwort-An-Adresse als SMTP-Fallback
                    if not sender_email:
                        try:
                            sender_email = msg.ReplyRecipients(1).Address or ""
                        except Exception:
                            pass
                    absender_display = sender_name or sender_email
                    betreff  = msg.Subject or ""
                    # Reinen Text bevorzugen; bei HTML-only Body aus HTMLBody extrahieren
                    body = (msg.Body or "").strip()
                    if not body:
                        import re as _re
                        body = _re.sub(r'<[^>]+>', ' ', msg.HTMLBody or "").strip()
                    self._items.append((datum, absender_display, sender_email, betreff, body))
                    count += 1
                except Exception:
                    pass
        except Exception as exc:
            self._items = []
            self._load_error = str(exc)
        else:
            self._load_error = ""

    def _setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        if self._load_error:
            err = QLabel(f"⚠️  Outlook konnte nicht gelesen werden:\n{self._load_error}")
            err.setWordWrap(True)
            err.setStyleSheet("color: #c0392b; font-size: 11px;")
            lay.addWidget(err)
        else:
            info = QLabel(
                f"{len(self._items)} E-Mails geladen — "
                "Doppelklick oder Auswählen + OK um eine E-Mail zu übernehmen."
            )
            info.setStyleSheet("color: #555; font-size: 11px;")
            lay.addWidget(info)

        self._table = QTableWidget(len(self._items), 3)
        self._table.setHorizontalHeaderLabels(["Datum", "Von", "Betreff"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setStyleSheet("font-size: 11px;")

        for row, (datum, absender_display, _sender_email, betreff, _body) in enumerate(self._items):
            self._table.setItem(row, 0, QTableWidgetItem(datum))
            self._table.setItem(row, 1, QTableWidgetItem(absender_display))
            self._table.setItem(row, 2, QTableWidgetItem(betreff))

        self._table.cellDoubleClicked.connect(self._accept_row)
        lay.addWidget(self._table, stretch=1)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("✅  E-Mail übernehmen")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Abbrechen")
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _accept_row(self, row: int, _col: int):
        self._table.selectRow(row)
        self._on_ok()

    def _on_ok(self):
        rows = self._table.selectionModel().selectedRows()
        if not rows:
            return
        row = rows[0].row()
        _datum, sender_name, sender_email, _betreff, body = self._items[row]
        self.selected_body         = body
        self.selected_sender_email = sender_email
        self.selected_sender_name  = sender_name
        self.accept()


# ── Hilfsfunktion: Deutschen E-Mail-Text in strukturierte Felder aufteilen ──────

def _parse_email_fields(text: str) -> dict:
    """
    Versucht aus einem deutschen E-Mail-Freitext folgende Felder zu extrahieren:
    name, email, flugnummer, datum, rueckflug

    Strategie (Priorität oben → unten):
    1. Explizite Label-Zeilen: "Vorname:", "Nachname:", "Name:", "Von:", "Absender:"
    2. Anredezeilen: "Anrede: Herr" + "Vorname: Hans" + "Nachname: ..."
    3. "Herr/Frau Vorname Nachname" im Fließtext
    4. From-Header / Absenderzeile
    """
    result = {"name": "", "anrede": "", "email": "", "flugnummer": "", "datum": "", "rueckflug": ""}

    # ── Anrede ──────────────────────────────────────────────────────────────────
    m_anr = re.search(r'Anrede[:\s]+(Herr|Frau)', text, re.IGNORECASE)
    if m_anr:
        result["anrede"] = m_anr.group(1).capitalize()

    lines = text.splitlines()

    # ── E-Mail ──────────────────────────────────────────────────────────────
    m = re.search(r'[\w.+\-]+@[\w\-]+\.[a-zA-Z]{2,}', text)
    result["email"] = m.group(0) if m else ""

    # ── Flugnummer  (EW583, FR 1234, LH456, 4U100) ──────────────────────────
    m = re.search(r'\b([A-Z0-9]{2})\s*(\d{3,4})\b', text)
    result["flugnummer"] = (m.group(1) + m.group(2)) if m else ""

    # ── Datum + Rückflug ─────────────────────────────────────────────────────
    _MONATE = {
        'januar': '01', 'februar': '02', 'märz': '03', 'april': '04',
        'mai': '05', 'juni': '06', 'juli': '07', 'august': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'dezember': '12',
    }
    dates: list[str] = []
    # dd.MM.yyyy / dd/MM/yyyy / dd-MM-yyyy
    for raw in re.finditer(r'\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\b', text):
        d, mo, y = raw.group(1), raw.group(2), raw.group(3)
        if len(y) == 2:
            y = "20" + y
        dates.append(f"{int(d):02d}.{int(mo):02d}.{y}")
    # "12. März 2026" / "12 März 2026"
    if not dates:
        for raw2 in re.finditer(
            r'(\d{1,2})[.\s]+(januar|februar|märz|april|mai|juni|juli|august|'
            r'september|oktober|november|dezember)[,.\s]+(\d{4})',
            text, re.IGNORECASE,
        ):
            mo_key = raw2.group(2).lower()
            dates.append(f"{int(raw2.group(1)):02d}.{_MONATE[mo_key]}.{raw2.group(3)}")
    result["datum"] = dates[0] if dates else ""
    result["rueckflug"] = dates[1] if len(dates) >= 2 else ""

    # Explizites Rückflug-Label überschreibt zweites Datum
    m_r = re.search(r'[Rr]ückflug[:\s]+(.{3,40}?)(?:\n|,|\.|$)', text)
    if m_r:
        result["rueckflug"] = m_r.group(1).strip()

    # ── Name ─────────────────────────────────────────────────────────────────
    # Strategie 1: explizite Zeilen "Vorname: ...", "Nachname: ..."
    vorname = ""
    nachname = ""
    for line in lines:
        if re.match(r'(?:Vorname|First\s*name)[:\s]+(.+)', line, re.IGNORECASE):
            vorname = re.split(r':\s*', line, 1)[-1].strip()
        if re.match(r'(?:Nachname|Last\s*name|Familienname)[:\s]+(.+)', line, re.IGNORECASE):
            nachname = re.split(r':\s*', line, 1)[-1].strip()
    if vorname or nachname:
        result["name"] = f"{vorname} {nachname}".strip()
        return result

    # Strategie 2: Label "Name: ..." in einer Zeile
    m_n = re.search(
        r'(?:^|\n)\s*(?:Name|Passagier)[:\s]+([A-ZÄÖÜ][^\n,]{2,40})',
        text, re.IGNORECASE,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()
        return result

    # Strategie 3: Anrede-Block "Anrede: Herr\nVorname: Hans\nNachname: Muster"
    anrede_idx = -1
    for i, line in enumerate(lines):
        if re.match(r'Anrede\s*:', line, re.IGNORECASE):
            anrede_idx = i
            break
    if anrede_idx >= 0:
        for line in lines[anrede_idx + 1: anrede_idx + 5]:
            lm = re.match(r'(?:Vorname|Nachname)[:\s]+(.+)', line, re.IGNORECASE)
            if lm:
                part = lm.group(1).strip()
                if not vorname:
                    vorname = part
                else:
                    nachname = part
        if vorname or nachname:
            result["name"] = f"{vorname} {nachname}".strip()
            return result

    # Strategie 4: "Herr/Frau Vorname Nachname" im Fließtext
    m_n = re.search(
        r'\b(?:Herr(?:n)?|Frau)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)+)',
        text,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()
        return result

    # Strategie 5: "Von: Hans Muster <...>" (E-Mail-Header)
    m_n = re.search(
        r'^Von:\s*([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)*)\s*[<\(]',
        text, re.MULTILINE,
    )
    if m_n:
        result["name"] = m_n.group(1).strip()

    return result


# ── Antwort-Vorlagen ───────────────────────────────────────────────────────────

_FLUGDATEN_BITTE = (
    "\n\nFür die Bearbeitung Ihrer Anfrage benötigen wir noch folgende Angaben:\n"
    "• Abflugdatum und -uhrzeit\n"
    "• Flugnummer und Reiseziel\n"
    "• Vor- und Nachname der zu betreuenden Person\n"
    "• Art der Einschränkung (WCH-R, WCH-S oder WCH-C)"
)

_SIG = (
    "\n\nMit freundlichen Grüßen\n\nIhr Team vom PRM-Service\n"
    "Am Köln-Bonn-Airport · Kennedystraße · 51147 Köln\n\n"
    "Telefon: +49 2203 40 - 2323  (24 Stunden täglich erreichbar)\n"
    "E-Mail:  flughafen@drk-koeln.de"
)

ANTWORT_KOMPLETT = (
    "Sehr geehrte Damen und Herren,\n\n"
    "wir haben die Passagiere in unserem System eingetragen.\n\n"
    "Die Buchung des PRM Services erfolgt jedoch nur für den Flughafen Köln / Bonn.\n"
    "Bei Buchungen, die nicht über die Airline erfolgen, kann sich die Airline das Recht "
    "nehmen, den PRM Service vor Ort abzulehnen.\n\n"
    "Um einen reibungslosen Ablauf zu gewährleisten, sind folgende Dinge zu beachten:\n\n"
    "• Bitte kommen Sie mindestens 2–2,5 Stunden vor Abflug am Flughafen Köln/Bonn an.\n"
    "• Melden Sie sich beim zuständigen Check-In-Schalter und weisen Sie darauf hin,\n"
    "  dass Sie den PRM Service in Anspruch nehmen möchten.\n"
    "  Wichtig: Die Bestätigung muss über den Check-In erfolgen, da sonst keine\n"
    "  Abholung am Service-Point stattfindet.\n"
    "• Das Check-In-Personal verweist Sie zu einem Service-Point. Dort holen wir Sie\n"
    "  in der Regel eine Stunde vor Abflug ab. In seltenen Fällen kann es zu leichten "
    "Verzögerungen kommen.\n\n"
    "Bei Fragen stehen wir Ihnen jederzeit zur Verfügung – unser Team ist 24 Stunden "
    "täglich für Sie erreichbar. Gerne beantworten wir Ihre Fragen auch telefonisch."
) + _SIG

ANTWORT_FEHLENDE_DATEN = (
    "Sehr geehrte Damen und Herren,\n\n"
    "gerne sind wir Ihnen bei Ihrem Flug behilflich. Leider liegen uns derzeit noch "
    "nicht alle erforderlichen Informationen vor.\n\n"
    "Für die Organisation des Services benötigen wir noch folgende Angaben:\n\n"
    "• Abflugdatum und -uhrzeit\n"
    "• Flugnummer und Reiseziel\n"
    "• Vor- und Nachname der zu betreuenden Person\n"
    "• Art der Einschränkung (WCH-R, WCH-S oder WCH-C)\n\n"
    "Sobald uns diese Informationen vorliegen, kümmern wir uns umgehend um die "
    "weitere Koordination.\n\n"
    "Unser Team ist 24 Stunden täglich für Sie erreichbar – zögern Sie nicht, "
    "uns auch telefonisch zu kontaktieren."
) + _SIG

ANTWORT_PARKPLATZ = (
    "Sehr geehrte Damen und Herren,\n\n"
    "vielen Dank für Ihre Anfrage.\n\n"
    "Eine Abholung direkt am Parkplatz ist kein Problem. Bitte rufen Sie uns am "
    "Reisetag nochmals an, damit wir die genaue Abholung koordinieren können.\n\n"
    "Unser Team ist 24 Stunden täglich für Sie erreichbar und freut sich auf Ihren Anruf."
) + _SIG

ANTWORT_INFO_SERVICE = (
    "Sehr geehrte Damen und Herren,\n\n"
    "vielen Dank für Ihre Anfrage zum PRM-Service am Flughafen Köln/Bonn.\n\n"
    "Der PRM-Service (Persons with Reduced Mobility) steht allen Reisenden zur "
    "Verfügung, die am Flughafen Unterstützung benötigen – ob aufgrund körperlicher "
    "Einschränkungen, vorübergehender Verletzungen oder aus anderen Gründen.\n\n"
    "So funktioniert der PRM-Service am Flughafen Köln/Bonn:\n\n"
    "1. Anmeldung\n"
    "   Bitte melden Sie den Bedarf möglichst frühzeitig bei Ihrer Airline an.\n"
    "   Die Airline informiert uns direkt über Ihre Anforderung.\n\n"
    "2. Ankunft am Flughafen\n"
    "   Bitte kommen Sie mindestens 2–2,5 Stunden vor Abflug am Flughafen an.\n\n"
    "3. Check-In\n"
    "   Melden Sie sich beim zuständigen Check-In-Schalter und weisen Sie auf den\n"
    "   gebuchten PRM-Service hin. Die Bestätigung durch den Check-In ist zwingend\n"
    "   erforderlich, da sonst keine Abholung am Service-Point erfolgt.\n\n"
    "4. Abholung am Service-Point\n"
    "   Das Check-In-Personal verweist Sie zu einem unserer Service-Points.\n"
    "   Dort holen wir Sie in der Regel rund eine Stunde vor Abflug ab.\n\n"
    "5. Begleitung\n"
    "   Unser Team begleitet Sie durch alle Sicherheitskontrollen, zum Gate und –\n"
    "   falls gewünscht – bis an Bord des Flugzeugs.\n\n"
    "Gerne beantworten wir Ihre Fragen auch telefonisch – unser Team ist "
    "24 Stunden täglich für Sie erreichbar."
) + _SIG


# ── Szenario 5: PRM-Ablehnung durch Airline ──────────────────────────────────

def _build_prm_ablehnung(
    kontakte: list[dict],
    prm_passagiere: list[str],
    flugnummer: str,
    datum: str,
    airline: str,
) -> str:
    """Erzeugt den vollständigen E-Mail-Text für eine PRM-Ablehnung."""

    # Anrede-Zeile
    if len(kontakte) == 1:
        k = kontakte[0]
        anr, vor = k["anrede"], k["vorname"]
        if vor and anr == "Frau":
            begruessung = f"Sehr geehrte Frau {vor},"
        elif vor and anr == "Herr":
            begruessung = f"Sehr geehrter Herr {vor},"
        elif vor:
            begruessung = f"Sehr geehrte/r {vor},"
        else:
            begruessung = "Sehr geehrte Damen und Herren,"
    elif len(kontakte) > 1:
        parts = []
        for k in kontakte:
            if k["vorname"] and k["anrede"] in ("Herr", "Frau"):
                parts.append(f"{k['anrede']} {k['vorname']}")
            elif k["vorname"]:
                parts.append(k["vorname"])
        if parts:
            joined = (
                f"{parts[0]} und {parts[1]}" if len(parts) == 2
                else ", ".join(parts[:-1]) + f" und {parts[-1]}"
            )
            begruessung = f"Sehr geehrte {joined},"
        else:
            begruessung = "Sehr geehrte Damen und Herren,"
    else:
        begruessung = "Sehr geehrte Damen und Herren,"

    # Flug-Referenz
    if flugnummer and datum:
        flug_ref = f"{flugnummer} am {datum}"
    elif flugnummer:
        flug_ref = flugnummer
    elif datum:
        flug_ref = f"der Flug am {datum}"
    else:
        flug_ref = "der genannte Flug"

    # PRM-Passagier(e)
    if not prm_passagiere:
        schilderung_passagier = "für den Flug"
    elif len(prm_passagiere) == 1:
        schilderung_passagier = f"für den Flug von {prm_passagiere[0]}"
    else:
        joined_p = (
            f"{prm_passagiere[0]} und {prm_passagiere[1]}" if len(prm_passagiere) == 2
            else ", ".join(prm_passagiere[:-1]) + f" und {prm_passagiere[-1]}"
        )
        schilderung_passagier = f"für den Flug von {joined_p}"

    airline_str = airline if airline else "der Fluggesellschaft"

    schilderung = (
        f"Sie schildern, dass {flug_ref} die Anzahl der PRM-Plätze begrenzt hat "
        f"und das Kontingent {schilderung_passagier} bereits ausgeschöpft ist. "
        "Die Festlegung der PRM-Kontingente sowie die Entscheidung, ob eine weitere "
        "Anmeldung angenommen oder abgelehnt wird, liegt ausschließlich in der "
        "Verantwortung der Fluggesellschaft. Auf diese Entscheidung haben weder "
        "der Flughafen Köln/Bonn noch wir als Deutsches Rotes Kreuz Einfluss."
    )

    return (
        f"{begruessung}\n\n"
        "vielen Dank für Ihre Nachricht.\n\n"
        "Grundsätzlich muss der PRM-Service über die jeweilige Fluggesellschaft angemeldet "
        "werden. Die Airline nimmt die Anmeldung entgegen und übermittelt diese anschließend "
        "an uns, das Deutsche Rote Kreuz. Wir führen den Service am Flughafen Köln/Bonn im "
        "Auftrag des Flughafens durch.\n\n"
        "Wir können den Service daher nur auf Grundlage einer entsprechenden Meldung durch die "
        "Fluggesellschaft leisten. Eine eigene Buchung am System der Airline vorbei oder eine "
        "Änderung der durch die Fluggesellschaft festgelegten Kontingente ist uns leider nicht "
        "möglich.\n\n"
        f"{schilderung}\n\n"
        f"Bitte wenden Sie sich daher bezüglich einer möglichen Lösung oder Umbuchung direkt an "
        f"{airline_str}. Sie haben in diesem Fall allerdings das Recht, kostenlos umzubuchen. "
        "Die weitere Klärung und Umsetzung der Umbuchung muss direkt über die Fluggesellschaft "
        "erfolgen.\n\n"
        "Unabhängig davon können wir Ihnen, sofern verfügbar, einen Leihrollstuhl zur Verfügung "
        "stellen. Mit diesem könnten Sie Ihre/n Angehörige/n selbstständig bis zum Gate "
        "begleiten.\n\n"
        "Bitte beachten Sie hierbei:\n\n"
        "• Der Leihrollstuhl ersetzt nicht den von der Airline zu organisierenden PRM-Service.\n"
        "• Wir können in diesem Fall keine Verantwortung für den Transport bis zum Flugzeug "
        "oder für die Beförderung durch die Fluggesellschaft übernehmen.\n"
        "• Die endgültige Entscheidung zur Mitnahme liegt weiterhin bei der jeweiligen "
        "Fluggesellschaft.\n\n"
        "Bitte melden Sie sich gerne bei uns, falls ein Leihrollstuhl für Sie infrage kommt.\n\n"
        "Für weitere Fragen stehen wir Ihnen selbstverständlich gerne zur Verfügung."
    ) + _SIG


class _KontaktRow(QWidget):
    """Zeile: Anrede-Combo + Vorname + E-Mail + Entfernen-Button."""

    removed = Signal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 2, 0, 2)
        lay.setSpacing(4)

        self._anrede = QComboBox()
        self._anrede.addItems(["–", "Herr", "Frau"])
        self._anrede.setFixedWidth(68)
        self._anrede.setStyleSheet(
            "QComboBox{border:1px solid #c8d0d8;border-radius:3px;padding:2px 6px;font-size:11px;}"
            "QComboBox:focus{border-color:#0070f3;}"
        )

        self._vorname = QLineEdit()
        self._vorname.setPlaceholderText("Vorname")
        self._vorname.setStyleSheet(
            "QLineEdit{border:1px solid #c8d0d8;border-radius:3px;padding:2px 6px;font-size:11px;}"
            "QLineEdit:focus{border-color:#0070f3;}"
        )

        self._email = QLineEdit()
        self._email.setPlaceholderText("E-Mail-Adresse")
        self._email.setStyleSheet(
            "QLineEdit{border:1px solid #c8d0d8;border-radius:3px;padding:2px 6px;font-size:11px;}"
            "QLineEdit:focus{border-color:#0070f3;}"
        )

        btn = QPushButton("✕")
        btn.setFixedSize(26, 26)
        btn.setStyleSheet(
            "QPushButton{background:#e74c3c;color:white;border:none;border-radius:3px;font-size:11px;}"
            "QPushButton:hover{background:#c0392b;}"
        )
        btn.clicked.connect(lambda: self.removed.emit(self))

        lay.addWidget(self._anrede)
        lay.addWidget(self._vorname, 1)
        lay.addWidget(self._email, 2)
        lay.addWidget(btn)

    def data(self) -> dict:
        return {
            "anrede": self._anrede.currentText(),
            "vorname": self._vorname.text().strip(),
            "email": self._email.text().strip(),
        }

    def set_email(self, email: str):
        self._email.setText(email)


class _PassagierRow(QWidget):
    """Zeile: Name + Entfernen-Button."""

    removed = Signal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 2, 0, 2)
        lay.setSpacing(4)

        self._name = QLineEdit()
        self._name.setPlaceholderText("Vollständiger Name")
        self._name.setStyleSheet(
            "QLineEdit{border:1px solid #c8d0d8;border-radius:3px;padding:2px 6px;font-size:11px;}"
            "QLineEdit:focus{border-color:#0070f3;}"
        )

        btn = QPushButton("✕")
        btn.setFixedSize(26, 26)
        btn.setStyleSheet(
            "QPushButton{background:#e74c3c;color:white;border:none;border-radius:3px;font-size:11px;}"
            "QPushButton:hover{background:#c0392b;}"
        )
        btn.clicked.connect(lambda: self.removed.emit(self))

        lay.addWidget(self._name, 1)
        lay.addWidget(btn)

    def data(self) -> str:
        return self._name.text().strip()


class Szenario5Dialog(QDialog):
    """Eingabe-Dialog für Szenario 5 – Ablehnung PRM durch Airline."""

    _CC_FIXED = "Flughafen2@drk-koeln.de; hildegard.eichler@koeln-bonn-airport.de"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🚫  Szenario 5 – Ablehnung PRM durch Airline")
        self.setMinimumWidth(700)
        self.resize(720, 660)
        self._kontakt_rows: list[_KontaktRow] = []
        self._passagier_rows: list[_PassagierRow] = []

        # Ergebniswerte
        self.result_to = ""
        self.result_cc = self._CC_FIXED
        self.result_subject = ""
        self.result_body = ""

        self._setup_ui()
        self._add_kontakt()
        self._add_passagier()

    # ── UI ──────────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(0, 0, 4, 0)
        inner_lay.setSpacing(10)
        scroll.setWidget(inner)
        outer.addWidget(scroll, stretch=1)

        _GRP = (
            "QGroupBox{border:1px solid #c8d0d8;border-radius:6px;margin-top:14px;"
            "padding-top:6px;font-size:11px;font-weight:bold;color:#2c3e50;}"
            "QGroupBox::title{subcontrol-origin:margin;subcontrol-position:top left;"
            "left:10px;padding:0 6px;color:#1565a8;}"
        )

        # ── Kontaktpersonen ───────────────────────────────────────────────────
        grp_k = QGroupBox("👤  Kontaktpersonen  (Empfänger der E-Mail)")
        grp_k.setStyleSheet(_GRP)
        k_lay = QVBoxLayout(grp_k)
        k_lay.setSpacing(4)

        hdr_k = QHBoxLayout()
        for txt, stretch in [("Anrede", 0), ("Vorname", 1), ("E-Mail", 2)]:
            lbl = QLabel(txt)
            lbl.setStyleSheet("color:#555;font-size:10px;font-weight:bold;")
            if stretch:
                hdr_k.addWidget(lbl, stretch)
            else:
                lbl.setFixedWidth(72)
                hdr_k.addWidget(lbl)
        hdr_k.addSpacing(30)
        k_lay.addLayout(hdr_k)

        self._kontakt_container = QVBoxLayout()
        self._kontakt_container.setSpacing(2)
        k_lay.addLayout(self._kontakt_container)

        btn_add_k = QPushButton("＋  Kontaktperson hinzufügen")
        btn_add_k.setStyleSheet(
            "QPushButton{background:#ecf0f1;border:1px dashed #aaa;border-radius:4px;"
            "padding:4px 10px;font-size:11px;color:#555;}"
            "QPushButton:hover{background:#dfe6e9;}"
        )
        btn_add_k.clicked.connect(self._add_kontakt)
        k_lay.addWidget(btn_add_k)
        inner_lay.addWidget(grp_k)

        # ── PRM-Passagiere ────────────────────────────────────────────────────
        grp_p = QGroupBox("♿  PRM-Passagiere  (vom Service abgelehnte Personen)")
        grp_p.setStyleSheet(_GRP)
        p_lay = QVBoxLayout(grp_p)
        p_lay.setSpacing(4)

        hdr_p = QHBoxLayout()
        lbl_n = QLabel("Vollständiger Name")
        lbl_n.setStyleSheet("color:#555;font-size:10px;font-weight:bold;")
        hdr_p.addWidget(lbl_n, 1)
        hdr_p.addSpacing(30)
        p_lay.addLayout(hdr_p)

        self._passagier_container = QVBoxLayout()
        self._passagier_container.setSpacing(2)
        p_lay.addLayout(self._passagier_container)

        btn_add_p = QPushButton("＋  PRM-Passagier hinzufügen")
        btn_add_p.setStyleSheet(
            "QPushButton{background:#ecf0f1;border:1px dashed #aaa;border-radius:4px;"
            "padding:4px 10px;font-size:11px;color:#555;}"
            "QPushButton:hover{background:#dfe6e9;}"
        )
        btn_add_p.clicked.connect(self._add_passagier)
        p_lay.addWidget(btn_add_p)
        inner_lay.addWidget(grp_p)

        # ── Flugdaten ─────────────────────────────────────────────────────────
        grp_f = QGroupBox("✈️  Flugdaten & Airline")
        grp_f.setStyleSheet(_GRP)
        f_lay = QHBoxLayout(grp_f)
        f_lay.setSpacing(10)

        _LS = (
            "QLineEdit{border:1px solid #c8d0d8;border-radius:3px;padding:3px 8px;"
            "font-size:11px;} QLineEdit:focus{border-color:#0070f3;}"
        )

        f_lay.addWidget(QLabel("Flugnummer:"))
        self._e_flug = QLineEdit()
        self._e_flug.setPlaceholderText("z. B. PC 1608")
        self._e_flug.setStyleSheet(_LS)
        f_lay.addWidget(self._e_flug, 1)

        f_lay.addWidget(QLabel("Datum:"))
        self._e_datum = QLineEdit()
        self._e_datum.setPlaceholderText("z. B. 04.06.2026")
        self._e_datum.setStyleSheet(_LS)
        f_lay.addWidget(self._e_datum, 1)

        f_lay.addWidget(QLabel("Airline:"))
        self._e_airline = QLineEdit()
        self._e_airline.setPlaceholderText("z. B. Pegasus Airlines")
        self._e_airline.setStyleSheet(_LS)
        f_lay.addWidget(self._e_airline, 2)
        inner_lay.addWidget(grp_f)

        # ── E-Mail-Felder ─────────────────────────────────────────────────────
        grp_e = QGroupBox("📧  E-Mail-Versand")
        grp_e.setStyleSheet(_GRP)
        e_lay = QVBoxLayout(grp_e)
        e_lay.setSpacing(6)

        row_to = QHBoxLayout()
        lbl_to = QLabel("An:")
        lbl_to.setFixedWidth(30)
        lbl_to.setStyleSheet("font-weight:bold;font-size:11px;")
        self._e_to = QLineEdit()
        self._e_to.setPlaceholderText("Empfänger-E-Mail (aus Kontaktperson oder manuell)")
        self._e_to.setStyleSheet(_LS)
        row_to.addWidget(lbl_to)
        row_to.addWidget(self._e_to)
        e_lay.addLayout(row_to)

        row_cc = QHBoxLayout()
        lbl_cc = QLabel("CC:")
        lbl_cc.setFixedWidth(30)
        lbl_cc.setStyleSheet("font-weight:bold;font-size:11px;")
        self._e_cc = QLineEdit(self._CC_FIXED)
        self._e_cc.setStyleSheet(_LS)
        row_cc.addWidget(lbl_cc)
        row_cc.addWidget(self._e_cc)
        e_lay.addLayout(row_cc)
        inner_lay.addWidget(grp_e)

        # ── Buttons ───────────────────────────────────────────────────────────
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("📧  Entwurf übernehmen")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Abbrechen")
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        outer.addWidget(btns)

    # ── Zeilen verwalten ────────────────────────────────────────────────────────

    def _add_kontakt(self):
        row = _KontaktRow(self)
        row.removed.connect(self._remove_kontakt)
        self._kontakt_rows.append(row)
        self._kontakt_container.addWidget(row)

    def _remove_kontakt(self, row: _KontaktRow):
        if len(self._kontakt_rows) <= 1:
            return  # mindestens eine Zeile behalten
        self._kontakt_rows.remove(row)
        self._kontakt_container.removeWidget(row)
        row.deleteLater()
        self._sync_to_from_kontakte()

    def _add_passagier(self):
        row = _PassagierRow(self)
        row.removed.connect(self._remove_passagier)
        self._passagier_rows.append(row)
        self._passagier_container.addWidget(row)

    def _remove_passagier(self, row: _PassagierRow):
        if len(self._passagier_rows) <= 1:
            return
        self._passagier_rows.remove(row)
        self._passagier_container.removeWidget(row)
        row.deleteLater()

    def _sync_to_from_kontakte(self):
        """Füllt das An-Feld automatisch mit den E-Mails der Kontaktpersonen."""
        emails = [r.data()["email"] for r in self._kontakt_rows if r.data()["email"]]
        self._e_to.setText("; ".join(emails))

    # ── Validierung & Ergebnis ──────────────────────────────────────────────────

    def _on_ok(self):
        kontakte = [r.data() for r in self._kontakt_rows]
        passagiere = [r.data() for r in self._passagier_rows if r.data()]
        flugnummer = self._e_flug.text().strip()
        datum = self._e_datum.text().strip()
        airline = self._e_airline.text().strip()

        # An-Feld: falls leer, aus Kontaktpersonen befüllen
        to = self._e_to.text().strip()
        if not to:
            emails = [k["email"] for k in kontakte if k["email"]]
            to = "; ".join(emails)

        if not to:
            QMessageBox.warning(
                self, "E-Mail fehlt",
                "Bitte mindestens eine E-Mail-Adresse im Feld 'An:' oder bei einer Kontaktperson eintragen."
            )
            return

        subject_parts = ["PRM – Ablehnung durch Airline"]
        if passagiere:
            subject_parts.append(passagiere[0])
        if flugnummer:
            subject_parts.append(f"Flug {flugnummer}")
        if datum:
            subject_parts.append(datum)

        self.result_to = to
        self.result_cc = self._e_cc.text().strip()
        self.result_subject = " | ".join(subject_parts)
        self.result_body = _build_prm_ablehnung(kontakte, passagiere, flugnummer, datum, airline)
        self.accept()


# ── Widget ─────────────────────────────────────────────────────────────────────

class PassagieranfragenWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    # ── UI-Aufbau ──────────────────────────────────────────────────────────────

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 10, 16, 10)
        root.setSpacing(8)

        # Titelzeile
        title = QLabel("✉️  Passagieranfragen")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Light))
        title.setStyleSheet(f"color: {FIORI_TEXT};")
        root.addWidget(title)

        hint = QLabel(
            "E-Mail einfügen → Daten extrahieren → Szenario wählen → Outlook-Entwurf öffnen"
        )
        hint.setStyleSheet("color: #666; font-size: 11px;")
        root.addWidget(hint)

        # Haupt-Splitter: Links Eingabe / Rechts Ausgabe
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, stretch=1)

        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setSizes([440, 720])

    def _build_left(self) -> QWidget:
        left = QWidget()
        layout = QVBoxLayout(left)
        layout.setContentsMargins(0, 0, 8, 0)
        layout.setSpacing(6)

        # Eingabebereich
        grp_in = QGroupBox("📥  Passagier-E-Mail einfügen")
        grp_in.setStyleSheet(self._grp_style())
        in_lay = QVBoxLayout(grp_in)
        in_lay.setSpacing(6)

        self._text_input = QTextEdit()
        self._text_input.setPlaceholderText(
            "E-Mail-Text des Passagiers hier einfügen …"
        )
        self._text_input.setFont(QFont("Segoe UI", 10))
        in_lay.addWidget(self._text_input)

        btn_row = QHBoxLayout()
        btn_posteingang = QPushButton("📬  Posteingang")
        btn_posteingang.setToolTip("Outlook-Posteingang öffnen und E-Mail auswählen")
        btn_posteingang.setStyleSheet(f"""
            QPushButton {{
                background-color: #5d6d7e;
                color: white; border: none; border-radius: 4px;
                padding: 6px 12px; font-size: 11px; font-weight: bold;
            }}
            QPushButton:hover  {{ background-color: #485460; color: white; }}
            QPushButton:pressed {{ background-color: #34495e; }}
        """)
        btn_posteingang.setMinimumHeight(34)
        btn_posteingang.clicked.connect(self._load_from_inbox)
        btn_row.addWidget(btn_posteingang)

        btn_analyse = QPushButton("🔍  Daten extrahieren")
        btn_analyse.setStyleSheet(self._btn_primary_style())
        btn_analyse.setMinimumHeight(34)
        btn_analyse.clicked.connect(self._extract)
        btn_row.addWidget(btn_analyse)
        in_lay.addLayout(btn_row)

        layout.addWidget(grp_in, stretch=1)

        # Extrahierte Felder
        grp_fields = QGroupBox("📋  Extrahierte Daten  (bearbeitbar)")
        grp_fields.setStyleSheet(self._grp_style())
        f_lay = QVBoxLayout(grp_fields)
        f_lay.setSpacing(5)

        self._f_name       = self._add_field(f_lay, "Name:")

        # Anrede-Combo
        anr_row = QHBoxLayout()
        anr_lbl = QLabel("Anrede:")
        anr_lbl.setFixedWidth(90)
        anr_lbl.setStyleSheet(f"color: {FIORI_TEXT}; font-weight: bold; font-size: 11px;")
        self._f_anrede = QComboBox()
        self._f_anrede.addItems(["–", "Herr", "Frau"])
        self._f_anrede.setMinimumHeight(26)
        self._f_anrede.setStyleSheet("""
            QComboBox { border: 1px solid #c8d0d8; border-radius: 4px;
                        padding: 3px 8px; font-size: 11px; }
            QComboBox:focus { border-color: #0070f3; }
        """)
        anr_row.addWidget(anr_lbl)
        anr_row.addWidget(self._f_anrede)
        f_lay.addLayout(anr_row)

        self._f_email      = self._add_field(f_lay, "E-Mail:")
        self._f_flugnummer = self._add_field(f_lay, "Flugnummer:")
        self._f_datum      = self._add_field(f_lay, "Datum:")
        self._f_rueckflug  = self._add_field(f_lay, "Rückflug:")

        layout.addWidget(grp_fields)
        return left

    def _add_field(self, parent_layout: QVBoxLayout, label: str) -> QLineEdit:
        row = QHBoxLayout()
        lbl = QLabel(label)
        lbl.setFixedWidth(90)
        lbl.setStyleSheet(f"color: {FIORI_TEXT}; font-weight: bold; font-size: 11px;")
        edit = QLineEdit()
        edit.setStyleSheet(self._input_style())
        edit.setMinimumHeight(26)
        row.addWidget(lbl)
        row.addWidget(edit)
        parent_layout.addLayout(row)
        return edit

    def _build_right(self) -> QWidget:
        right = QWidget()
        layout = QVBoxLayout(right)
        layout.setContentsMargins(8, 0, 0, 0)
        layout.setSpacing(6)

        # Szenario-Buttons
        grp_sz = QGroupBox("🎯  Szenario wählen")
        grp_sz.setStyleSheet(self._grp_style())
        sz_lay = QVBoxLayout(grp_sz)
        sz_lay.setSpacing(5)

        _BTN_COLOR = "#1e5799"   # einheitliches Dunkelblau – gut lesbar bei Hover
        _BTN_HOVER  = "#154060"

        scenarios = [
            ("✅  Szenario 1 – Alle Angaben vorhanden",      ANTWORT_KOMPLETT),
            ("⚠️  Szenario 2 – Fehlende Informationen",      ANTWORT_FEHLENDE_DATEN),
            ("🅿️  Szenario 3 – Abholung am Parkplatz",       ANTWORT_PARKPLATZ),
            ("ℹ️  Szenario 4 – Allgemeine PRM-Service-Info", ANTWORT_INFO_SERVICE),
        ]
        for label, text in scenarios:
            btn = QPushButton(label)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {_BTN_COLOR};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 7px 14px;
                    font-size: 11px;
                    font-weight: bold;
                    text-align: left;
                }}
                QPushButton:hover  {{ background-color: {_BTN_HOVER}; color: #ffffff; }}
                QPushButton:pressed {{ background-color: #0e2d42; }}
            """)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.setMinimumHeight(36)
            btn.clicked.connect(lambda _, t=text: self._set_antwort(t))
            sz_lay.addWidget(btn)

        btn_s5 = QPushButton("🚫  Szenario 5 – Ablehnung PRM durch Airline")
        btn_s5.setStyleSheet(f"""
            QPushButton {{
                background-color: #7d3c98;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 7px 14px;
                font-size: 11px;
                font-weight: bold;
                text-align: left;
            }}
            QPushButton:hover  {{ background-color: #6c3483; color: #ffffff; }}
            QPushButton:pressed {{ background-color: #5b2c6f; }}
        """)
        btn_s5.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        btn_s5.setMinimumHeight(36)
        btn_s5.clicked.connect(self._open_szenario5)
        sz_lay.addWidget(btn_s5)

        # Checkbox: Flugdaten anfordern
        self._chk_flugdaten = QCheckBox("  + Flugdaten anfordern (Hinweis in Antwort einfügen)")
        self._chk_flugdaten.setStyleSheet(f"color: {FIORI_TEXT}; font-size: 11px;")
        sz_lay.addWidget(self._chk_flugdaten)

        layout.addWidget(grp_sz)

        # Antwort-Bereich
        grp_ant = QGroupBox("📝  Antwort  (bearbeitbar)")
        grp_ant.setStyleSheet(self._grp_style())
        ant_lay = QVBoxLayout(grp_ant)
        ant_lay.setSpacing(6)

        self._text_antwort = QTextEdit()
        self._text_antwort.setPlaceholderText(
            "Szenario wählen oder Antwort hier manuell eingeben …"
        )
        self._text_antwort.setFont(QFont("Segoe UI", 10))
        ant_lay.addWidget(self._text_antwort)

        btn_outlook = QPushButton("📧  Outlook-Entwurf erstellen")
        btn_outlook.setStyleSheet(self._btn_primary_style())
        btn_outlook.setMinimumHeight(38)
        btn_outlook.clicked.connect(self._open_outlook)
        ant_lay.addWidget(btn_outlook)

        layout.addWidget(grp_ant, stretch=1)
        return right

    # ── Logik ──────────────────────────────────────────────────────────────────

    def _load_from_inbox(self):
        """Outlook-Posteingang öffnen, E-Mail auswählen und Felder befüllen."""
        dlg = OutlookInboxDialog(self)
        if dlg.exec() == OutlookInboxDialog.DialogCode.Accepted and dlg.selected_body:
            self._text_input.setPlainText(dlg.selected_body)
            # Zuerst Textextraktion laufen lassen
            self._extract()
            # Absender-E-Mail direkt aus Outlook überschreiben (Regex findet sie nicht im Body)
            if dlg.selected_sender_email:
                self._f_email.setText(dlg.selected_sender_email)
            # Absender-Name ergänzen, falls Extraktion leer war
            if dlg.selected_sender_name and not self._f_name.text().strip():
                self._f_name.setText(dlg.selected_sender_name)

    def _extract(self):
        text = self._text_input.toPlainText()
        fields = _parse_email_fields(text)
        self._f_name.setText(fields["name"])
        anrede = fields.get("anrede", "")
        self._f_anrede.setCurrentText(anrede if anrede in ("Herr", "Frau") else "–")
        self._f_email.setText(fields["email"])
        self._f_flugnummer.setText(fields["flugnummer"])
        self._f_datum.setText(fields["datum"])
        self._f_rueckflug.setText(fields["rueckflug"])

    def _set_antwort(self, template: str):
        text = template.strip()

        # 1. Anrede personalisieren
        name   = self._f_name.text().strip()
        anrede = self._f_anrede.currentText()   # "–", "Herr", "Frau"
        nachname = name.split()[-1] if name else ""

        if nachname and anrede in ("Herr", "Frau"):
            gen = "geehrter" if anrede == "Herr" else "geehrte"
            greeting = f"Sehr {gen} {anrede} {nachname},"
        elif nachname:
            greeting = f"Sehr geehrte/r {nachname},"
        else:
            greeting = "Sehr geehrte Damen und Herren,"

        text = text.replace("Sehr geehrte Damen und Herren,", greeting, 1)

        # 2. Bezug-Zeile mit Flugdaten einfügen (direkt nach Anrede + Leerzeile)
        flug  = self._f_flugnummer.text().strip()
        datum = self._f_datum.text().strip()
        ref_parts = []
        if flug:
            ref_parts.append(f"Flug {flug}")
        if datum:
            ref_parts.append(datum)
        if ref_parts:
            bezug = "Bezug: " + ", ".join(ref_parts) + "\n\n"
            # nach "greeting\n\n" einfügen
            marker = greeting + "\n\n"
            if marker in text:
                text = text.replace(marker, marker + bezug, 1)

        # 3. Flugdaten-Bitte vor der Signatur einfügen
        if self._chk_flugdaten.isChecked():
            sig_marker = "\n\nMit freundlichen Grüßen"
            if sig_marker in text:
                text = text.replace(sig_marker, _FLUGDATEN_BITTE + sig_marker, 1)
            else:
                text += _FLUGDATEN_BITTE

        self._text_antwort.setPlainText(text)

    def _open_szenario5(self):
        """Öffnet den Szenario-5-Dialog und sendet den Entwurf direkt an Outlook."""
        dlg = Szenario5Dialog(self)
        # Vorausfüllen aus bereits extrahierten Feldern
        if hasattr(self, '_f_email') and self._f_email.text().strip():
            dlg._e_to.setText(self._f_email.text().strip())
            # Synchron auch erste Kontaktzeile befüllen
            if dlg._kontakt_rows:
                row0 = dlg._kontakt_rows[0]
                row0._email.setText(self._f_email.text().strip())
                anrede = self._f_anrede.currentText()
                if anrede in ("Herr", "Frau"):
                    row0._anrede.setCurrentText(anrede)
                name_parts = self._f_name.text().strip().split()
                if name_parts:
                    row0._vorname.setText(name_parts[0])
        if hasattr(self, '_f_flugnummer'):
            dlg._e_flug.setText(self._f_flugnummer.text().strip())
        if hasattr(self, '_f_datum'):
            dlg._e_datum.setText(self._f_datum.text().strip())

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        from functions.mail_functions import create_outlook_draft
        import os
        logo = _LOGO_PATH if os.path.isfile(_LOGO_PATH) else None
        try:
            create_outlook_draft(
                to=dlg.result_to,
                cc=dlg.result_cc,
                subject=dlg.result_subject,
                body_text=dlg.result_body,
                logo_path=logo,
            )
            # Text auch in Antwort-Feld übernehmen
            self._text_antwort.setPlainText(dlg.result_body)
            self._f_email.setText(dlg.result_to.split(";")[0].strip())
            QMessageBox.information(
                self,
                "Outlook",
                "Outlook-Entwurf (Szenario 5) wurde geöffnet.\nBitte prüfen und absenden.",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Outlook-Fehler",
                f"Entwurf konnte nicht erstellt werden:\n{exc}\n\n"
                "Bitte sicherstellen, dass Outlook geöffnet und pywin32 installiert ist.",
            )

    def _open_outlook(self):
        """Erstellt Outlook-Entwurf via win32com mit DRK-Logo als Signatur."""
        from functions.mail_functions import create_outlook_draft

        recipient = self._f_email.text().strip()
        body = self._text_antwort.toPlainText().strip()

        subject_parts = ["PRM-Service – Flughafen Köln/Bonn"]
        if self._f_name.text().strip():
            subject_parts.append(self._f_name.text().strip())
        if self._f_flugnummer.text().strip():
            subject_parts.append(f"Flug {self._f_flugnummer.text().strip()}")
        if self._f_datum.text().strip():
            subject_parts.append(self._f_datum.text().strip())
        subject = " | ".join(subject_parts)

        logo = _LOGO_PATH if os.path.isfile(_LOGO_PATH) else None

        try:
            create_outlook_draft(
                to=recipient,
                subject=subject,
                body_text=body,
                logo_path=logo,
            )
            QMessageBox.information(
                self,
                "Outlook",
                "Outlook-Entwurf wurde geöffnet.\nBitte prüfen und absenden.",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Outlook-Fehler",
                f"Entwurf konnte nicht erstellt werden:\n{exc}\n\n"
                "Bitte sicherstellen, dass Outlook geöffnet und pywin32 installiert ist.",
            )

    def refresh(self):
        pass

    # ── Styles ─────────────────────────────────────────────────────────────────

    def _grp_style(self) -> str:
        return f"""
            QGroupBox {{
                border: 1px solid #c8d0d8;
                border-radius: 6px;
                margin-top: 14px;
                padding-top: 6px;
                font-size: 11px;
                font-weight: bold;
                color: {FIORI_TEXT};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                padding: 0 6px;
                color: {FIORI_BLUE};
            }}
        """

    def _btn_primary_style(self) -> str:
        return f"""
            QPushButton {{
                background-color: {FIORI_BLUE};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 7px 18px;
                font-size: 12px;
                font-weight: bold;
            }}
            QPushButton:hover  {{ background-color: #1a5276; color: white; }}
            QPushButton:pressed {{ background-color: #154360; }}
        """

    def _input_style(self) -> str:
        return f"""
            QLineEdit {{
                border: 1px solid #c8d0d8;
                border-radius: 4px;
                padding: 3px 8px;
                font-size: 11px;
                color: {FIORI_TEXT};
            }}
            QLineEdit:focus {{ border-color: {FIORI_BLUE}; }}
        """
