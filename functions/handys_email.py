"""
Handys – E-Mail-Funktion
Öffnet Outlook mit vorausgefüllter E-Mail und Excel-Anhang.
"""
import os
import sys
import urllib.parse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import HANDYS_EMAIL_EMPFAENGER
from functions.handys_excel_export import export_handys_excel


def _absender_name() -> str:
    try:
        return os.getlogin()
    except Exception:
        return "DRK Schichtleiter"


def _outlook_mail_oeffnen(
    an: str,
    betreff: str,
    text: str,
    anhang_pfad: str,
) -> None:
    """Öffnet eine neue Outlook-Mail mit Empfänger, Betreff, Text und Anhang."""
    import win32com.client  # type: ignore
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)  # olMailItem
    mail.Display()               # Outlook lädt Standardsignatur in HTMLBody
    signatur = mail.HTMLBody     # Signatur sichern

    mail.To      = an
    mail.Subject = betreff

    # Text als HTML + gespeicherte Signatur dahinter
    zeilen_html = "".join(
        f"<p style='margin:0'>{z if z.strip() else '&nbsp;'}</p>"
        for z in text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").split("\n")
    )
    body_html = (
        "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"
        + zeilen_html
        + "</body></html>"
    )
    mail.HTMLBody = body_html + (signatur or "")

    # Anhang zuletzt hinzufügen (nach dem Body-Setzen)
    mail.Attachments.Add(anhang_pfad)

def _email_text(absender: str, datum: str) -> str:
    linie = "─" * 42
    return (
        f"Sehr geehrter Herr Peters,\n\n"
        f"anbei erhalten Sie die aktuelle Übersicht aller Diensthandys der\n"
        f"DRK Erste-Hilfe-Station am Flughafen Köln/Bonn.\n\n"
        f"Die beigefügte Excel-Datei enthält:\n"
        f"  \u2022 Eine vollständige Geräteliste mit aktuellem Status "
        f"(Aktiv/Defekt/Außer Betrieb/Reserve)\n"
        f"  \u2022 Eine Übersicht der letzten Zustandsänderungen (90-Tage-Verlauf)\n\n"
        f"Bei Rückfragen stehen wir gerne zur Verfügung.\n\n"
        f"Mit freundlichen Grüßen\n\n"
        f"{absender}\n"
        f"Schichtleiter | DRK Kreisverband Köln e.V.\n"
        f"Erste-Hilfe-Station Flughafen Köln/Bonn\n\n"
        f"{linie}\n"
        f"Deutsches Rotes Kreuz\n"
        f"Kreisverband Köln e.V.\n"
        f"Erste-Hilfe-Station Flughafen Köln/Bonn\n"
        f"Terminal 2 | 51147 Köln\n"
        f"Tel.: +49 (0)2203 / 40 40 - 0\n"
        f"E-Mail: erste-hilfe-station-flughafen@drk-koeln.de\n"
        f"{linie}"
    )


def sende_handys_email(absender_name: str | None = None) -> tuple[bool, str]:
    """
    Exportiert die Handy-Übersicht als Excel und öffnet Outlook mit vorausgefüllter E-Mail.

    Rückgabe: (erfolg: bool, meldung: str)
    """
    if absender_name is None:
        absender_name = _absender_name()

    datum_anzeige = datetime.now().strftime("%d.%m.%Y")
    betreff = f"Diensthandy-Übersicht – DRK EHS CGN – {datum_anzeige}"

    # Schritt 1: Excel erzeugen (Fehler wird immer reportiert, unabhängig von E-Mail)
    try:
        excel_pfad = export_handys_excel(open_after=False)
    except Exception as e:
        return False, f"Excel-Export fehlgeschlagen: {e}"

    text = _email_text(absender_name, datum_anzeige)

    # Schritt 2: Outlook via pywin32 (Variante A – bevorzugt)
    try:
        _outlook_mail_oeffnen(HANDYS_EMAIL_EMPFAENGER, betreff, text, excel_pfad)
        return True, f"Outlook geöffnet. Anhang: {os.path.basename(excel_pfad)}"
    except ImportError:
        pass
    except Exception as e:
        return False, f"Outlook-Fehler: {e}\n\nExcel gespeichert unter:\n{excel_pfad}"

    # Schritt 3: Fallback – mailto-Link (kein pywin32 / kein Outlook)
    try:
        mailto = (
            f"mailto:{urllib.parse.quote(HANDYS_EMAIL_EMPFAENGER)}"
            f"?subject={urllib.parse.quote(betreff)}"
            f"&body={urllib.parse.quote(text)}"
        )
        import subprocess
        subprocess.Popen(["cmd", "/c", "start", "", mailto], shell=False)
        meldung = (
            f"Standard-E-Mail-Client geöffnet (kein Outlook/pywin32 verfügbar).\n"
            f"Bitte Anhang manuell hinzufügen:\n{excel_pfad}"
        )
        return True, meldung
    except Exception as e:
        return False, (
            f"E-Mail-Client konnte nicht geöffnet werden: {e}\n\n"
            f"Excel gespeichert unter:\n{excel_pfad}\n\n"
            f"Empfänger: {HANDYS_EMAIL_EMPFAENGER}\n"
            f"Betreff:   {betreff}"
        )


def _bericht_email_text(absender: str, inventarnummer: str, bericht_typ: str, datum: str) -> str:
    linie = "─" * 42
    return (
        f"Sehr geehrter Herr Peters,\n\n"
        f"anbei erhalten Sie den {bericht_typ} für das Gerät {inventarnummer}\n"
        f"der DRK Erste-Hilfe-Station am Flughafen Köln/Bonn.\n\n"
        f"Die beigefügte Word-Datei enthält:\n"
        f"  \u2022 Gerätedaten und Zustandsbeschreibung\n"
        f"  \u2022 Eingeleitete Maßnahmen\n\n"
        f"Bei Rückfragen stehen wir gerne zur Verfügung.\n\n"
        f"Mit freundlichen Grüßen\n\n"
        f"{absender}\n"
        f"Schichtleiter | DRK Kreisverband Köln e.V.\n"
        f"Erste-Hilfe-Station Flughafen Köln/Bonn\n\n"
        f"{linie}\n"
        f"Deutsches Rotes Kreuz\n"
        f"Kreisverband Köln e.V.\n"
        f"Erste-Hilfe-Station Flughafen Köln/Bonn\n"
        f"Terminal 2 | 51147 Köln\n"
        f"Tel.: +49 (0)2203 / 40 40 - 0\n"
        f"E-Mail: erste-hilfe-station-flughafen@drk-koeln.de\n"
        f"{linie}"
    )


def sende_bericht_email(
    pfad: str,
    inventarnummer: str,
    bericht_typ: str,
    absender_name: str | None = None,
) -> tuple[bool, str]:
    """
    Öffnet Outlook mit vorausgefüllter E-Mail und dem angegebenen Word-Bericht als Anhang.

    Rückgabe: (erfolg: bool, meldung: str)
    """
    if absender_name is None:
        absender_name = _absender_name()

    datum_anzeige = datetime.now().strftime("%d.%m.%Y")
    betreff = f"{bericht_typ} {inventarnummer} – DRK EHS CGN – {datum_anzeige}"
    text = _bericht_email_text(absender_name, inventarnummer, bericht_typ, datum_anzeige)
    dateiname = os.path.basename(pfad)

    # Schritt 1: Outlook via pywin32
    try:
        _outlook_mail_oeffnen(HANDYS_EMAIL_EMPFAENGER, betreff, text, pfad)
        return True, f"Outlook geöffnet. Anhang: {dateiname}"
    except ImportError:
        pass
    except Exception as e:
        return False, f"Outlook-Fehler: {e}\n\nBericht: {pfad}"

    # Schritt 2: Fallback – mailto (ohne Anhang, da mailto keinen Anhang unterstützt)
    try:
        mailto = (
            f"mailto:{urllib.parse.quote(HANDYS_EMAIL_EMPFAENGER)}"
            f"?subject={urllib.parse.quote(betreff)}"
            f"&body={urllib.parse.quote(text)}"
        )
        import subprocess
        subprocess.Popen(["cmd", "/c", "start", "", mailto], shell=False)
        meldung = (
            f"Standard-E-Mail-Client geöffnet (kein Outlook/pywin32 verfügbar).\n"
            f"Bitte Anhang manuell hinzufügen:\n{pfad}"
        )
        return True, meldung
    except Exception as e:
        return False, (
            f"E-Mail-Client konnte nicht geöffnet werden: {e}\n\n"
            f"Bericht gespeichert unter:\n{pfad}\n\n"
            f"Empfänger: {HANDYS_EMAIL_EMPFAENGER}\n"
            f"Betreff:   {betreff}"
        )
