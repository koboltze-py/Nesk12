# Nesk3 – E-Mail-Funktion: Handys

## Übersicht

Der Button **„📧 Per E-Mail senden"** im Handys-Widget löst folgende Schritte aus:

1. Excel-Export ausführen (Datei wird lokal gespeichert, s. `handys_excel_export.md`)
2. Standard-E-Mail-Client öffnen (Outlook / MAPI) **oder** direkter SMTP-Versand
3. E-Mail vorausgefüllt mit Empfänger, Betreff, Text und Anhang

---

## E-Mail-Parameter

### Empfänger

```
erste-hilfe-station-flughafen@drk-koeln.de
```

### Betreff

```
Diensthandy-Übersicht – DRK EHS CGN – {DATUM}
```

Beispiel: `Diensthandy-Übersicht – DRK EHS CGN – 15.07.2025`

### E-Mail-Text (Vorlage)

```
Sehr geehrter Herr Peters,

anbei erhalten Sie die aktuelle Übersicht aller Diensthandys der
DRK Erste-Hilfe-Station am Flughafen Köln/Bonn.

Die beigefügte Excel-Datei enthält:
  • Eine vollständige Geräteliste mit aktuellem Status (Aktiv/Defekt/
    Außer Betrieb/Reserve)
  • Eine Übersicht der letzten Zustandsänderungen (90-Tage-Verlauf)

Bei Rückfragen stehen wir gerne zur Verfügung.

Mit freundlichen Grüßen

[ABSENDER_NAME]
Schichtleiter | DRK Kreisverband Köln e.V.
Erste-Hilfe-Station Flughafen Köln/Bonn

──────────────────────────────────────────
Deutsches Rotes Kreuz
Kreisverband Köln e.V.
Erste-Hilfe-Station Flughafen Köln/Bonn
Terminal 2 | 51147 Köln
Tel.: +49 (0)2203 / 40 40 - 0
E-Mail: erste-hilfe-station-flughafen@drk-koeln.de
──────────────────────────────────────────
```

`[ABSENDER_NAME]` wird aus `config.py` oder dem angemeldeten Benutzer befüllt.

### Anhang

Erzeugte Excel-Datei:
```
Handys_Uebersicht_YYYY-MM-DD.xlsx
```

---

## Implementierung (`handys_email.py`)

### Variante A – MAPI / Outlook (empfohlen für DRK-Umgebung)

Öffnet Outlook mit vorausgefüllter E-Mail; Benutzer sendet manuell ab.
Kein separates SMTP-Passwort nötig.

```python
import os
import win32com.client   # pip install pywin32
from datetime import datetime
from config import HANDYS_EXPORT_PATH
from modules.handys.handys_excel_export import export_handys_excel

EMPFAENGER = "erste-hilfe-station-flughafen@drk-koeln.de"


def sende_handys_email(absender_name: str = "DRK Schichtleiter") -> str:
    """
    Exportiert Excel, öffnet Outlook-E-Mail mit Anhang.
    Gibt den Pfad der gespeicherten Excel-Datei zurück.
    """
    # Schritt 1: Excel erzeugen und speichern
    excel_pfad = export_handys_excel(open_after=False)

    # Schritt 2: Outlook-Mail aufbauen
    datum_anzeige = datetime.now().strftime('%d.%m.%Y')

    betreff = f"Diensthandy-Übersicht – DRK EHS CGN – {datum_anzeige}"

    text = f"""Sehr geehrter Herr Peters,

anbei erhalten Sie die aktuelle Übersicht aller Diensthandys der
DRK Erste-Hilfe-Station am Flughafen Köln/Bonn.

Die beigefügte Excel-Datei enthält:
  \u2022 Eine vollständige Geräteliste mit aktuellem Status (Aktiv/Defekt/
    Außer Betrieb/Reserve)
  \u2022 Eine Übersicht der letzten Zustandsänderungen (90-Tage-Verlauf)

Bei Rückfragen stehen wir gerne zur Verfügung.

Mit freundlichen Grüßen

{absender_name}
Schichtleiter | DRK Kreisverband Köln e.V.
Erste-Hilfe-Station Flughafen Köln/Bonn

\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
Deutsches Rotes Kreuz
Kreisverband Köln e.V.
Erste-Hilfe-Station Flughafen Köln/Bonn
Terminal 2 | 51147 Köln
Tel.: +49 (0)2203 / 40 40 - 0
E-Mail: erste-hilfe-station-flughafen@drk-koeln.de
\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500"""

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = EMPFAENGER
    mail.Subject = betreff
    mail.Body = text
    mail.Attachments.Add(excel_pfad)
    mail.Display(True)   # True = modaler Dialog; False = im Hintergrund öffnen

    return excel_pfad
```

> **Hinweis:** `mail.Display(True)` öffnet das Outlook-Fenster. Der Benutzer kann den Text
> noch anpassen und sendet manuell ab. Für automatischen Versand `mail.Send()` verwenden –
> dann ist aber eine Outlook-Sicherheitsabfrage möglich.

---

### Variante B – Fallback: `mailto`-Link (kein pywin32)

Falls `pywin32` nicht verfügbar (z. B. andere Outlook-Version):

```python
import urllib.parse
import webbrowser

def sende_handys_email_mailto(absender_name: str = "DRK Schichtleiter") -> str:
    excel_pfad = export_handys_excel(open_after=False)
    datum_anzeige = datetime.now().strftime('%d.%m.%Y')

    betreff = f"Diensthandy-Übersicht – DRK EHS CGN – {datum_anzeige}"
    # mailto kann keinen Anhang setzen – Benutzer muss manuell anfügen
    hinweis = f"\n\n>>> BITTE DATEI MANUELL ANFÜGEN:\n{excel_pfad}\n\n"

    params = urllib.parse.urlencode({
        'to': EMPFAENGER,
        'subject': betreff,
        'body': hinweis + EMAIL_TEXT_TEMPLATE.format(absender_name=absender_name)
    }, quote_via=urllib.parse.quote)

    webbrowser.open(f"mailto:?{params}")
    return excel_pfad
```

> Bei `mailto` ist kein automatischer Anhang möglich. Der Benutzer erhält einen
> Dialog-Hinweis mit dem Dateipfad und muss die Datei selbst einhängen.

---

## Aufruf aus dem Widget

```python
# Button-Handler im HandysWidget
def on_email_senden_clicked(self):
    from config import ABSENDER_NAME  # z. B. aus Einstellungen
    try:
        pfad = sende_handys_email(absender_name=ABSENDER_NAME)
        QMessageBox.information(
            self,
            "E-Mail geöffnet",
            f"Outlook wurde geöffnet.\n\nExcel-Datei gespeichert unter:\n{pfad}"
        )
    except Exception as e:
        QMessageBox.critical(self, "Fehler beim E-Mail-Versand", str(e))
```

---

## Ablauf im Überblick

```
Benutzer klickt [📧 Per E-Mail senden]
        │
        ▼
export_handys_excel()
   → Datei schreiben nach HANDYS_EXPORT_PATH
   → Pfad zurückgeben
        │
        ▼
sende_handys_email()
   → Outlook öffnen
   → Empfänger, Betreff, Text, Anhang setzen
   → Fenster anzeigen (Benutzer sendet ab)
        │
        ▼
QMessageBox: „E-Mail geöffnet, Datei gespeichert unter: ..."
```

---

## Konfiguration in `config.py`

```python
# E-Mail
HANDYS_EMAIL_EMPFAENGER  = "erste-hilfe-station-flughafen@drk-koeln.de"
HANDYS_EMAIL_ABSENDER    = "DRK Schichtleiter"   # Kann auch aus Benutzereinstellungen kommen

# Exportpfad
HANDYS_EXPORT_PATH = (
    r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - "
    r"Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - "
    r"DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3\Daten\Handys"
)
```

---

## Abhängigkeiten

| Paket | Zweck | Installation |
|-------|-------|-------------|
| `pywin32` | Outlook MAPI-Zugriff | `pip install pywin32` |
| `openpyxl` | Excel-Erstellung | bereits in Nesk3 vorhanden |

`pywin32` ist nur auf Windows verfügbar – passt zur Nesk3-Umgebung.
