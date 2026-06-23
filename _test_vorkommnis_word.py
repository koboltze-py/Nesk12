# Testskript: Vorkommnis Word-Export
# Erstellt mehrere Word-Dateien mit verschiedenen Abschnitts-Konstellationen
# Ausgabe: Desktop/bei/vorkomtest/
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pathlib import Path
from gui.vorkommnisse import VorkommnisseWidget

AUSGABE_DIR = Path(
    "C:/Users/DRKairport/OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V"
    "/Desktop/bei/vorkomtest"
)
AUSGABE_DIR.mkdir(parents=True, exist_ok=True)

# ─── Minimaler PySide6-Kontext (keine echtes Fenster, nur Widget-Instanz) ────
from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication(sys.argv)
w = VorkommnisseWidget()

def _export(dateiname: str, daten: dict):
    pfad = str(AUSGABE_DIR / dateiname)
    try:
        w._erstelle_word(pfad, daten)
        print(f"  OK  {dateiname}")
    except Exception as exc:
        print(f"  FEHLER  {dateiname}: {exc}")
        raise

print("=== Vorkommnis Word-Export Tests ===\n")

# ── Test 1: Einfaches PRM-Vorkommnis, Standardtitel ──────────────────────────
_export("Test1_PRM_Standard.docx", {
    "flug": "XQ983",
    "typ": "PRM-Betreuung",
    "bereich": "Flüge",
    "datum": "23.06.2026",
    "ort": "Köln/Bonn (CGN)",
    "offblock_plan": "08:30 Uhr",
    "offblock_ist": "09:05 Uhr (+35 Min.)",
    "erstellt_von": "Max Mustermann",
    "passagiere": [
        ["Müller, Hans", "PRM Passagier", "WCHS", "Rollstuhl vorgezogen"],
        ["Schmidt, Anna", "Passagier", "", "Begleitperson"],
    ],
    "personal": [
        ["Kamencic, Elvedin", "PRM-Begleitung", ""],
        ["Groß, Stefan",      "Fahrer",         "Bulmor C58"],
    ],
    "chronologie": [
        ["06:45", "Benachrichtigung über PRM-Anforderung"],
        ["07:10", "Fahrzeug C58 an Terminal 2 Abflug positioniert"],
        ["07:35", "Passagier am Check-in abgeholt"],
        ["08:00", "Passagier am Gate D82 übergeben"],
        ["09:05", "Offblock Flug XQ983 – 35 Min. Verspätung"],
    ],
    "ursache": "• Verspätete Benachrichtigung durch Airline\n• Gate-Änderung kurzfristig (D60 → D82)",
    "ergebnis": "Passagier wurde erfolgreich befördert. Verspätung dokumentiert und an Disposition weitergegeben.",
    "abschnitt_titel": {
        "1": "Betroffene Personen",
        "2": "Eingeteiltes Personal",
        "3": "Chronologischer Ablauf",
        "4": "Ursachenanalyse",
        "5": "Ergebnis",
    },
    "extra_abschnitte": [],
})

# ── Test 2: Fahrzeugschaden, geänderte Abschnittstitel ────────────────────────
_export("Test2_Fahrzeugschaden_CustomTitel.docx", {
    "flug": "C58",
    "typ": "Fahrzeugschaden",
    "bereich": "Fahrzeuge",
    "datum": "21.06.2026",
    "ort": "Vorfeld Süd, Gate 10",
    "offblock_plan": "",
    "offblock_ist": "",
    "erstellt_von": "Stefan Groß",
    "passagiere": [],
    "personal": [
        ["Groß, Stefan", "Fahrer", ""],
    ],
    "chronologie": [
        ["10:00", "Schaden am linken Hinterrad festgestellt"],
        ["10:15", "Fahrzeug außer Betrieb genommen"],
        ["11:30", "Werkstatt informiert"],
    ],
    "ursache": "Reifenschaden durch Fremdkörper auf dem Vorfeld.",
    "ergebnis": "Fahrzeug wurde abgeschleppt. Einsatz durch Ersatzfahrzeug C42 sichergestellt.",
    "abschnitt_titel": {
        "1": "Beteiligte Fahrzeuge",
        "2": "Verantwortliches Personal",
        "3": "Schadenshergang",
        "4": "Schadensursache",
        "5": "Maßnahmen & Ergebnis",
    },
    "extra_abschnitte": [],
})

# ── Test 3: Medizinischer Notfall mit 2 Zusatzabschnitten ─────────────────────
_export("Test3_MedNotfall_Zusatzabschnitte.docx", {
    "flug": "PC5066",
    "typ": "Medizinischer Notfall",
    "bereich": "Flüge",
    "datum": "23.06.2026",
    "ort": "Gate D82",
    "offblock_plan": "",
    "offblock_ist": "",
    "erstellt_von": "Elvedin Kamencic",
    "passagiere": [
        ["Schneider, Karl", "Patient", "", "Kreislaufkollaps"],
    ],
    "personal": [
        ["Kamencic, Elvedin", "PRM-Begleitung", ""],
        ["Mustermann, Lisa",  "Servicepoint",   "Notarzt alarmiert"],
    ],
    "chronologie": [
        ["11:49", "Patient am Gate D82 bewusstlos aufgefunden"],
        ["11:51", "Notarzt alarmiert, Erste Hilfe eingeleitet"],
        ["12:10", "Notarzt eingetroffen"],
        ["12:25", "Patient ins Krankenhaus Köln-Merheim eingeliefert"],
    ],
    "ursache": "Kreislaufkollaps, vermutlich durch Hitzebelastung und längere Wartezeit.",
    "ergebnis": "Patient stabil in Klinik übergeben. Flug PC5066 mit 45 Min. Verspätung abgeflogen.",
    "abschnitt_titel": {
        "1": "Betroffene Personen",
        "2": "Eingesetztes Personal",
        "3": "Chronologischer Ablauf",
        "4": "Ursache",
        "5": "Ergebnis & Folgeaktionen",
    },
    "extra_abschnitte": [
        {
            "titel": "Rückmeldung Klinik",
            "text": "Lt. Rückmeldung Krankenhaus Köln-Merheim (13:30 Uhr):\n"
                    "Patient wurde stationär aufgenommen. Zustand stabil.",
        },
        {
            "titel": "Meldung an Flughafen-Sicherheitsdienst",
            "text": "Meldung an FHF-Sicherheit erstattet (Ref. 2026-0623-04).\n"
                    "Kein weiterer Handlungsbedarf lt. Sicherheitsdienst.",
        },
    ],
})

# ── Test 4: Sicherheitsvorfall, viele Personen, langer Text ──────────────────
_export("Test4_Sicherheitsvorfall.docx", {
    "flug": "W4-3918",
    "typ": "Sicherheitsvorfall",
    "bereich": "Flüge",
    "datum": "06.06.2026",
    "ort": "Check-in Bereich C",
    "offblock_plan": "14:00 Uhr",
    "offblock_ist": "15:20 Uhr (+80 Min.)",
    "erstellt_von": "Leitungsdienst",
    "passagiere": [
        ["Yilmaz, Mehmet",   "Passagier",     "", "Verweigert Boarding"],
        ["Yilmaz, Fatma",    "Passagier",     "", "Begleitperson"],
        ["Unbekannt, Kind",  "Passagier",     "", "UMNR unter Obhut"],
    ],
    "personal": [
        ["Kamencic, Elvedin", "Leitung",       "Einsatzleiter vor Ort"],
        ["Groß, Stefan",      "Fahrer",        "Transfer-Bereitschaft"],
        ["Bie, Anna",         "Servicepoint",  "Koordination Check-in"],
    ],
    "chronologie": [
        ["13:45", "Meldung durch Airline: Passagier verweigert Sicherheitscheck"],
        ["13:50", "DRK-Leitung informiert und zum Gate C entsandt"],
        ["14:05", "Bundespolizei angefordert"],
        ["14:30", "Bundespolizei eingetroffen, Lage beruhigt"],
        ["15:00", "Passagier freiwillig zum Check weitergeleitet"],
        ["15:20", "Boarding abgeschlossen, Offblock"],
    ],
    "ursache": (
        "• Passagier verweigerte erweiterten Sicherheitscheck aufgrund sprachlicher Missverständnisse\n"
        "• Kommunikationsbarriere (kein Deutsch/Englisch)\n"
        "• Eskalation durch fehlendes Dolmetscherangebot seitens der Airline"
    ),
    "ergebnis": (
        "Vorfall wurde durch Bundespolizei und DRK-Leitung deeskaliert.\n"
        "Passagier und Begleitperson haben Flug W4-3918 mit 80 Min. Verspätung angetreten.\n"
        "Vorfallbericht an Flughafenkoordination und Airline weitergeleitet."
    ),
    "abschnitt_titel": {
        "1": "Beteiligte Personen",
        "2": "Eingesetztes DRK-Personal",
        "3": "Ereignisablauf",
        "4": "Ursachenanalyse",
        "5": "Ergebnis & Folgemaßnahmen",
    },
    "extra_abschnitte": [
        {
            "titel": "Empfehlungen",
            "text": (
                "• Bereitstellung von Dolmetschdiensten an Check-in-Schaltern prüfen\n"
                "• Schulung DRK-Personal: Deeskalation bei sprachlichen Barrieren\n"
                "• Abstimmung mit Bundespolizei für schnellere Reaktionszeit"
            ),
        },
    ],
})

print(f"\nAlle Dateien gespeichert in:\n  {AUSGABE_DIR}")
import subprocess
subprocess.Popen(f'explorer "{AUSGABE_DIR}"')
