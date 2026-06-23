# GUI-Module – Dokumentation

## Übersicht Sidebar-Navigation

Die Sidebar ist links im Hauptfenster. Navigation via `QStackedWidget`. Seiten werden per Index adressiert.

| Index | Icon | Seitenname | Widget-Klasse |
|---|---|---|---|
| 0 | 🏠 | Dashboard | `DashboardWidget` |
| 1 | 📋 | Aufgaben | `AufgabenHauptWidget` |
| 2 | ⭐ | Sonderaufgaben | `SonderaufgabenWidget` |
| 3 | 🚑 | Fahrzeuge | `FahrzeugeWidget` |
| 4 | 👥 | Mitarbeiter | `MitarbeiterWidget` |
| 5 | 📅 | Dienstplan | `DienstplanWidget` |
| 6 | 🔄 | Übergabe | `UebergabeWidget` |
| 7 | ⚠️ | Vorkommnisse | `VorkommnisseWidget` |
| 8 | ✅ | Checklisten | `ChecklistenWidget` |
| 9 | 🎓 | Schulungen | `SchulungenKalenderWidget` |
| 10 | 📊 | Bericht | `BerichtWidget` |
| 11 | 🔀 | Workflow | `WorkflowWidget` |
| 12 | 📞 | Telefon | `TelefonnummernWidget` |
| 13 | 📁 | Dienstliches | `DienstlichesWidget` |
| 14 | 🩺 | SANMAT | _(sanmat/)_ |
| 15 | ✈️ | Passagiere | `PassagiereWidget` |
| 16 | ✉️ | Passagieranfragen | `PassagieranfragenWidget` |
| 17 | 📱 | Handys | `HandysWidget` |
| 18 | 📞 | Transkription | `CallTranscriptionWidget` |
| 19 | ⚙️ | Einstellungen | `EinstellungenWidget` |
| 20 | 💾 | Backup | `BackupWidget` |
| 21 | ❓ | Hilfe | `HilfeDialog` |

---

## `gui/main_window.py` – Hauptfenster

### Klasse: `MainWindow`

**Aufgaben:**
- Erstellt Sidebar + QStackedWidget
- Verwaltet Navigation (Sidebar-Buttons → Seitenwechsel)
- Fade-Animation bei Seitenwechsel (180ms, OutCubic)
- Zeigt animiertes `_NeskLogoWidget` oben in Sidebar
- Hintergrund-Thread für Turso-Sync

### Klasse: `_NeskLogoWidget`

- Animiertes NeSk-Logo mit Doppelring (Teal vorwärts, Gold rückwärts)
- QTimer 30ms (~33 FPS)
- Shimmer-Effekt auf „NeSk"-Schriftzug
- Doppelklick öffnet Slot-Machine (`gui/slot_machine.py`)

### Fade-Animation (Seitenwechsel)

```python
effect = QGraphicsOpacityEffect(widget)
widget.setGraphicsEffect(effect)
anim = QPropertyAnimation(effect, b"opacity")
anim.setDuration(180)
anim.setStartValue(0.0)
anim.setEndValue(1.0)
anim.setEasingCurve(QEasingCurve.Type.OutCubic)
anim.finished.connect(lambda: widget.setGraphicsEffect(None))
anim.start()
```

---

## `gui/dashboard.py` – Dashboard

### Klassen

| Klasse | Beschreibung |
|---|---|
| `DashboardWidget` | Haupt-Widget mit Statistiken, Kalender, Fahrzeugterminen |
| `StatCard` | Einzelne Statistik-Karte (Titel, Wert, Icon, Farbe) |
| `_TerminKalender` | QCalendarWidget mit farbigen Punkten für Termine/Notizen |
| `_SkyWidget` | Animierter Himmel mit Flugzeug-Animation (QPainter + QTimer) |

### StatCard-Verwendung
```python
card = StatCard(
    title="Mitarbeiter",
    value="47",
    icon="👥",
    color=FIORI_BLUE
)
card.set_value("48")  # Wert dynamisch aktualisieren
```

### Kalender-Punkte
- **Blauer Punkt**: Fahrzeug-Termin an diesem Tag
- **Grüner Punkt**: Notiz an diesem Tag
- Beide gleichzeitig: Blauer Punkt links, grüner Punkt rechts

---

## `gui/aufgaben.py` + `aufgaben_tag.py` + `aufgaben_haupt.py` – Aufgaben

### Aufgaben-Container (`aufgaben_haupt.py`)
- `AufgabenHauptWidget` enthält QTabWidget mit:
  - Tab 0: Tagdienst-Aufgaben (`aufgaben_tag.py`)
  - Tab 1: Nachtdienst-Aufgaben (`aufgaben.py`)

### Nachtdienst (`aufgaben.py`)
- Erfassung aller Aufgaben für die Nachtschicht
- Felder: Schichtleiter, Dispo, Betreuer, Service-Point, Bemerkungen
- Excel-Export nach `Daten/Aufgaben/`

### Tagdienst (`aufgaben_tag.py`)
- Wie Nacht, aber für Tagschicht (09:00–21:00)
- Zusätzlich: Frühschicht, Spätschicht

---

## `gui/sonderaufgaben.py` – Sonderaufgaben

### Funktionen

| Methode | Beschreibung |
|---|---|
| `_save_to_excel()` | Speichert Formular als Excel (`Daten/Sonderaufgaben/`) |
| `_print_dialog()` | Druckdialog: Anzahl pro Dokument wählbar |
| `_restore_last()` | Dropdown-Auswahl gespeicherter Dateien → Formular befüllen |
| `_load_from_excel(path)` | Liest alle Felder (Aufgaben, Service-Point, Bemerkung) aus Excel |
| `_open_archiv_ordner()` | Öffnet `Backup Data/Dokumente/Sonderaufgaben` im Explorer |

### Abschnitt „👷 Vorfeldmitarbeiter"
- 3 Gruppen (09:00–14:00, 14:00–19:00, 19:00–00:00)
- Je 3 Mitarbeiter-Slots als Dropdowns
- Befüllt aus aktuellem Dienstplan (Tag + Nacht dedupliziert)
- Suffix: Schichttyp (T/T10/N/N10), Bulmorfahrer (B), E-Mobby-Fahrer (EM)
- Excel-Export der Vorfeldmitarbeiterliste (Querformat A4, s/w)

### Druckdialog (PowerShell COM-Automation)
```python
# Drucken via Excel.Application COM:
# - Wartet auf Abschluss vor nächster Datei
# - Spinner: individuelle Kopienanzahl pro Dokument
# Standard: Sonderaufgaben = 2×, Vorfeldmitarbeiterliste = 3×
```

---

## `gui/fahrzeuge.py` – Fahrzeugverwaltung

**Funktionen:**
- Fahrzeugliste mit aktuellem Status
- Statushistorie (fahrbereit/defekt/werkstatt/ausser_dienst)
- Schadensmeldungen erfassen und per E-Mail senden
- Wartungstermine (TÜV, Service, HU)
- Kalenderansicht mit Terminen
- Schadensberichte als Word-Dokument exportieren (`gui/schadensbericht_dialog.py`)

**Status-Werte:**
- `fahrbereit` (Grün)
- `defekt` (Rot)
- `werkstatt` (Orange)
- `ausser_dienst` (Grau)
- `sonstiges` (Blau)

---

## `gui/mitarbeiter.py` – Mitarbeiterverwaltung

**Tabs:**
- 🗂️ Verwaltung: CRUD-Formular für Mitarbeiterstammdaten
- 📋 Übersicht: Tabellarische Ansicht aller Mitarbeiter

**Verbindung zu `mitarbeiter_dokumente.py`**: Sidebar enthält Kategorien + Sondereinträge (Ausdrucke, Krankmeldungen).

---

## `gui/mitarbeiter_dokumente.py` – Dokumentenbrowser

**Sidebar (Kategorien):**
- ● Verwaltung
- ● Ausbildung
- ● Sonstiges
- ────────────
- 🖨️ Ausdrucke → `DokumentBrowserWidget` für `Daten/Vordrucke/`
- 🤒 Krankmeldungen → `DokumentBrowserWidget` für `../../03_Krankmeldungen`

**6 Tabs** plus eingebettete DokumentBrowser-Instanzen für Ausdrucke und Krankmeldungen.

**Methode `_zeige_sonderkategorie(pfad)`**: Lädt DokumentBrowserWidget für den angegebenen Pfad.

---

## `gui/uebergabe.py` – Übergabeprotokolle

### Schichttypen
- **Tagdienst** (Orange `#e67e22`)
- **Nachtdienst** (Dunkelblau `#2c3e50`)

### Klassen
| Klasse | Beschreibung |
|---|---|
| `UebergabeWidget` | Hauptwidget mit Protokollliste + Formular |
| `_ProtokolListItem` | Listenelement in der Sidebar |

### Hauptfelder
- Datum, Schichttyp, Beginn/Ende
- Patienten-Anzahl, Personal
- Ereignisse, Massnahmen, Übergabe-Notiz
- Fahrzeug-Notizen (pro Fahrzeug)
- Handy-Einträge (Empfänger, Status)
- Verspätungen (Mitarbeiter, Minuten, Grund, Datum)

### E-Mail-Export
- HTML-E-Mail mit DRK-rotem Header-Banner
- Info-Tabelle: Datum, Schicht, Ersteller, Patienten-Anzahl
- Fahrzeuge-Sektion (KZ + Notiz)
- Verspätungen mit Datum
- Outlook-Entwurf via `win32com`

### Verspätungen-Logik
- Automatisches Laden aus `verspaetungen.db` für das Protokolldatum
- Blaue Einträge (aus Mitarbeiterdokumentation) immer angezeigt
- Datum-Feld bei manuellen Einträgen
- Dedup via `(mitarbeiter, dienstbeginn)`-Set

---

## `gui/vorkommnisse.py` – Vorkommnis-Berichte

### Vorkommnis-Typen
```python
VORKOMMNIS_TYPEN = [
    "PRM-Betreuung", "Medizinischer Notfall", "Sicherheitsvorfall",
    "Verspätung/Offblock", "Fahrzeugschaden", "Kommunikationsfehler", "Sonstiges"
]
```

### Bereich-Optionen
```python
BEREICH_OPTIONEN = ["Flüge", "Fahrzeuge", "Personal", "Sachschäden", "Infrastruktur", "Sonstiges"]
```

### PRM-Kategorien (IATA)
```python
PRM_KATEGORIEN = ["WCHS", "WCHR", "WCHC", "BLND", "DEAF", "DPNA", "UMNR", "STCR", "MEDA", "Sonstiges"]
```

### Word-Export
- DRK-Logo + Anschrift als Briefkopf
- Automatische Bereichs-Bezeichnung (z.B. „Flugnummer:", „Fahrzeug / Kennzeichen:")
- Mitarbeiter-Rollen-Tabelle

---

## `gui/passagieranfragen.py` – PRM-Passagieranfragen

### Features
- **Outlook-Posteingang**: Letzte 75 E-Mails (via `win32com`)
- **Automatische Extraktion** aus E-Mail-Text:
  - Name (5-Stufen-Strategie: Label, Anrede-Block, Fließtext, Von-Header)
  - E-Mail-Adresse, Flugnummer, Datum, Rückflug
  - Absender direkt aus `SenderEmailAddress` (Exchange-EX:/-Adressen werden übersprungen)
- **Anrede-Dropdown**: `–`, `Herr`, `Frau`
- **Bezug-Zeile**: `Bezug: Flug EW583, 19.03.2026`

### 4 Antwort-Szenarien
1. Alle Angaben vorhanden → Eintragungsbestätigung + Hinweise
2. Fehlende Informationen → Anforderung der 4 Pflichtfelder
3. Abholung am Parkplatz
4. Allgemeine PRM-Service-Info (5 Schritte)

### Outlook-Entwurf
```python
# create_outlook_draft()
# - DRK-Logo als CID-Inline-Bild
# - Outlook-Standardsignatur automatisch angehängt
# - Betreff: "PRM-Service – Flughafen Köln/Bonn | Name | Flug EW583 | 19.03.2026"
```

---

## `gui/schulungen_kalender.py` – Schulungskalender

### Tabs
- **📅 Kalender-Tab**: Jahresübersicht aller Schulungen
- **👥 Mitarbeiter-Liste**: Alle Mitarbeiter mit Schulungsstatus

### Mitarbeiter-Liste Features
- Freitextsuche nach Name (mit ✕-Löschen)
- Status-Filter: Alle / Abgelaufen / ≤1 Mon. / ≤2 Mon. / ≤3 Mon. / OK / Kein Eintrag
- Schulungs-Filter: Alle 14 Schulungstypen einzeln wählbar
- Matrix-Tabelle: EH, Refresher, ZÜP, Ärztl., FS-K. mit Farbkodierung
- MA ohne Einträge → grau ans Ende sortiert

### Dialoge
- **`_MitarbeiterDetailDialog`**: Doppelklick → alle 14 Schulungstypen in einer Tabelle
- **`_SchulungBearbeitenDialog`**: Datum-Picker + auto. Gültig-bis-Berechnung
  - `intervall`-Typen (EH +2J, Refresher +1J): auto-berechnet
  - `direkt`-Typen (ZÜP, Ärztl.): manuelles Gültig-bis-Feld
  - `einmalig`-Typen: kein Ablaufdatum

---

## `gui/beschwerden.py` – Beschwerden

- Beschwerde-Erfassung mit Priorität, Status, Kategorie
- Antworten/Notizen je Beschwerde
- Statusverlauf (offen → in Bearbeitung → abgeschlossen)

---

## `gui/telefonnummern.py` – Telefonverzeichnis

- Telefonbuch für DRK-interne + externe Nummern
- Kategorien, Suchfunktion
- Import aus Excel/CSV möglich

---

## `gui/dienstplan.py` – Dienstplan

- Schichtplanung (Tagdienst, Nachtdienst, 10h-Schichten)
- Bulmorfahrer-Kennzeichnung
- E-Mobby-Fahrer-Kennzeichnung
- Word-Export (Stärkemeldung)
- HTML-Export
- Parser für bestehende Word-Dokumente (`functions/dienstplan_parser.py`)

---

## `gui/splash_screen.py` – Ladebildschirm

- Gleiche Animationslogik wie `_NeskLogoWidget`
- Zeigt sich beim App-Start
- Verschwindet nach Initialisierung aller Module

---

## `gui/einstellungen.py` – Einstellungen

- Konfigurierbare App-Parameter
- Speichert in `settings`-Tabelle der `nesk3.db`
- Dark/Light Mode-Option, Farbschema

---

## `gui/backup_widget.py` – Backup-UI

- Manuelle Backups auslösen
- Backup-Liste anzeigen
- ZIP-Backup des gesamten Nesk3-Ordners erstellen
- Restore-Funktion

---

## `gui/handys_widget.py` – Handy-Verwaltung

- Übersicht aller DRK-Handys mit Status
- Zuweisung an Mitarbeiter
- Ausgabe-/Rückgabe-Protokoll
- Excel-Export
- E-Mail-Bericht

---

## `gui/call_transcription.py` – Anruf-Transkription

- Aufnahme oder Import von Audiodateien
- Transkription via Google Gemini API (`GEMINI_API_KEY`)
- Datenextraktion aus Transkript
- Speicherung in `call_transcription.db`

---

## `gui/workflow.py` – Workflow

- Aufgaben-Tracking mit Status (offen, in Bearbeitung, erledigt)
- Zuweisung an Mitarbeiter
- Fälligkeitsdaten

---

## `gui/laufzettel_dialog.py` – Laufzettel

- Dialog für Laufzettel-Generierung
- Mitarbeiter-Auswahl
- PDF-Export

---

## `gui/dokument_browser.py` – Datei-Browser

- Generisches Datei-Browser-Widget für beliebige Ordner
- Dateivorschau (PDF, Bilder, Office)
- Doppelklick öffnet Datei mit Standardprogramm

---

## Gemeinsame Styles (Tab-Leisten)

Alle Tab-Leisten verwenden einheitlich:

```python
tab_style = """
    QTabBar::tab {
        min-width: 160px;
        padding: 8px 16px;
        font-size: 12px;
        font-family: 'Segoe UI';
        color: #666;
        background: transparent;
        border-bottom: 3px solid transparent;
    }
    QTabBar::tab:selected {
        color: #1565a8;
        font-weight: bold;
        border-bottom: 3px solid #1565a8;
    }
    QTabBar::tab:hover:!selected {
        color: #1565a8;
        border-bottom: 3px solid #ccddf5;
    }
"""
```
