# GUI-Module – Beschreibung aller Widgets (Stand: 02.06.2026)

## gui/main_window.py

**Klassen:** `_NeskLogoWidget`, `SidebarButton`, `MainWindow`

```python
class MainWindow(QMainWindow):
    # Fenstergröße: 1280×800 (min. 900×600)
    # Sidebar: 220px, FIORI_SIDEBAR_BG (#354a5e)
    # Content: QStackedWidget, 15 Seiten
    
    def _build_ui(self): ...        # Sidebar + Content zusammenbauen
    def _build_sidebar(self): ...   # Logo + ScrollArea + Nav-Buttons
    def _build_content(self): ...   # QStackedWidget mit allen Widgets
    def _navigate(self, index): ... # Seite wechseln + Button-Status
    def _check_termine_startup(self): ...  # Termin-Prüfung nach Start
```

---

## gui/dashboard.py

**Klassen:** `_TerminKalender`, `StatCard`, `_SkyWidget`, `DashboardWidget`

### StatCard
```python
StatCard(title, value, icon, color)
# Weiße Karte, farbiger linker Rand, Titel + Icon oben, große Zahl unten
# .set_value(value)  → Wert aktualisieren
```

### _TerminKalender
- Erbt von `QCalendarWidget`
- Blauer Punkt = Fahrzeug-Termin (`set_termin_dates()`)
- Grüner Punkt = Notiz vorhanden (`set_notiz_dates()`)

### DashboardWidget
```python
def _zeige_notizen(self):
    # Nutzt functions.notizen_db.lade_fenster()
    # Zeigt Notizen von heute-5 bis heute+10
    # Header: 🟢 Heute, 🔵 Gestern, 📍 ältere Tage, 📅 Morgen/zukünftig
    # Jede Notiz: [Titel] + Buttons [👁 Gelesen] [↩️ Rückgängig] [✅ Erledigt] [🗑 Löschen]
    
def _notiz_als_offen(self, nid: int):
    # Erledigt rückgängig → Status auf 'offen'
    
def _notiz_archiv_dialog(self):
    # Zeigt alle Notizen im Archiv-Dialog mit Suchfilter
```

---

## gui/mitarbeiter.py

**Klasse:** `MitarbeiterHauptWidget`

- 2 Tabs: **Mitarbeiter** (Tabelle, CRUD) + **Dokumente** (PDF/Word je MA)
- Felder: Vorname, Nachname, Personalnummer, Funktion, Position, Abteilung, E-Mail, Telefon, Status
- Funktionen: Hinzufügen, Bearbeiten, Löschen, Filtern, Excel-Export
- Dokumente-Tab: Datei-Upload/-Download per MA

---

## gui/dienstplan.py

**Klasse:** `DienstplanWidget`

- Lädt Word-Dienstpläne via `functions/dienstplan_parser.py`
- Anzeige: Wochenübersicht mit Farbkodierung
- Export: HTML (`functions/dienstplan_html_export.py`)
- Einstellung: Ordnerpfad für Dienstplandateien (in `settings`-Tabelle)

---

## gui/fahrzeuge.py

**Klasse:** `FahrzeugeWidget`

- 3 Sub-Tabs:
  - **Fahrzeuge**: CRUD, Kennzeichen, Typ, TÜV-Datum
  - **Schäden**: Pro Fahrzeug, Schwere (gering/mittel/schwer), E-Mail-Versand via Outlook
  - **Termine**: Typ (TÜV/Inspektion/...), Fälligkeitsdatum, Erinnerung
- Termin-Integration: Kalenderpunkte im Dashboard

---

## gui/vorkommnisse.py

**Klasse:** `VorkommnisseWidget`

```python
# Felder des Formulars:
# flug, typ, datum, ort, offblock_plan, offblock_ist
# erstellt_von, ursache, ergebnis
# Tabellen: passagiere (JSON), personal (JSON), chronologie (JSON)

def _erstelle(self):
    # E-Mail-Dialog:
    # AUTO-SAVE: wenn _current_id is None und Flugnummer vorhanden
    #   → speichern() vor dem Outlook-Entwurf
```

---

## gui/handys_widget.py

**Klasse:** `HandysWidget`

```
4 Tabs:
1. "Geräteübersicht"
   - QTableWidget mit allen Handys
   - Filter nach Zustand (Aktiv/Defekt/...)
   - Doppelklick: Detailformular öffnet sich
   - Zustände: Aktiv, Defekt, Außer Betrieb, Reserve, Verloren
   
2. "Historien"
   - Alle Änderungen aus handys_historie
   - Filter nach Inventarnummer

3. "Berichte"
   - Word-Bericht erstellen (functions/handys_bericht.py)
   - Excel-Export (functions/handys_excel_export.py)
   - E-Mail via Outlook (functions/handys_email.py)

4. "Einstellungen"
   - Export-Pfad: HANDYS_EXPORT_PATH
```

---

## gui/schadensbericht_dialog.py

- Dialog für Schadensmeldungen einzelner Handys
- Erstellt Word-Dokument und sendet per Outlook

---

## gui/uebergabe.py

**Klasse:** `UebergabeWidget`

- Protokoll-Felder: Datum, Schicht-Typ (Tag/Nacht), Zeiten, Patientenzahl, Personal
- Sub-Abschnitte: Fahrzeug-Notizen, Handy-Einträge, Verspätungen
- Archivierung: `archiviert`-Flag setzt Protokoll als abgeschlossen

---

## gui/passagiere.py & gui/passagieranfragen.py

- **passagiere.py**: Passagier-Datenbank (pax_db), Suche, Eintrag
- **passagieranfragen.py**: Anfragen-Tracking + Beschwerden-Integration

---

## gui/beschwerden.py

**Klasse:** Beschwerden-Widget

- Kategorien: Verhalten, Wartezeit, Kommunikation, Ausstattung, Hygiene, Dokumentation, Behandlung, Sonstiges
- Prioritäten: Niedrig, Mittel, Hoch, Kritisch
- Status: Offen, In Bearbeitung, Erledigt, Abgewiesen
- Maßnahmen als Sub-Liste, Stellungnahmen als Sub-Liste

---

## gui/code19.py

**Klasse:** `Code19Widget`

- Code-19-Protokoll-Führung (Sicherheitsvorfälle)
- Zeitstempel-Animation (Uhren-Widget)
- Protokoll-Liste und Detail-Ansicht

---

## gui/bericht.py

**Klasse:** `BerichtWidget`

- Kombinierter Monatsbericht:
  - Verspätungen (uebergabe_verspaetungen)
  - Schulungen (schulungen_db)
  - Einsätze (tages_einsaetze)
  - Passagiere (tages_pax)
- Export: Excel + E-Mail via Outlook (`functions/staerkemeldung_export.py`)

---

## gui/schulungen_kalender.py

- Schulungskalender mit Monatsansicht
- Schulungen anlegen, bearbeiten, löschen
- Farbkodierung nach Schulungstyp

---

## gui/einstellungen.py

**Klasse:** `EinstellungenWidget`

- App-Name, Version anzeigen
- Dienstplan-Ordner konfigurieren
- E-Mobby-Fahrerliste konfigurieren (`functions/emobby_functions.py`)
- Turso-Sync-Status

---

## gui/backup_widget.py

**Klasse:** `BackupWidget`

- Manuellen Backup erstellen
- Backup-Liste anzeigen (mit Datum, Größe)
- Backup wiederherstellen
- Nutzt `backup/backup_manager.py`

---

## gui/dienstliches.py

**Klasse:** `DienstlichesWidget`

- Dienstliche Protokolle (Einsätze)
- Dienstanweisungen (dienstanweisungen_db)

---

## gui/aufgaben_haupt.py

**Klasse:** `AufgabenHauptWidget`

```
2 Tabs:
- "Tag" → AufgabenTagWidget (gui/aufgaben_tag.py)
  - Tägliche Checklisten, Tagesaufgaben
- "Haupt" → AufgabenWidget (gui/aufgaben.py)
  - Aufgaben Nacht, Sonderaufgaben, AOCC-Aufgaben
```

---

## gui/checklisten.py

- Checklisten-Editor und -Anzeige
- Checklisten abhaken und abschließen

---

## gui/sonderaufgaben.py

- Sonderaufgaben erfassen und verwalten
- Verbunden mit `database/sonderaufgaben_db.py`

---

## gui/telefonnummern.py

**Klasse:** `TelefonnummernWidget`

- FKB Gate-/Check-In-Nummern
- DRK-Kontakte
- Suche / Filterung
- Datenbasis: `functions/telefonnummern_db.py`

---

## gui/call_transcription.py

**Klasse:** (Transkriptions-Widget)

- Audioaufnahme oder Datei-Upload
- Transkription via Google Gemini API
- Ergebnis-Anzeige und Speicherung
- Datenbasis: `functions/call_transcription_db.py`

---

## gui/hilfe_dialog.py

- Modaler Dialog mit Hilfetext/Dokumentation
- Tastaturkürzeln-Übersicht

---

## gui/splash_screen.py

- Animierter Startbildschirm (identisches Logo wie Sidebar)
- Lade-Fortschrittsbalken
- Versionsnummer
- Schließt sich nach App-Start selbst

---

## gui/slot_machine.py & gui/slot_symbols.py

**Easter Egg:** "Alice's Wunderrad"

- Aktivierung: Doppelklick auf NeSk-Logo in Sidebar
- Slot-Machine-Animation mit benutzerdefinierten Symbolen
- Rein dekorativ, kein Produktivbezug

---

## gui/sanmat/ (Unterordner)

- SanMat (Sanitätsmaterial) Verwaltung
- Verbunden mit `database/sanmat_db.py`
- Inventur, Ablaufdaten, Bestandsübersicht

---

## gui/dokument_browser.py

- Generischer Datei-Browser für Dokumente
- Öffnen, Vorschau, Herunterladen von Dateien aus `Daten/`

---

## gui/mitarbeiter_dokumente.py

- Dokumenten-Verwaltung je Mitarbeiter
- Upload/Download per Mitarbeiter-ID
- Verbunden mit `functions/mitarbeiter_dokumente_functions.py`

---

## gui/workflow.py  *(neu v3.8.0)*

**Klassen:** `WorkflowWidget`, `_DetailDialog`, `_LadeThread`

```python
class WorkflowWidget(QWidget):
    # Monat-Selektor: Dropdown + "Neuer Monat" + "Monat entfernen"
    # Geladene SM/DP-Pfade werden je Monat in DB gespeichert (workflow_session)
    # Beim nächsten Start: letzter Monat wird automatisch geladen + Abgleich gestartet

    def _monat_gewaehlt(self, idx): ...    # Lädt Session, startet Abgleich
    def _neuer_monat(self): ...            # Dialog: Monat + Jahr wählen
    def _monat_loeschen(self): ...         # Entfernt Monat aus Session (Carmen/Notizen bleiben)
    def _speichere_session_wenn_monat(self): ...  # Auto-Speichern nach Datei-Laden

    def _starte_abgleich(self): ...        # Startet _LadeThread
    def _zeige_ergebnisse(self, ergebnisse): ...   # Füllt Tabelle
    def _detail_zeigen(self, index): ...   # Öffnet _DetailDialog (Doppelklick)
    def _kontext_menu_erg(self, pos): ...  # Carmen-Toggle + Notiz + Dateien öffnen

class _DetailDialog(QDialog):
    # 9 Spalten: Status | SM-Name | SM-Dienst | SM-Von | SM-Bis |
    #            DP-Name | DP-Dienst | DP-Von | DP-Bis (+ DP-Notiz)
    # SL/Dispo-Gruppe oben, dann Trennzeile "── Betreuer ──", dann Betreuer
    # Sortierung Dispo: SL+DT3 → SL+DN3 → DT → DN
    # Sortierung Betreuer: T → T10 → T8 → N → N10

class _LadeThread(QThread):
    # Lädt SM (Word) + DP (Excel) parallel, führt Abgleich durch
    # Signals: fortschritt(str), fertig(list), fehler(str)
```

### Persistenz
- `database/workflow_db.py`: zwei Tabellen
  - `workflow_tag`: Carmen-Status + Notiz pro Tag+SM-Datei
  - `workflow_session`: SM/DP-Pfadlisten pro Monat

### Abgleich-Logik
- Fuzzy-Matching via `difflib.get_close_matches` (cutoff 0.82)
- Pre-Resolve für compound-Namen (El Mojahid, Moeeni Mahvelati)
- Abbrev-Einträge (z.B. "Blei Pa") werden vor einfachen Namen gematcht
- Rufbereitschaft (R) und "lars peters" automatisch ausgeschlossen
- `functions/dienstplan_parser.py`: DP-Notizen aus Zellen rechts von ENDE
