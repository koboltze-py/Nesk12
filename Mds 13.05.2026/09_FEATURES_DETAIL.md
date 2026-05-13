# Features im Detail (Stand: 13.05.2026)

## 1. Dashboard-Notizen

### Datenbasis
- Datenbank: `notizen.db`, Tabelle `notizen`
- Schema: `id, titel, text, datum (dd.MM.yyyy), erstellt_am, status`
- Status-Werte: `offen`, `gelesen`, `erledigt`

### Anzeige-Fenster
```
Dashboard zeigt: heute-5 Tage bis heute+10 Tage
→ Sortierung nach datum-Feld (aufsteigend)
→ Funktion: functions/notizen_db.lade_fenster()
```

### Notiz-Karten im Dashboard
Jede Notiz zeigt:
- Datum-Header (🟢 Heute, 🔵 Gestern, 📍 ältere, 📅 zukünftige)
- Titel + Text
- Status-Badge (offen/gelesen/erledigt)
- Buttons:
  - 👁 **Gelesen** → `als_gelesen(nid)`
  - ↩️ **Rückgängig** (nur bei erledigt) → `als_offen(nid)`
  - ✅ **Erledigt** → `als_erledigt(nid)`
  - 🗑 **Löschen** → `loeschen(nid)` (mit Bestätigungsdialog)

### Neue Notiz erstellen
- Button `➕ Neue Notiz` (oben rechts im Notizen-Panel)
- Felder: Titel (Pflicht), Text (optional), Datum (default: heute)
- Datum-Picker: im Format `dd.MM.yyyy`

### Archiv-Dialog
- Button `🗄️ Archiv` (neben Neue Notiz)
- Öffnet Dialog: alle Notizen, neueste zuerst
- Suchfilter: Freitext über Titel und Text
- Pro Notiz: Status, Datum, Buttons (Löschen)
- Keine Paginierung – alle Notizen auf einmal

---

## 2. Vorkommnisse

### Datenmodell
```
vorkommnisse.db / vorkommnisse-Tabelle:
- flug, typ, datum, ort
- offblock_plan, offblock_ist  (Zeitverschiebung)
- erstellt_von, ursache, ergebnis
- passagiere_json   → JSON-Array [{name, alter, ...}]
- personal_json     → JSON-Array [{name, rolle, ...}]
- chronologie_json  → JSON-Array [{uhrzeit, ereignis}]
```

### Auto-Save beim E-Mail-Dialog
```python
# In gui/vorkommnisse.py → _erstelle():
# Wenn _current_id is None und Flugnummer vorhanden:
#   → Automatisch speichern bevor Outlook-Entwurf geöffnet wird
#   → Verhindert Datenverlust wenn "Abbrechen" nach E-Mail
```

### Word-Export
- Erstellt Vorkommnis-Bericht als Word-Datei (python-docx)
- Speichert in: `Daten/` (konfigurierbar)
- E-Mail-Versand via Outlook (Anhang)

---

## 3. Handys-Modul

### Geräte-Verwaltung
- 4 Tabs: Übersicht, Historien, Berichte, Einstellungen
- Vollständige CRUD-Operationen
- Automatische Historien-Logung bei jedem Feldwechsel

### Defekt-Tracking
- Zustände: Aktiv, Defekt, Außer Betrieb, Reserve, Verloren
- Bei Defekt/Außer Betrieb/Verloren: `defekt_datum`, `defekt_beschreibung`, `defekt_gemeldet_von`
- Schadensbericht-Dialog: Word-Dokument + Outlook-Versand

### Berichte
- Excel-Tabelle aller Geräte
- Word-Zustandsbericht
- E-Mail mit Anhängen via `win32com.client`

---

## 4. Fahrzeugverwaltung

### Schäden
- Schwere: gering / mittel / schwer
- E-Mail-Versand-Tracking: `gesendet`-Flag
- Schadens-E-Mail via Outlook

### Termine
- Typen: TÜV, Inspektion, Reparatur, Hauptuntersuchung, Sonstiges
- Dashboard-Kalender: blauer Punkt für Tage mit Terminen
- Startup-Benachrichtigung: 800ms nach Start prüft App bald fällige Termine

---

## 5. Übergabe-Protokolle

### Felder
- Datum, Schicht-Typ (tagdienst/nachtdienst), Zeiten
- Patientenzahl, Personal-Text
- Ereignisse, Maßnahmen, Übergabe-Notiz
- Ersteller, Abzeichner, Status (offen/abgeschlossen)

### Sub-Daten
- **Fahrzeug-Notizen**: Pro Fahrzeug eine Notiz (UNIQUE constraint)
- **Handy-Einträge**: Anzahl + Notiz pro Gerät
- **Verspätungen**: Mitarbeiter + Soll/Ist-Zeit

### Archivierung
- `archiviert`-Flag: Protokoll aus aktiver Liste entfernen
- Archivierte Protokolle bleiben in DB (kein Löschen)

---

## 6. Bericht (Stärkemeldung)

### Kombinierter Monatsbericht
Aus 4 Datenquellen:
1. `uebergabe_verspaetungen` → Verspätungen-Tabelle
2. `schulungen`-Tabelle → Schulungsstunden
3. `tages_einsaetze` → Einsatzzahlen
4. `tages_pax` → Passagierzahlen

### Export
- Excel-Datei (openpyxl)
- E-Mail-Versand via Outlook

### Auto-Export
- Am **1. jeden Monats** beim App-Start
- Speichert Vormonats-Bericht automatisch
- Pfad: `Daten/Spät/Monatsliste/Auto Liste/`

---

## 7. Dienstplan

### Word-Import
- Liest strukturierte Word-Dokumente ein (`functions/dienstplan_parser.py`)
- Parst Tabellen-Format der Wochenpläne
- Importiert MA-Namen, Datum, Schichtzeiten

### Anzeige
- Wochenansicht mit Farbkodierung nach Schichttyp
- Regulär / Nacht / Bereitschaft

---

## 8. Call-Transkription (KI)

### Technologie
- Google Gemini API (`GEMINI_API_KEY` in config.py)
- Model: Gemini (aktuell konfiguriertes Modell)

### Workflow
1. Audio-Datei hochladen oder Aufnahme starten
2. Senden an Gemini API
3. Transkriptions-Text anzeigen
4. Textbausteine einfügen
5. Speichern in `call_transcription.db`

### Datenbasis
```
call_transcription.db:
  call_logs     – Transkriptionen
  textbausteine – Vordefinierte Text-Templates
```

---

## 9. Code-19-Protokoll

- Sicherheitsvorfall-Dokumentation
- Zeitstempel-geführt mit Uhren-Animation
- Protokoll-Liste und Druckexport

---

## 10. Passagier-Modul

### tages_pax-Tabelle
- Tägliche Passagierzahlen (YYYY-MM-DD, eindeutig pro Tag)
- Wird im Dashboard als Statistik angezeigt
- Geht in den Monatsbericht ein

### Anfragen & Beschwerden
- `passagieranfragen.py` verknüpft mit Beschwerden-Modul
- Vollständiges Beschwerdemanagement mit Kategorisierung, Status, Maßnahmen

---

## 11. SanMat (Sanitätsmaterial)

- Inventar mit Ablaufdaten
- Bestandsüberwachung
- Datenbank: `sanmat.db` (`database/sanmat_db.py`)
- Widget in `gui/sanmat/`

---

## 12. Dienstliches / PSA / Stellungnahmen

### Dienstliches
- Einsatz-Protokolle
- Dienstanweisungen (`functions/dienstanweisungen_db.py`)

### PSA (Persönliche Schutzausrüstung)
- PSA-Verstöße erfassen (`psa.db`)
- Turso-Sync

### Stellungnahmen
- Zu Beschwerden oder Vorfällen
- HTML-Export (`functions/stellungnahmen_html_export.py`)

---

## 13. E-Mobby Integration

- `functions/emobby_functions.py`
- Fahrerliste-Verwaltung
- Eingestellt in `gui/einstellungen.py`
