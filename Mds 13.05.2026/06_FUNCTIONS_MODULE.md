# Functions-Module – API-Referenz (Stand: 13.05.2026)

## functions/notizen_db.py → `notizen.db`

```python
speichern(titel: str, text: str = "", datum: str = "") -> int
    # Neue Notiz anlegen. datum: dd.MM.yyyy (default: heute)
    # Gibt die neue ID zurück

als_gelesen(nid: int)
    # Status → 'gelesen'

als_erledigt(nid: int)
    # Status → 'erledigt'

als_offen(nid: int)
    # Status → 'offen' (Erledigt rückgängig)

loeschen(nid: int)
    # Notiz dauerhaft löschen

lade_aktive() -> list[Row]
    # Notizen der letzten 5 Tage (nach erstellt_am)

lade_fenster() -> list[Row]
    # Notizen von heute-5 bis heute+10 (nach datum-Feld), aufsteigend sortiert
    # → wird vom Dashboard verwendet

lade_zukunft() -> list[Row]
    # Notizen von morgen bis heute+10

lade_alle() -> list[Row]
    # Alle Notizen, neueste zuerst (für Archiv-Dialog)
```

---

## functions/vorkommnisse_db.py → `vorkommnisse.db`

```python
speichern(daten: dict) -> int
    # Neues Vorkommnis speichern
    # daten: {flug, typ, datum, ort, offblock_plan, offblock_ist,
    #         erstellt_von, ursache, ergebnis,
    #         passagiere: [...], personal: [...], chronologie: [...]}
    # JSON-Arrays werden via json.dumps() serialisiert
    # Turso-Push nach Speicherung

aktualisieren(vid: int, daten: dict)
    # Vorhandenes Vorkommnis aktualisieren + Turso-Push

loeschen(vid: int)
    # Dauerhaft löschen

lade_alle() -> list[Row]
    # Alle Vorkommnisse, neueste zuerst

lade_eins(vid: int) -> Row | None
    # Einzelnes Vorkommnis laden
```

---

## functions/beschwerden_db.py → `beschwerden.db`

**Konstanten:**
```python
KATEGORIEN = ["Verhalten Mitarbeiter", "Wartezeit / Reaktionszeit",
              "Kommunikation", "Ausstattung / Material", "Hygiene",
              "Dokumentation", "Behandlung / Versorgung", "Sonstiges"]
PRIORITAETEN = ["Niedrig", "Mittel", "Hoch", "Kritisch"]
STATUS_OPTIONEN = ["Offen", "In Bearbeitung", "Erledigt", "Abgewiesen"]
QUELLEN = ["Freitext", "Word-Datei", "PDF-Datei", "E-Mail"]
```

```python
speichern(daten: dict) -> int
aktualisieren(bid: int, daten: dict)
loeschen(bid: int)
lade_alle() -> list[Row]
lade_eins(bid: int) -> Row | None
massnahme_hinzufuegen(bid: int, text: str)
stellungnahme_hinzufuegen(bid: int, text: str, ersteller: str)
```

---

## functions/handys_db.py → `handys.db`

**Zustände:** `Aktiv`, `Defekt`, `Außer Betrieb`, `Reserve`, `Verloren`

```python
get_all() -> list[Row]
    # Alle Handys, nach Inventarnummer sortiert

get_by_id(handy_id: int) -> Row | None

get_by_inventarnummer(inv: str) -> Row | None

create(daten: dict) -> int
    # Felder: inventarnummer (UNIQUE), hersteller, modell, rufnummer,
    #         sim_nummer, standort, zustand, defekt_beschreibung,
    #         defekt_datum, defekt_gemeldet_von, anschaffungsdatum,
    #         notizen, kartennummer, pin, pin2, puk, puk2

update(handy_id: int, daten: dict)
    # Schreibt Änderungen und loggt jedes geänderte Feld in handys_historie

delete(handy_id: int)

get_historie(handy_id: int) -> list[Row]
    # Alle Änderungen für ein Gerät

filter_by_zustand(zustand: str) -> list[Row]
```

---

## functions/handys_bericht.py

```python
erstelle_bericht(handys: list[dict], pfad: str) -> str
    # Erstellt Word-Dokument (python-docx)
    # Tabelle: Inventarnummer, Zustand, Rufnummer, Standort, Notizen
    # Gibt den Dateipfad zurück
```

---

## functions/handys_email.py

```python
erstelle_outlook_entwurf(betreff: str, body: str, anlagen: list[str])
    # Öffnet Outlook-Entwurf via win32com.client
    # anlagen: Liste von Dateipfaden
```

---

## functions/handys_excel_export.py

```python
export_excel(handys: list[dict], pfad: str) -> str
    # Erstellt Excel-Datei mit openpyxl
    # Tabellenblatt: Geräteübersicht
    # Gibt Dateipfad zurück
```

---

## functions/verspaetung_db.py → nesk3.db (uebergabe_verspaetungen)

```python
speichern(protokoll_id: int, mitarbeiter: str, soll: str, ist: str) -> int
lade_fuer_protokoll(protokoll_id: int) -> list[Row]
loeschen(vid: int)
lade_monat(monat: int, jahr: int) -> list[Row]
    # Für Monats-Export
```

---

## functions/verspaetung_functions.py

```python
exportiere_monats_excel(monat: int, jahr: int, pfad: str) -> str
    # Excel-Export aller Verspätungen eines Monats
    # Format: Datum, Mitarbeiter, Soll, Ist, Differenz
    
auto_export_vormonat()
    # Wird am 1. jeden Monats aufgerufen
    # Speichert in: Daten/Spät/Monatsliste/Auto Liste/
```

---

## functions/mitarbeiter_functions.py → nesk3.db

```python
lade_alle() -> list[Row]
lade_aktive() -> list[Row]
speichern(daten: dict) -> int
aktualisieren(mid: int, daten: dict)
loeschen(mid: int)
suche(text: str) -> list[Row]
```

---

## functions/dienstplan_functions.py → nesk3.db

```python
lade_woche(datum: str) -> list[Row]
    # datum: YYYY-MM-DD (erster Tag der Woche)
lade_monat(monat: int, jahr: int) -> list[Row]
speichern(mitarbeiter_id: int, datum: str, start: str, ende: str, ...) -> int
loeschen(did: int)
```

---

## functions/dienstplan_parser.py

```python
parse_word_datei(pfad: str) -> list[dict]
    # Liest Word-Dienstplan ein (python-docx)
    # Gibt Liste von {mitarbeiter, datum, start, ende, position} zurück
```

---

## functions/fahrzeug_functions.py → nesk3.db

```python
lade_fahrzeuge(nur_aktive=True) -> list[Row]
lade_status(fahrzeug_id: int) -> Row | None
lade_schaeden(fahrzeug_id: int) -> list[Row]
lade_termine(fahrzeug_id: int) -> list[Row]
speichern_fahrzeug(daten: dict) -> int
aktualisieren_status(fahrzeug_id: int, status: str, grund: str)
schaden_hinzufuegen(fahrzeug_id: int, daten: dict) -> int
termin_hinzufuegen(fahrzeug_id: int, daten: dict) -> int
```

---

## functions/schulungen_db.py → nesk3.db

```python
lade_monat(monat: int, jahr: int) -> list[Row]
speichern(daten: dict) -> int
loeschen(sid: int)
lade_alle() -> list[Row]
```

---

## functions/uebergabe_functions.py → nesk3.db

```python
lade_protokolle(archiviert=False) -> list[Row]
lade_protokoll(pid: int) -> Row | None
speichern(daten: dict) -> int
aktualisieren(pid: int, daten: dict)
abschliessen(pid: int)
archivieren(pid: int)
```

---

## functions/telefonnummern_db.py → nesk3.db

```python
lade_alle() -> list[Row]
suche(text: str) -> list[Row]
speichern(daten: dict) -> int
loeschen(tid: int)
importiere_csv(pfad: str) -> int
```

---

## functions/settings_functions.py → nesk3.db

```python
get(schluessel: str, default: str = "") -> str
    # Einstellung lesen

set(schluessel: str, wert: str)
    # Einstellung schreiben / anlegen

# Bekannte Schlüssel:
# "dienstplan_ordner"  → Pfad zum Dienstplan-Ordner
```

---

## functions/mail_functions.py

```python
erstelle_outlook_entwurf(an: str, betreff: str, body: str,
                          anlagen: list[str] | None = None)
    # Öffnet Outlook via win32com.client
    # Erstellt Draft (nicht automatisch senden)
```

---

## functions/staerkemeldung_export.py

```python
exportiere_excel(monat: int, jahr: int, pfad: str) -> str
    # Kombinierter Bericht: Verspätungen + Schulungen + Einsätze + Passagiere
    
def erstelle_dashboard_export(monat: int, jahr: int) -> dict
    # Gibt Zusammenfassung für Dashboard-Kacheln zurück
```

---

## database/connection.py

```python
def get_connection() -> sqlite3.Connection:
    """Gibt eine SQLite-Verbindung zu nesk3.db zurück."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn
```

---

## database/models.py

Datenzugriffsschicht für alle Tabellen in `nesk3.db`.  
Stellt typisierte Funktionen bereit, die `get_connection()` nutzen.

---

## database/pax_db.py

```python
# Tabelle: tages_pax
speichern(datum: str, pax_zahl: int)
    # datum: YYYY-MM-DD
    # UPSERT (INSERT OR REPLACE)

lade(datum: str) -> Row | None
lade_monat(monat: int, jahr: int) -> list[Row]
```

---

## database/sanmat_db.py

SanMat = Sanitätsmaterial-Datenbank.  
Verwaltet Inventar mit Ablaufdaten und Bestandszahlen.

---

## database/sonderaufgaben_db.py

Verwaltet Sonderaufgaben (außerplanmäßige Aufgaben).  
Verbunden mit `gui/sonderaufgaben.py`.

---

## database/export_historie_db.py

Speichert eine Historik aller getätigten Exporte (Excel, Word, PDF).  
Verhindert Doppel-Exporte und erlaubt Audit-Trail.
