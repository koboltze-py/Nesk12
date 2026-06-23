# Functions-Module – Geschäftslogik & Datenbankoperationen

## Übersicht

Alle Datenbankoperationen und Geschäftslogik sind von den GUI-Modulen getrennt in `functions/` abgelegt.

---

## `functions/mitarbeiter_functions.py`

**Funktionen:**

| Funktion | Signatur | Beschreibung |
|---|---|---|
| `lade_alle_mitarbeiter` | `() → list[dict]` | Alle aktiven Mitarbeiter aus DB |
| `lade_mitarbeiter_namen` | `() → list[str]` | Nur Vor- + Nachname als Liste |
| `lade_mitarbeiter_by_id` | `(id: int) → dict\|None` | Einzelner Mitarbeiter |
| `speichere_mitarbeiter` | `(daten: dict) → int` | INSERT, gibt neue ID zurück |
| `aktualisiere_mitarbeiter` | `(id: int, daten: dict) → bool` | UPDATE |
| `loesche_mitarbeiter` | `(id: int) → bool` | DELETE (soft: status='inaktiv') |
| `suche_mitarbeiter` | `(query: str) → list[dict]` | Volltextsuche Name/Personalnr. |

---

## `functions/fahrzeug_functions.py`

**Funktionen:**

| Funktion | Signatur | Beschreibung |
|---|---|---|
| `lade_alle_fahrzeuge` | `() → list[dict]` | Alle aktiven Fahrzeuge |
| `aktueller_status` | `(fahrzeug_id: int) → str` | Letzter Status-Eintrag |
| `setze_status` | `(fahrzeug_id, status, von, bis, grund) → int` | Neuen Status speichern |
| `lade_schaeden_letzte_tage` | `(tage: int) → list[dict]` | Schäden der letzten N Tage |
| `markiere_schaden_gesendet` | `(id: int) → None` | `gesendet=1` setzen |
| `lade_termine` | `(fahrzeug_id: int) → list[dict]` | Wartungstermine |
| `alle_termin_dates` | `() → set[str]` | Alle Termindaten als `YYYY-MM-DD`-Set |
| `erstelle_termin` | `(fahrzeug_id, typ, datum, notiz) → int` | Neuen Termin anlegen |
| `erledige_termin` | `(id: int) → None` | `erledigt=1` setzen |

---

## `functions/dienstplan_functions.py`

**Funktionen:**

| Funktion | Signatur | Beschreibung |
|---|---|---|
| `lade_dienstplan_fuer_datum` | `(datum: str) → list[dict]` | Alle Schichten eines Tages |
| `lade_dienstplan_fuer_woche` | `(von: str, bis: str) → list[dict]` | Schichten in Zeitraum |
| `speichere_schicht` | `(daten: dict) → int` | Neue Schicht anlegen |
| `loesche_schicht` | `(id: int) → bool` | Schicht löschen |
| `get_mitarbeiter_fuer_datum` | `(datum: str) → list[str]` | Mitarbeiternamen an Datum (für Dropdowns) |

---

## `functions/uebergabe_functions.py`

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `erstelle_protokoll(daten)` | Neues Übergabeprotokoll anlegen → ID |
| `aktualisiere_protokoll(id, daten)` | Protokoll aktualisieren |
| `lade_protokolle()` | Alle Protokolle (neueste zuerst) |
| `lade_protokoll_by_id(id)` | Einzelnes Protokoll mit allen Unter-Daten |
| `loesche_protokoll(id)` | Protokoll löschen (CASCADE auf Unter-Tabellen) |
| `schliesse_protokoll_ab(id)` | `status='abgeschlossen'` setzen |
| `speichere_fahrzeug_notizen(protokoll_id, notizen)` | Fahrzeug-Notizen speichern |
| `lade_fahrzeug_notizen(protokoll_id)` | Fahrzeug-Notizen laden |
| `speichere_handy_eintraege(protokoll_id, eintraege)` | Handy-Einträge speichern |
| `lade_handy_eintraege(protokoll_id)` | Handy-Einträge laden |
| `speichere_verspaetungen(protokoll_id, liste)` | Verspätungsliste speichern |
| `lade_verspaetungen(protokoll_id)` | Verspätungen laden |

---

## `functions/verspaetung_db.py`

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `verspaetung_speichern(daten)` | Verspätung in `verspaetungen.db` speichern |
| `lade_verspaetungen_fuer_datum(datum)` | Alle Verspätungen eines Tages |
| `lade_verspaetungen_letzter_zeitraum(tage)` | Verspätungen der letzten N Tage |

**Datenbankpfad**: `database SQL/verspaetungen.db`

**Schema:**
```sql
CREATE TABLE IF NOT EXISTS verspaetungen (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    mitarbeiter TEXT NOT NULL,
    dienstbeginn TEXT DEFAULT '',
    verspaetung_min INTEGER DEFAULT 0,
    grund       TEXT DEFAULT '',
    datum       TEXT NOT NULL,
    erstellt_am TEXT DEFAULT (datetime('now','localtime'))
);
```

---

## `functions/vorkommnisse_db.py`

**Datenbankpfad**: `database SQL/vorkommnisse.db`

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `speichere_vorkommnis(daten)` | Neues Vorkommnis → ID |
| `lade_alle_vorkommnisse()` | Alle Vorkommnisse (neueste zuerst) |
| `lade_vorkommnis_by_id(id)` | Einzelnes Vorkommnis |
| `aktualisiere_vorkommnis(id, daten)` | Vorkommnis updaten |
| `loesche_vorkommnis(id)` | Vorkommnis löschen |

---

## `functions/beschwerden_db.py`

**Datenbankpfad**: `database SQL/beschwerden.db`

**Tabellen**: `beschwerden`, `beschwerde_antworten`

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `speichere_beschwerde(daten)` | Neue Beschwerde → ID |
| `lade_alle_beschwerden()` | Alle Beschwerden |
| `aktualisiere_beschwerde(id, daten)` | Beschwerde updaten |
| `loesche_beschwerde(id)` | Beschwerde + Antworten löschen |
| `speichere_antwort(beschwerde_id, text, autor)` | Neue Antwort |
| `lade_antworten(beschwerde_id)` | Alle Antworten zu einer Beschwerde |

---

## `functions/schulungen_db.py`

**Datenbankpfad**: `database SQL/mitarbeiter.db` (Tabelle `schulungseintraege`)

**14 Schulungstypen** mit Intervall-Logik:

| Typ | Kategorie | Gültigkeitsdauer |
|---|---|---|
| EH | intervall | +2 Jahre |
| Refresher | intervall | +1 Jahr |
| ZÜP | direkt | manuell |
| Ärztl. Untersuchung | direkt | manuell |
| Führerschein Klasse K | direkt | manuell |
| _+ weitere 9 Typen_ | einmalig | kein Ablauf |

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `lade_alle_schulungseintraege()` | Alle Einträge |
| `lade_schulungen_fuer_mitarbeiter(ma_id)` | Alle Schulungen eines MA |
| `lade_mitarbeiter_mit_schulungen()` | MA + aktuellste Einträge pro Typ |
| `speichere_schuleintrag(daten)` | INSERT/UPDATE Schuleintrag |
| `berechne_status(typ, gueltig_bis)` | `'ok'|'warnung'|'abgelaufen'|'kein_eintrag'` |

---

## `functions/telefonnummern_db.py`

**Datenbankpfad**: `database SQL/telefonnummern.db`

**Tabellen**: `telefonnummern`, `tel_import_log`

**Schema:**
```sql
CREATE TABLE IF NOT EXISTS telefonnummern (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    nummer      TEXT NOT NULL,
    kategorie   TEXT DEFAULT '',
    notiz       TEXT DEFAULT '',
    erstellt_am TEXT DEFAULT (datetime('now','localtime'))
);
```

---

## `functions/handys_db.py`

**Datenbankpfad**: `database SQL/handys.db`

**Tabellen**: `handys`, `handys_historie`

**Schema:**
```sql
CREATE TABLE IF NOT EXISTS handys (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    nummer      TEXT DEFAULT '',
    typ         TEXT DEFAULT '',
    status      TEXT DEFAULT 'verfuegbar'
                CHECK (status IN ('verfuegbar','ausgegeben','defekt','gesperrt')),
    aktueller_nutzer TEXT DEFAULT '',
    notiz       TEXT DEFAULT '',
    erstellt_am TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS handys_historie (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    handy_id    INTEGER REFERENCES handys(id) ON DELETE CASCADE,
    aktion      TEXT NOT NULL,   -- 'ausgegeben' | 'zurueck' | 'defekt' ...
    mitarbeiter TEXT DEFAULT '',
    datum       TEXT NOT NULL,
    notiz       TEXT DEFAULT '',
    erstellt_am TEXT DEFAULT (datetime('now','localtime'))
);
```

---

## `functions/handys_email.py`

- Erstellt HTML-E-Mail-Bericht der Handy-Ausgaben
- Sendet via `win32com.client.Dispatch("Outlook.Application")`

---

## `functions/handys_excel_export.py`

- Exportiert Handy-Übersicht als Excel (`.xlsx`)
- Speichert in `HANDYS_EXPORT_PATH` (`Daten/Handys/`)

---

## `functions/mail_functions.py`

**Funktion `create_outlook_draft(to, subject, html_body, attachments=None)`:**
```python
# Erstellt einen Outlook-Entwurf (nicht direkt gesendet)
# - DRK-Logo als CID-Inline-Bild eingebettet
# - Outlook-Standardsignatur automatisch angehängt
# - to: Empfänger-E-Mail-Adresse
# - attachments: Liste von Dateipfaden
```

---

## `functions/dienstplan_functions.py` – Stärkemeldung

**Funktionen:**

| Funktion | Beschreibung |
|---|---|
| `erstelle_staerkemeldung(datum)` | Word-Dokument der Tagesbesetzung |
| `berechne_staerke(datum)` | Anzahl MA je Funktion für einen Tag |

---

## `functions/staerkemeldung_export.py`

- Word-Export der Stärkemeldung
- Kopfzeile: DRK-Logo, Datum, Schicht
- Tabelle: Funktion | Name | Beginn | Ende

---

## `functions/dienstplan_html_export.py`

- HTML-Tabellenexport des Dienstplans
- Farbkodierung nach Schichttyp
- Kann als E-Mail-Anhang oder Datei gespeichert werden

---

## `functions/dienstplan_parser.py`

- Liest bestehende Word-Dienstplan-Dokumente
- Extrahiert Mitarbeiternamen, Zeiten und Rollen
- Importiert Daten in die SQLite-DB

---

## `functions/archiv_functions.py`

- Archiviert abgeschlossene Protokolle und Dokumente
- Speichert in `archiv.db`

---

## `functions/dokument_archiv.py`

- Verwaltet das Dokument-Archiv-Verzeichnis
- Suche, Kategorisierung, Datum-basierte Ablage

---

## `functions/mitarbeiter_dokumente_functions.py`

- Hilfsfunktionen für den Dokumentenbrowser
- Dateipfad-Auflösung für Mitarbeiter-Dokumente

---

## `functions/mitarbeiter_sync.py`

- Synchronisiert Mitarbeiterdaten zwischen `nesk3.db` und `mitarbeiter.db`
- Wird beim App-Start aufgerufen

---

## `functions/laufzettel_functions.py`

- Erstellt Laufzettel als Word-Dokument
- Mitarbeiter-Auswahl, Stempel, Unterschrift-Feld

---

## `functions/settings_functions.py`

```python
def get_setting(schluessel: str, default: str = "") -> str:
    """Liest einen Wert aus der settings-Tabelle."""

def set_setting(schluessel: str, wert: str) -> None:
    """Schreibt/Aktualisiert einen Wert in der settings-Tabelle."""
```

---

## `functions/emobby_functions.py`

- E-Mobby-Fahrerdaten für Dienstplan
- Kennzeichnung im Sonderaufgaben-Formular (EM-Suffix)

---

## `functions/call_transcription_db.py`

**Datenbankpfad**: `database SQL/call_transcription.db`

**Tabellen**: `call_logs`, `textbausteine`

- Speichert Anruf-Protokolle mit Transkription
- Textbausteine für häufige Antworten

---

## `functions/psa_db.py`

**Datenbankpfad**: `database SQL/psa.db`

- PSA-Verstöße (Persönliche Schutzausrüstung) dokumentieren

---

## `functions/notizen_db.py`

**Datenbankpfad**: `database SQL/notizen.db`

- Allgemeine Notizen mit Datum und Kategorie

---

## `functions/stellungnahmen_db.py`

**Datenbankpfad**: `database SQL/stellungnahmen.db`

- Dienstliche Stellungnahmen erfassen und verwalten

---

## `functions/stellungnahmen_html_export.py`

- HTML-Export von Stellungnahmen
- Briefkopf-Format, DRK-Design

---

## `functions/dienstanweisungen_db.py`

- Dienstanweisungen verwalten und kategorisieren
