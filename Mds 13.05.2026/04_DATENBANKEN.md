# Datenbanken – vollständige Schemata (Stand: 02.06.2026)

## Übersicht

| DB-Datei | Tabellen | Erstellt durch |
|---|---|---|
| `nesk3.db` | 18 Tabellen | `database/migrations.py` |
| `notizen.db` | `notizen` | `functions/notizen_db.py` |
| `vorkommnisse.db` | `vorkommnisse` | `functions/vorkommnisse_db.py` |
| `beschwerden.db` | `beschwerden`, `massnahmen`, `stellungnahmen` | `functions/beschwerden_db.py` |
| `handys.db` | `handys`, `handys_historie` | `functions/handys_db.py` |
| `sanmat.db` | SanMat-Einträge | `database/sanmat_db.py` |
| `archiv.db` | Archiv-Einträge | `functions/archiv_functions.py` |
| `mitarbeiter.db` | (sync) | `functions/mitarbeiter_sync.py` |

---

## 1. nesk3.db – Haupt-Datenbank

### `abteilungen`
```sql
CREATE TABLE IF NOT EXISTS abteilungen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,
    beschreibung    TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```
**Defaults:** `Erste-Hilfe-Station`, `Sanitaetsdienst`, `Verwaltung`

---

### `positionen`
```sql
CREATE TABLE IF NOT EXISTS positionen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,
    kuerzel         TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```
**Defaults:** RS, RA, NFS, SH, SL (Schichtleiter), EL (Einsatzleiter)

---

### `mitarbeiter`
```sql
CREATE TABLE IF NOT EXISTS mitarbeiter (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    vorname         TEXT NOT NULL,
    nachname        TEXT NOT NULL,
    personalnummer  TEXT UNIQUE,
    funktion        TEXT DEFAULT 'Schichtleiter'
                    CHECK (funktion IN ('Schichtleiter','Dispo','Betreuer')),
    position        TEXT DEFAULT '',
    abteilung       TEXT DEFAULT '',
    email           TEXT DEFAULT '',
    telefon         TEXT DEFAULT '',
    eintrittsdatum  TEXT,
    status          TEXT DEFAULT 'aktiv'
                    CHECK (status IN ('aktiv','inaktiv','beurlaubt')),
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
-- Nachgerüstet via ALTER TABLE:
-- funktion        TEXT DEFAULT 'stamm'
```

---

### `dienstplan`
```sql
CREATE TABLE IF NOT EXISTS dienstplan (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    mitarbeiter_id  INTEGER REFERENCES mitarbeiter(id) ON DELETE SET NULL,
    datum           TEXT NOT NULL,
    start_uhrzeit   TEXT NOT NULL,
    end_uhrzeit     TEXT NOT NULL,
    position        TEXT DEFAULT '',
    schicht_typ     TEXT DEFAULT 'regulaer'
                    CHECK (schicht_typ IN ('regulaer','nacht','bereitschaft')),
    notizen         TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `backup_log`
```sql
CREATE TABLE IF NOT EXISTS backup_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    dateiname       TEXT NOT NULL,
    typ             TEXT DEFAULT 'manuell',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `settings`
```sql
CREATE TABLE IF NOT EXISTS settings (
    schluessel  TEXT PRIMARY KEY,
    wert        TEXT NOT NULL DEFAULT ''
);
```
**Bekannte Keys:** `dienstplan_ordner`

---

### `uebergabe_protokolle`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_protokolle (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    datum               TEXT NOT NULL,
    schicht_typ         TEXT NOT NULL
                        CHECK (schicht_typ IN ('tagdienst','nachtdienst')),
    beginn_zeit         TEXT DEFAULT '',
    ende_zeit           TEXT DEFAULT '',
    patienten_anzahl    INTEGER DEFAULT 0,
    personal            TEXT DEFAULT '',
    ereignisse          TEXT DEFAULT '',
    massnahmen          TEXT DEFAULT '',
    uebergabe_notiz     TEXT DEFAULT '',
    ersteller           TEXT DEFAULT '',
    abzeichner          TEXT DEFAULT '',
    status              TEXT DEFAULT 'offen'
                        CHECK (status IN ('offen','abgeschlossen')),
    erstellt_am         TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am        TEXT DEFAULT (datetime('now','localtime'))
);
-- Nachgerüstet:
-- handys_anzahl    INTEGER DEFAULT 0
-- handys_notiz     TEXT DEFAULT ''
-- archiviert       INTEGER DEFAULT 0
```

---

### `fahrzeuge`
```sql
CREATE TABLE IF NOT EXISTS fahrzeuge (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    kennzeichen     TEXT NOT NULL UNIQUE,
    typ             TEXT DEFAULT '',
    marke           TEXT DEFAULT '',
    modell          TEXT DEFAULT '',
    baujahr         INTEGER,
    fahrgestellnr   TEXT DEFAULT '',
    tuev_datum      TEXT DEFAULT '',
    notizen         TEXT DEFAULT '',
    aktiv           INTEGER DEFAULT 1,
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `fahrzeug_status`
```sql
CREATE TABLE IF NOT EXISTS fahrzeug_status (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fahrzeug_id     INTEGER NOT NULL REFERENCES fahrzeuge(id) ON DELETE CASCADE,
    status          TEXT NOT NULL
                    CHECK (status IN ('fahrbereit','defekt','werkstatt','ausser_dienst','sonstiges')),
    von             TEXT NOT NULL,
    bis             TEXT DEFAULT '',
    grund           TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `fahrzeug_schaeden`
```sql
CREATE TABLE IF NOT EXISTS fahrzeug_schaeden (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fahrzeug_id     INTEGER NOT NULL REFERENCES fahrzeuge(id) ON DELETE CASCADE,
    datum           TEXT NOT NULL,
    beschreibung    TEXT NOT NULL,
    schwere         TEXT DEFAULT 'gering'
                    CHECK (schwere IN ('gering','mittel','schwer')),
    kommentar       TEXT DEFAULT '',
    behoben         INTEGER DEFAULT 0,
    behoben_am      TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
-- Nachgerüstet:
-- gesendet         INTEGER DEFAULT 0  (E-Mail-Tracking)
```

---

### `fahrzeug_termine`
```sql
CREATE TABLE IF NOT EXISTS fahrzeug_termine (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fahrzeug_id     INTEGER NOT NULL REFERENCES fahrzeuge(id) ON DELETE CASCADE,
    datum           TEXT NOT NULL,
    uhrzeit         TEXT DEFAULT '',
    typ             TEXT DEFAULT 'sonstiges'
                    CHECK (typ IN ('tuev','inspektion','reparatur','hauptuntersuchung','sonstiges')),
    titel           TEXT NOT NULL,
    beschreibung    TEXT DEFAULT '',
    kommentar       TEXT DEFAULT '',
    erledigt        INTEGER DEFAULT 0,
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `uebergabe_fahrzeug_notizen`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_fahrzeug_notizen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER NOT NULL REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    fahrzeug_id     INTEGER NOT NULL REFERENCES fahrzeuge(id) ON DELETE CASCADE,
    notiz           TEXT DEFAULT '',
    UNIQUE(protokoll_id, fahrzeug_id)
);
```

---

### `uebergabe_handy_eintraege`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_handy_eintraege (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER NOT NULL REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    geraet_nr       TEXT NOT NULL,
    notiz           TEXT DEFAULT ''
);
```

---

### `uebergabe_verspaetungen`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_verspaetungen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER NOT NULL REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    mitarbeiter     TEXT NOT NULL,
    soll_zeit       TEXT DEFAULT '',
    ist_zeit        TEXT DEFAULT ''
);
```

---

### `drk_daten_backup_log`
```sql
CREATE TABLE IF NOT EXISTS drk_daten_backup_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    dateiname       TEXT NOT NULL,
    pfad_nesk       TEXT DEFAULT '',
    pfad_lokal      TEXT DEFAULT '',
    groesse_mb      REAL DEFAULT 0,
    gesicherte_ordner INTEGER DEFAULT 0,
    fehler_ordner   INTEGER DEFAULT 0,
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `tages_pax`
```sql
CREATE TABLE IF NOT EXISTS tages_pax (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    datum       TEXT NOT NULL UNIQUE,
    pax_zahl    INTEGER NOT NULL DEFAULT 0,
    erfasst_am  TEXT DEFAULT (datetime('now','localtime'))
);
```

---

### `tages_einsaetze`
```sql
CREATE TABLE IF NOT EXISTS tages_einsaetze (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    datum           TEXT NOT NULL UNIQUE,
    einsaetze_zahl  INTEGER NOT NULL DEFAULT 0,
    erfasst_am      TEXT DEFAULT (datetime('now','localtime'))
);
```

---

## 2. notizen.db

```sql
CREATE TABLE IF NOT EXISTS notizen (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    titel       TEXT    NOT NULL,
    text        TEXT    NOT NULL DEFAULT '',
    datum       TEXT    NOT NULL,      -- Format: dd.MM.yyyy
    erstellt_am TEXT    NOT NULL,      -- Format: YYYY-MM-DD HH:MM:SS
    status      TEXT    NOT NULL DEFAULT 'offen'
                -- Werte: 'offen', 'gelesen', 'erledigt'
);
```

**CRUD-Funktionen (functions/notizen_db.py):**
- `speichern(titel, text, datum)` → `int` (neue ID)
- `als_gelesen(nid)` → Status `'gelesen'`
- `als_erledigt(nid)` → Status `'erledigt'`
- `als_offen(nid)` → Status zurück auf `'offen'`
- `loeschen(nid)`
- `lade_aktive()` → letzte 5 Tage nach `erstellt_am`
- `lade_fenster()` → heute-5 bis heute+10 nach `datum`-Feld, aufsteigend
- `lade_zukunft()` → morgen bis heute+10
- `lade_alle()` → alle, neueste zuerst

---

## 3. vorkommnisse.db

```sql
CREATE TABLE IF NOT EXISTS vorkommnisse (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    flug             TEXT    NOT NULL DEFAULT '',
    typ              TEXT    NOT NULL DEFAULT '',
    datum            TEXT    NOT NULL DEFAULT '',
    ort              TEXT    NOT NULL DEFAULT '',
    offblock_plan    TEXT    NOT NULL DEFAULT '',
    offblock_ist     TEXT    NOT NULL DEFAULT '',
    erstellt_von     TEXT    NOT NULL DEFAULT '',
    ursache          TEXT    NOT NULL DEFAULT '',
    ergebnis         TEXT    NOT NULL DEFAULT '',
    passagiere_json  TEXT    NOT NULL DEFAULT '[]',   -- JSON-Array
    personal_json    TEXT    NOT NULL DEFAULT '[]',   -- JSON-Array
    chronologie_json TEXT    NOT NULL DEFAULT '[]',   -- JSON-Array
    erstellt_am      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    geaendert_am     TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
```

---

## 4. handys.db

```sql
CREATE TABLE IF NOT EXISTS handys (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    inventarnummer      TEXT    NOT NULL UNIQUE,
    hersteller          TEXT    DEFAULT '',
    modell              TEXT    DEFAULT '',
    rufnummer           TEXT    DEFAULT '',
    sim_nummer          TEXT    DEFAULT '',
    standort            TEXT    DEFAULT '',
    zustand             TEXT    NOT NULL DEFAULT 'Aktiv'
                        CHECK(zustand IN ('Aktiv','Defekt','Außer Betrieb','Reserve','Verloren')),
    defekt_beschreibung TEXT    DEFAULT '',
    defekt_datum        TEXT    DEFAULT '',
    defekt_gemeldet_von TEXT    DEFAULT '',
    anschaffungsdatum   TEXT    DEFAULT '',
    notizen             TEXT    DEFAULT '',
    kartennummer        TEXT    DEFAULT '',
    pin                 TEXT    DEFAULT '',
    pin2                TEXT    DEFAULT '',
    puk                 TEXT    DEFAULT '',
    puk2                TEXT    DEFAULT '',
    erstellt_am         TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    geaendert_am        TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS handys_historie (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    handy_id        INTEGER NOT NULL,
    inventarnummer  TEXT    NOT NULL,
    feld            TEXT    NOT NULL,
    alter_wert      TEXT    DEFAULT '',
    neuer_wert      TEXT    DEFAULT '',
    benutzer        TEXT    DEFAULT '',
    geaendert_am    TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);

-- Trigger: aktualisiert geaendert_am automatisch bei UPDATE
CREATE TRIGGER IF NOT EXISTS trg_handys_geaendert_am
AFTER UPDATE ON handys
FOR EACH ROW
BEGIN
    UPDATE handys SET geaendert_am = datetime('now','localtime') WHERE id = NEW.id;
END;
```

**Zustände:** `Aktiv`, `Defekt`, `Außer Betrieb`, `Reserve`, `Verloren`  
**Export-Pfad:** `Daten/Handys/` (konfiguriert in `HANDYS_EXPORT_PATH`)

---

## 5. beschwerden.db

**Kategorien:** Verhalten Mitarbeiter, Wartezeit/Reaktionszeit, Kommunikation, Ausstattung/Material, Hygiene, Dokumentation, Behandlung/Versorgung, Sonstiges  
**Prioritäten:** Niedrig, Mittel, Hoch, Kritisch  
**Status:** Offen, In Bearbeitung, Erledigt, Abgewiesen  
**Quellen:** Freitext, Word-Datei, PDF-Datei, E-Mail

---

## 6. SQLite-Konfiguration

Alle Verbindungen verwenden:
```python
PRAGMA journal_mode = WAL     # Write-Ahead-Logging für bessere Concurrent-Reads
PRAGMA synchronous  = NORMAL  # Kompromiss Performance/Sicherheit
PRAGMA busy_timeout  = 5000   # 5s warten wenn DB gesperrt
PRAGMA foreign_keys = ON      # FK-Constraints aktiviert
```

---

## 7. Datums-Formate

| Kontext | Format |
|---|---|
| `notizen.datum` | `dd.MM.yyyy` (z.B. `"13.05.2026"`) |
| `notizen.erstellt_am` | `YYYY-MM-DD HH:MM:SS` |
| SQLite DEFAULT | `datetime('now','localtime')` |
| Dienstplan | ISO-8601 Text (`YYYY-MM-DD`) |

---

## 8. workflow.db  *(neu v3.8.0)*

**Pfad:** `database SQL/workflow.db`  
**Modul:** `database/workflow_db.py`

### Tabelle `workflow_tag`
```sql
CREATE TABLE workflow_tag (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    datum                 TEXT    NOT NULL,           -- YYYY-MM-DD
    sm_datei              TEXT    NOT NULL,           -- Dateiname Stärkemeldung
    dp_datei              TEXT    NOT NULL DEFAULT '',
    abgeglichen_carmen    INTEGER NOT NULL DEFAULT 0, -- 0/1
    abgeglichen_carmen_am TEXT,                       -- YYYY-MM-DD HH:MM:SS
    notiz                 TEXT    NOT NULL DEFAULT '',
    geaendert_am          TEXT,
    UNIQUE(datum, sm_datei)
);
```

### Tabelle `workflow_session`
```sql
CREATE TABLE workflow_session (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    monat     TEXT    NOT NULL UNIQUE, -- YYYY-MM
    sm_pfade  TEXT    NOT NULL DEFAULT '[]', -- JSON-Array Pfade
    dp_pfade  TEXT    NOT NULL DEFAULT '[]', -- JSON-Array Pfade
    last_used TEXT                           -- YYYY-MM-DD HH:MM:SS
);
```

**Wichtig:** `workflow_tag` persistiert Carmen-Status und Notizen auch wenn der Monat aus `workflow_session` gelöscht wird.
