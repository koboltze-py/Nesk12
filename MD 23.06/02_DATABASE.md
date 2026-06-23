# Datenbank – Schema & Verbindungen

## Übersicht Datenbankdateien

| Datei | Speicherort | Turso-Prefix | Inhalt |
|---|---|---|---|
| `nesk3.db` | `database SQL/` | `nesk3__` | Hauptdatenbank |
| `mitarbeiter.db` | `database SQL/` | `ma__` | Mitarbeiterstammdaten |
| `einsaetze.db` | `database SQL/` | `einsaetze__` | Einsätze |
| `verspaetungen.db` | `database SQL/` | `vers__` | Verspätungsdaten |
| `telefonnummern.db` | `database SQL/` | `tel__` | Telefonverzeichnis |
| `patienten_station.db` | `database SQL/` | `pat__` | Patienten-/Medikamentendaten |
| `call_transcription.db` | `database SQL/` | `call__` | Anruf-Protokolle |
| `psa.db` | `database SQL/` | `psa__` | PSA-Verstöße |
| `stellungnahmen.db` | `database SQL/` | `stelg__` | Stellungnahmen |
| `beschwerden.db` | `database SQL/` | `bschw__` | Beschwerden |
| `vorkommnisse.db` | `database SQL/` | `vork__` | Vorkommnisse |
| `handys.db` | `database SQL/` | `handys__` | Handy-Verwaltung |
| `workflow.db` | `database SQL/` | _(kein)_ | Workflow-Daten |
| `sanmat.db` | `database SQL/` | _(kein)_ | SANMAT-Daten |
| `notizen.db` | `database SQL/` | _(kein)_ | Notizen |

---

## Verbindungsmanagement (`database/connection.py`)

### Funktion: `get_connection()`

```python
def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=3, check_same_thread=False)
    conn.row_factory = _row_factory     # Zeilen als dict
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn
```

- **WAL-Modus**: Mehrere Leser + ein Schreiber gleichzeitig (für Multi-PC via OneDrive)
- **Row-Factory**: Alle Zeilen werden als `dict` zurückgegeben (`{spaltenname: wert}`)
- **Timeout**: 5 Sekunden Wartezeit bei gesperrter DB

### Kontextmanager: `db_cursor(commit=False)`

```python
with db_cursor(commit=True) as cur:
    cur.execute("INSERT INTO tabelle ...")
```

Öffnet eine Verbindung, liefert Cursor, commit/rollback automatisch.

### Funktion: `get_ma_connection()`

Wie `get_connection()`, aber für `mitarbeiter.db`.

---

## Schema – nesk3.db (Hauptdatenbank)

### `abteilungen`
```sql
CREATE TABLE IF NOT EXISTS abteilungen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,
    beschreibung    TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `positionen`
```sql
CREATE TABLE IF NOT EXISTS positionen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,
    kuerzel         TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

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
CREATE INDEX IF NOT EXISTS idx_ma_nachname ON mitarbeiter(nachname);
CREATE INDEX IF NOT EXISTS idx_ma_funktion ON mitarbeiter(funktion);
```

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
```

### `uebergabe_fahrzeug_notizen`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_fahrzeug_notizen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    fahrzeug_id     INTEGER,
    kennzeichen     TEXT DEFAULT '',
    notiz           TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `uebergabe_handy_eintraege`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_handy_eintraege (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    handy_name      TEXT DEFAULT '',
    empfaenger      TEXT DEFAULT '',
    status          TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `uebergabe_verspaetungen`
```sql
CREATE TABLE IF NOT EXISTS uebergabe_verspaetungen (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    protokoll_id    INTEGER REFERENCES uebergabe_protokolle(id) ON DELETE CASCADE,
    mitarbeiter     TEXT DEFAULT '',
    dienstbeginn    TEXT DEFAULT '',
    verspaetung_min INTEGER DEFAULT 0,
    grund           TEXT DEFAULT '',
    datum           TEXT DEFAULT '',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `fahrzeuge`
```sql
CREATE TABLE IF NOT EXISTS fahrzeuge (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    kennzeichen     TEXT NOT NULL UNIQUE,
    typ             TEXT DEFAULT '',      -- RTW, KTW, PKW ...
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
    gesendet        INTEGER DEFAULT 0,
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `fahrzeug_termine`
```sql
CREATE TABLE IF NOT EXISTS fahrzeug_termine (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    fahrzeug_id     INTEGER NOT NULL REFERENCES fahrzeuge(id) ON DELETE CASCADE,
    termin_typ      TEXT DEFAULT '',      -- TÜV, Service, HU ...
    faellig_am      TEXT NOT NULL,
    notiz           TEXT DEFAULT '',
    erledigt        INTEGER DEFAULT 0,
    erstellt_am     TEXT DEFAULT (datetime('now','localtime'))
);
```

### `settings`
```sql
CREATE TABLE IF NOT EXISTS settings (
    schluessel  TEXT PRIMARY KEY,
    wert        TEXT NOT NULL DEFAULT ''
);
```

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

## Schema – mitarbeiter.db

```sql
CREATE TABLE IF NOT EXISTS positionen (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    name  TEXT UNIQUE NOT NULL
);
INSERT OR IGNORE INTO positionen(name) VALUES
    ('Notfallsanitäter'),('Rettungssanitäter'),('Sanitätshelfer'),
    ('Arzt'),('Verwaltung'),('Führungskraft');

CREATE TABLE IF NOT EXISTS abteilungen (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    name  TEXT UNIQUE NOT NULL
);
INSERT OR IGNORE INTO abteilungen(name) VALUES
    ('Erste-Hilfe-Station'),('Sanitätsdienst'),('Verwaltung');

CREATE TABLE IF NOT EXISTS mitarbeiter (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    vorname         TEXT NOT NULL,
    nachname        TEXT NOT NULL,
    personalnummer  TEXT DEFAULT '',
    funktion        TEXT DEFAULT 'Schichtleiter'
                    CHECK (funktion IN ('Schichtleiter','Dispo','Betreuer')),
    position        TEXT DEFAULT '',
    abteilung       TEXT DEFAULT '',
    email           TEXT DEFAULT '',
    telefon         TEXT DEFAULT '',
    eintrittsdatum  TEXT,
    status          TEXT DEFAULT 'aktiv',
    erstellt_am     TEXT DEFAULT (datetime('now','localtime')),
    geaendert_am    TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_ma_nachname ON mitarbeiter(nachname);
CREATE INDEX IF NOT EXISTS idx_ma_funktion ON mitarbeiter(funktion);
```

---

## Datenmodelle (`database/models.py`)

### `Mitarbeiter` (dataclass)
```python
@dataclass
class Mitarbeiter:
    id:             Optional[int]  = None
    vorname:        str            = ""
    nachname:       str            = ""
    personalnummer: str            = ""
    funktion:       str            = "stamm"  # stamm | dispo
    position:       str            = ""
    abteilung:      str            = ""
    email:          str            = ""
    telefon:        str            = ""
    eintrittsdatum: Optional[date] = None
    status:         str            = "aktiv"  # aktiv | inaktiv | beurlaubt
    erstellt_am:    Optional[datetime] = None
    geaendert_am:   Optional[datetime] = None

    @property
    def vollname(self) -> str:
        return f"{self.vorname} {self.nachname}".strip()
```

### `Dienstplan` (dataclass)
```python
@dataclass
class Dienstplan:
    id:               Optional[int]  = None
    mitarbeiter_id:   Optional[int]  = None
    mitarbeiter_name: str            = ""
    datum:            Optional[date] = None
    start_uhrzeit:    Optional[time] = None
    end_uhrzeit:      Optional[time] = None
    position:         str            = ""
    schicht_typ:      str            = "regulär"  # regulär | nacht | bereitschaft
    notizen:          str            = ""
    erstellt_am:      Optional[datetime] = None
```

### `UebergabeProtokoll` (dataclass)
```python
@dataclass
class UebergabeProtokoll:
    id:               Optional[int]  = None
    datum:            str            = ""
    schicht_typ:      str            = "tagdienst"  # tagdienst | nachtdienst
    beginn_zeit:      str            = ""
    ende_zeit:        str            = ""
    patienten_anzahl: int            = 0
    personal:         str            = ""
    ereignisse:       str            = ""
    massnahmen:       str            = ""
    uebergabe_notiz:  str            = ""
    ersteller:        str            = ""
    abzeichner:       str            = ""
    status:           str            = "offen"  # offen | abgeschlossen
```

### `Fahrzeug` (dataclass)
```python
@dataclass
class Fahrzeug:
    id:           Optional[int] = None
    kennzeichen:  str           = ""
    typ:          str           = ""  # RTW, KTW, PKW ...
    marke:        str           = ""
    modell:       str           = ""
    baujahr:      Optional[int] = None
    fahrgestellnr: str          = ""
    tuev_datum:   str           = ""  # YYYY-MM-DD
    notizen:      str           = ""
    aktiv:        int           = 1
```

---

## Backup-System (DB-Ebene)

Bei jedem App-Start (`main.py`) werden automatische DB-Backups erstellt:

```
database SQL/
└── db_backups/
    └── YYYY-MM-DD/          ← Tages-Ordner (max. 7 Tage)
        ├── nesk3.db          (bis zu 5 Kopien pro Tag)
        ├── mitarbeiter.db
        └── ...
```

**Limits:**
- Max. 5 Backups pro Datenbank pro Tag
- Max. 7 Tages-Ordner (älteste werden beim Start gelöscht)
