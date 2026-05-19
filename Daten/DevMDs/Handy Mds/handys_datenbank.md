# Nesk3 – Datenbank: `handys.db`

## Allgemein

| Eigenschaft | Wert |
|-------------|------|
| Dateiname | `handys.db` |
| Typ | SQLite 3 |
| Journal-Modus | WAL (`PRAGMA journal_mode=WAL;`) |
| Encoding | UTF-8 |
| Speicherort | Gleicher Ordner wie `nesk3.db` (z. B. `Daten/`) |
| Backup | Via `BackupManager`, täglich, gleiche Retention wie Haupt-DB |

---

## Tabelle: `handys`

Speichert alle Diensthandys.

```sql
CREATE TABLE IF NOT EXISTS handys (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    inventarnummer      TEXT    NOT NULL UNIQUE,
    hersteller          TEXT,
    modell              TEXT,
    rufnummer           TEXT,
    sim_nummer          TEXT,
    standort            TEXT,
    zustand             TEXT    NOT NULL DEFAULT 'Aktiv'
                                CHECK(zustand IN ('Aktiv','Defekt','Außer Betrieb','Reserve')),
    defekt_beschreibung TEXT,
    anschaffungsdatum   TEXT,                        -- ISO 8601: YYYY-MM-DD
    notizen             TEXT,
    erstellt_am         TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    geaendert_am        TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
```

### Feld-Beschreibungen

| Feld | Typ | Pflicht | Beschreibung |
|------|-----|---------|-------------|
| `id` | INTEGER | auto | Primärschlüssel |
| `inventarnummer` | TEXT | ✓ | Eindeutige interne Nummer, z. B. `HY-001` |
| `hersteller` | TEXT | – | z. B. `Samsung`, `Apple`, `Nokia` |
| `modell` | TEXT | – | z. B. `Galaxy A55`, `iPhone SE` |
| `rufnummer` | TEXT | – | Dienstliche Mobilnummer |
| `sim_nummer` | TEXT | – | ICCID oder Kurzbezeichnung der SIM |
| `standort` | TEXT | – | Aktueller Standort/Ausgabebereich |
| `zustand` | TEXT | ✓ | Enum (s. u.) |
| `defekt_beschreibung` | TEXT | – | Freitext, relevant wenn `zustand = Defekt` |
| `anschaffungsdatum` | TEXT | – | ISO-Datum der Beschaffung |
| `notizen` | TEXT | – | Sonstige Anmerkungen |
| `erstellt_am` | TEXT | auto | Erstellungszeitpunkt (Lokalzeit) |
| `geaendert_am` | TEXT | auto | Letzter Änderungszeitpunkt |

### Enum: `zustand`

| Wert | Bedeutung |
|------|-----------|
| `Aktiv` | Gerät in Betrieb, einsatzbereit |
| `Defekt` | Gerät defekt, `defekt_beschreibung` sollte befüllt sein |
| `Außer Betrieb` | Temporär nicht im Einsatz (z. B. verloren, gesperrt) |
| `Reserve` | Vorrat / Ersatzgerät |

---

## Tabelle: `handys_historie`

Protokolliert jede Änderung an einem Handy-Datensatz (Audit Trail).

```sql
CREATE TABLE IF NOT EXISTS handys_historie (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    handy_id        INTEGER NOT NULL REFERENCES handys(id) ON DELETE CASCADE,
    inventarnummer  TEXT    NOT NULL,   -- Denormalisiert für lesbare Logs
    feld            TEXT    NOT NULL,   -- Name des geänderten Feldes
    alter_wert      TEXT,
    neuer_wert      TEXT,
    benutzer        TEXT,              -- Windows-Username oder Nesk3-User
    geaendert_am    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
```

### Feld-Beschreibungen

| Feld | Beschreibung |
|------|-------------|
| `handy_id` | FK auf `handys.id` |
| `inventarnummer` | Kopie der Inv.-Nr. für lesbaren Log (auch nach Löschung) |
| `feld` | Name des Datenbankfeldes, das geändert wurde |
| `alter_wert` | Wert vor der Änderung (als Text) |
| `neuer_wert` | Wert nach der Änderung (als Text) |
| `benutzer` | Erfassender Benutzer (z. B. `os.getlogin()`) |
| `geaendert_am` | Zeitstempel der Änderung |

---

## Trigger: Auto-Update `geaendert_am`

```sql
CREATE TRIGGER IF NOT EXISTS trg_handys_geaendert_am
AFTER UPDATE ON handys
FOR EACH ROW
BEGIN
    UPDATE handys SET geaendert_am = datetime('now','localtime')
    WHERE id = OLD.id;
END;
```

---

## Initialisierung (`handys_db.py`)

```python
import sqlite3
import os

HANDYS_DB_PATH = os.path.join(
    os.path.dirname(__file__), '..', '..', 'data', 'handys.db'
)

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(HANDYS_DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_connection() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS handys (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                inventarnummer      TEXT    NOT NULL UNIQUE,
                hersteller          TEXT,
                modell              TEXT,
                rufnummer           TEXT,
                sim_nummer          TEXT,
                standort            TEXT,
                zustand             TEXT NOT NULL DEFAULT 'Aktiv'
                                    CHECK(zustand IN ('Aktiv','Defekt','Außer Betrieb','Reserve')),
                defekt_beschreibung TEXT,
                anschaffungsdatum   TEXT,
                notizen             TEXT,
                erstellt_am         TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                geaendert_am        TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS handys_historie (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                handy_id        INTEGER NOT NULL REFERENCES handys(id) ON DELETE CASCADE,
                inventarnummer  TEXT    NOT NULL,
                feld            TEXT    NOT NULL,
                alter_wert      TEXT,
                neuer_wert      TEXT,
                benutzer        TEXT,
                geaendert_am    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TRIGGER IF NOT EXISTS trg_handys_geaendert_am
            AFTER UPDATE ON handys
            FOR EACH ROW
            BEGIN
                UPDATE handys SET geaendert_am = datetime('now','localtime')
                WHERE id = OLD.id;
            END;
        """)
```

---

## Backup-Integration

In `BackupManager` (bestehende Klasse) folgenden Eintrag ergänzen:

```python
# In der Liste der zu sichernden Datenbanken:
BACKUP_DATABASES = [
    "nesk3.db",
    "handys.db",   # <-- NEU
    # ... weitere DBs
]
```

- Backup-Zielordner: identisch mit bestehenden Backups
- Dateiname im Backup: `handys_YYYY-MM-DD_HHMMSS.db`
- Retention: gleiche Aufbewahrungsregel wie `nesk3.db`

---

## CRUD-Übersicht (`handys_db.py`)

| Funktion | Beschreibung |
|----------|-------------|
| `get_all_handys()` | Alle Handys abrufen (für Übersichtstabelle) |
| `get_handy_by_id(id)` | Einzelnes Gerät laden |
| `insert_handy(data: dict)` | Neues Gerät anlegen |
| `update_handy(id, data: dict, benutzer: str)` | Gerät aktualisieren + Historien-Eintrag |
| `delete_handy(id)` | Gerät löschen (Cascade auf Historie) |
| `get_historie(handy_id=None)` | Alle oder gefilterte Historien-Einträge |

### Beispiel: `update_handy` mit Historien-Eintrag

```python
def update_handy(id: int, data: dict, benutzer: str):
    with get_connection() as conn:
        old = dict(conn.execute("SELECT * FROM handys WHERE id=?", (id,)).fetchone())
        felder_aendern = {k: v for k, v in data.items() if old.get(k) != v}

        if felder_aendern:
            set_clause = ", ".join(f"{k}=?" for k in felder_aendern)
            conn.execute(
                f"UPDATE handys SET {set_clause} WHERE id=?",
                (*felder_aendern.values(), id)
            )
            for feld, neuer_wert in felder_aendern.items():
                conn.execute(
                    """INSERT INTO handys_historie
                       (handy_id, inventarnummer, feld, alter_wert, neuer_wert, benutzer)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (id, old['inventarnummer'], feld, str(old.get(feld)), str(neuer_wert), benutzer)
                )
```
