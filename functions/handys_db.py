"""
Handys-Datenbank
CRUD-Funktionen für Diensthandys und Historien-Log
"""
import sqlite3
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import HANDYS_DB_PATH

ZUSTAND_OPTIONEN = ["Aktiv", "Defekt", "Außer Betrieb", "Reserve", "Verloren"]

# Zustände bei denen "Festgestellt am / von" angezeigt wird
ZUSTAND_MIT_DEFEKT_DETAIL = {"Defekt", "Außer Betrieb", "Verloren"}

_SQL_SCHEMA = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

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

CREATE TRIGGER IF NOT EXISTS trg_handys_geaendert_am
AFTER UPDATE ON handys
FOR EACH ROW
BEGIN
    UPDATE handys SET geaendert_am = datetime('now','localtime')
    WHERE id = OLD.id;
END;
"""


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(HANDYS_DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.row_factory = sqlite3.Row
    return conn


def init_handys_db():
    """Erstellt alle Tabellen und Trigger; migriert bestehende Schemas."""
    os.makedirs(os.path.dirname(HANDYS_DB_PATH), exist_ok=True)
    conn = get_connection()
    try:
        # 1. Tabellen / Trigger erstellen (falls noch nicht vorhanden)
        conn.executescript(_SQL_SCHEMA)

        # 2. Neue Spalten sicher hinzufügen (ALTER TABLE ignoriert Fehler wenn Spalte existiert)
        existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(handys)")}
        for col, typedef in [
            ("defekt_datum",        "TEXT DEFAULT ''"),
            ("defekt_gemeldet_von", "TEXT DEFAULT ''"),
            ("kartennummer",        "TEXT DEFAULT ''"),
            ("pin",                 "TEXT DEFAULT ''"),
            ("pin2",                "TEXT DEFAULT ''"),
            ("puk",                 "TEXT DEFAULT ''"),
            ("puk2",                "TEXT DEFAULT ''"),
        ]:
            if col not in existing_cols:
                conn.execute(f"ALTER TABLE handys ADD COLUMN {col} {typedef}")
                conn.commit()

        # 3. CHECK-Constraint prüfen – falls 'Verloren' fehlt, Tabelle rebuild
        schema_row = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='handys'"
        ).fetchone()
        if schema_row and "Verloren" not in (schema_row[0] or ""):
            data = [dict(r) for r in conn.execute("SELECT * FROM handys")]
            conn.executescript("""
                DROP TRIGGER IF EXISTS trg_handys_geaendert_am;
                ALTER TABLE handys RENAME TO handys_v_old;
            """)
            conn.executescript(_SQL_SCHEMA)
            if data:
                new_cols = [row[1] for row in conn.execute("PRAGMA table_info(handys)")]
                for row in data:
                    common = {k: v for k, v in row.items() if k in new_cols and k != "id"}
                    cols = list(common.keys())
                    vals = list(common.values())
                    conn.execute(
                        f"INSERT INTO handys ({', '.join(cols)}) "
                        f"VALUES ({', '.join(['?' for _ in cols])})",
                        vals,
                    )
                conn.commit()
            conn.execute("DROP TABLE IF EXISTS handys_v_old")
            conn.commit()
    finally:
        conn.close()
    print("[OK] Handys-DB initialisiert.")


def _benutzer() -> str:
    try:
        return os.getlogin()
    except Exception:
        return "System"


# ─── Handys CRUD ─────────────────────────────────────────────────────────────

def lade_alle_handys() -> list[dict]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM handys ORDER BY inventarnummer"
        ).fetchall()
    return [dict(r) for r in rows]


def lade_handy(handy_id: int) -> dict | None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM handys WHERE id = ?", (handy_id,)
        ).fetchone()
    return dict(row) if row else None


def erstelle_handy(daten: dict) -> int:
    """
    Legt ein neues Handy an.
    Gibt die neue ID zurück.
    Pushed anschließend nach Turso.
    """
    felder = [
        "inventarnummer", "hersteller", "modell", "rufnummer",
        "sim_nummer", "standort", "zustand", "defekt_beschreibung",
        "defekt_datum", "defekt_gemeldet_von",
        "anschaffungsdatum", "notizen",
        "kartennummer", "pin", "pin2", "puk", "puk2",
    ]
    cols = ", ".join(felder)
    placeholders = ", ".join(["?" for _ in felder])
    werte = [daten.get(f, "") for f in felder]
    with get_connection() as conn:
        cur = conn.execute(
            f"INSERT INTO handys ({cols}) VALUES ({placeholders})", werte
        )
        new_id = cur.lastrowid
        row = dict(conn.execute("SELECT * FROM handys WHERE id = ?", (new_id,)).fetchone())

    _push_handy(row)

    # Erstell-Eintrag in Historie
    _schreibe_historie(new_id, daten.get("inventarnummer", ""), "erstellt", "", "neu")

    return new_id


def aktualisiere_handy(handy_id: int, neu: dict) -> bool:
    """
    Aktualisiert einen Handy-Datensatz und schreibt geänderte Felder in die Historie.
    """
    alt = lade_handy(handy_id)
    if not alt:
        return False

    felder = [
        "inventarnummer", "hersteller", "modell", "rufnummer",
        "sim_nummer", "standort", "zustand", "defekt_beschreibung",
        "defekt_datum", "defekt_gemeldet_von",
        "anschaffungsdatum", "notizen",
        "kartennummer", "pin", "pin2", "puk", "puk2",
    ]
    set_teile = ", ".join([f"{f} = ?" for f in felder])
    werte = [neu.get(f, alt.get(f, "")) for f in felder]
    werte.append(handy_id)

    with get_connection() as conn:
        conn.execute(
            f"UPDATE handys SET {set_teile} WHERE id = ?", werte
        )
        row = dict(conn.execute("SELECT * FROM handys WHERE id = ?", (handy_id,)).fetchone())

    _push_handy(row)

    # Geänderte Felder in Historie schreiben
    inv = neu.get("inventarnummer", alt.get("inventarnummer", ""))
    for f in felder:
        alter_val = str(alt.get(f, "") or "")
        neuer_val = str(neu.get(f, alt.get(f, "")) or "")
        if alter_val != neuer_val:
            _schreibe_historie(handy_id, inv, f, alter_val, neuer_val)

    return True


def loesche_handy(handy_id: int) -> bool:
    handy = lade_handy(handy_id)
    if not handy:
        return False

    with get_connection() as conn:
        conn.execute("DELETE FROM handys_historie WHERE handy_id = ?", (handy_id,))
        conn.execute("DELETE FROM handys WHERE id = ?", (handy_id,))

    _push_delete_handy(handy_id)
    return True


# ─── Historie ─────────────────────────────────────────────────────────────────

def _schreibe_historie(handy_id: int, inventarnummer: str,
                       feld: str, alter_wert: str, neuer_wert: str):
    benutzer = _benutzer()
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO handys_historie "
            "(handy_id, inventarnummer, feld, alter_wert, neuer_wert, benutzer) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (handy_id, inventarnummer, feld, alter_wert, neuer_wert, benutzer),
        )
        row = dict(conn.execute(
            "SELECT * FROM handys_historie WHERE id = last_insert_rowid()"
        ).fetchone())

    _push_historie(row)


def lade_historie(handy_id: int | None = None, tage: int = 90) -> list[dict]:
    """
    Lädt Historien-Einträge.
    handy_id=None → alle Geräte; tage=0 → gesamte Geschichte.
    """
    with get_connection() as conn:
        if tage > 0:
            datum_filter = f"AND h.geaendert_am >= datetime('now','localtime','-{tage} days')"
        else:
            datum_filter = ""

        if handy_id is not None:
            rows = conn.execute(
                f"SELECT h.* FROM handys_historie h "
                f"WHERE h.handy_id = ? {datum_filter} "
                f"ORDER BY h.geaendert_am DESC",
                (handy_id,),
            ).fetchall()
        else:
            rows = conn.execute(
                f"SELECT h.* FROM handys_historie h "
                f"WHERE 1=1 {datum_filter} "
                f"ORDER BY h.geaendert_am DESC"
            ).fetchall()
    return [dict(r) for r in rows]


# ─── Turso-Push ───────────────────────────────────────────────────────────────

def _push_handy(row: dict):
    try:
        from database.turso_sync import push_row
        push_row(HANDYS_DB_PATH, "handys", row)
    except Exception as e:
        print(f"[Turso] Handys-Push fehlgeschlagen: {e}")


def _push_delete_handy(handy_id: int):
    try:
        from database.turso_sync import push_delete
        push_delete(HANDYS_DB_PATH, "handys", handy_id)
    except Exception as e:
        print(f"[Turso] Handys-Delete fehlgeschlagen: {e}")


def _push_historie(row: dict):
    try:
        from database.turso_sync import push_row
        push_row(HANDYS_DB_PATH, "handys_historie", row)
    except Exception as e:
        print(f"[Turso] Handys-Historie-Push fehlgeschlagen: {e}")
