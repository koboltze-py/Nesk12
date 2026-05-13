"""SQLite-CRUD für persönliche Dashboard-Notizen."""
import sqlite3
from datetime import datetime, timedelta


def _db_path() -> str:
    from config import NOTIZEN_DB_PATH
    return NOTIZEN_DB_PATH


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _init_db():
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS notizen (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                titel       TEXT    NOT NULL,
                text        TEXT    NOT NULL DEFAULT '',
                datum       TEXT    NOT NULL,
                erstellt_am TEXT    NOT NULL,
                status      TEXT    NOT NULL DEFAULT 'offen'
            )
        """)
        conn.commit()


def speichern(titel: str, text: str = "", datum: str = "") -> int:
    """Neue Notiz anlegen. datum im Format dd.MM.yyyy."""
    _init_db()
    if not datum:
        datum = datetime.today().strftime("%d.%m.%Y")
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO notizen (titel, text, datum, erstellt_am, status) "
            "VALUES (?, ?, ?, ?, 'offen')",
            (titel, text, datum, jetzt),
        )
        conn.commit()
        return cur.lastrowid


def als_gelesen(nid: int):
    """Status auf 'gelesen' setzen."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("UPDATE notizen SET status='gelesen' WHERE id=?", (nid,))
        conn.commit()


def als_erledigt(nid: int):
    """Status auf 'erledigt' setzen (Notiz verschwindet nach 5 Tagen nicht mehr)."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("UPDATE notizen SET status='erledigt' WHERE id=?", (nid,))
        conn.commit()


def loeschen(nid: int):
    """Notiz dauerhaft löschen."""
    _init_db()
    with _get_conn() as conn:
        conn.execute("DELETE FROM notizen WHERE id=?", (nid,))
        conn.commit()


def lade_aktive() -> list[dict]:
    """
    Notizen der letzten 5 Tage (nach erstellt_am).
    Erledigte werden ebenfalls angezeigt (als durchgestrichen / grau).
    """
    _init_db()
    grenze = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen WHERE erstellt_am >= ? ORDER BY erstellt_am DESC",
            (grenze,),
        ).fetchall()
    return [dict(r) for r in rows]


def lade_fenster() -> list[dict]:
    """
    Notizen deren `datum` im Fenster [heute-5 Tage … heute+10 Tage] liegt.
    Sortiert nach datum aufsteigend (Vergangenheit → Zukunft).
    """
    _init_db()
    heute = datetime.today().date()
    von   = heute - timedelta(days=5)
    bis   = heute + timedelta(days=10)
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen ORDER BY datum ASC, erstellt_am ASC"
        ).fetchall()
    ergebnis = []
    for r in rows:
        row = dict(r)
        try:
            d = datetime.strptime(row["datum"], "%d.%m.%Y").date()
            if von <= d <= bis:
                ergebnis.append(row)
        except ValueError:
            pass
    return ergebnis


def lade_zukunft() -> list[dict]:
    """
    Notizen der nächsten 10 Tage (ab morgen), anhand des Feldes `datum` (dd.MM.yyyy).
    Gibt alle Status zurück, sortiert nach datum aufsteigend.
    """
    _init_db()
    heute = datetime.today()
    morgen = heute + timedelta(days=1)
    in_10 = heute + timedelta(days=10)
    # Alle Notizen laden und im Python filtern (datum ist kein ISO-Format)
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen ORDER BY datum ASC, erstellt_am ASC"
        ).fetchall()
    ergebnis = []
    for r in rows:
        row = dict(r)
        try:
            d = datetime.strptime(row["datum"], "%d.%m.%Y")
            if morgen.date() <= d.date() <= in_10.date():
                ergebnis.append(row)
        except ValueError:
            pass
    return ergebnis


def lade_alle() -> list[dict]:
    """Alle Notizen, neueste zuerst."""
    _init_db()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen ORDER BY erstellt_am DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def lade_fuer_datum(datum_de: str) -> list[dict]:
    """Alle Notizen für ein bestimmtes Datum (Format dd.MM.yyyy)."""
    _init_db()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM notizen WHERE datum=? ORDER BY erstellt_am DESC",
            (datum_de,),
        ).fetchall()
    return [dict(r) for r in rows]
