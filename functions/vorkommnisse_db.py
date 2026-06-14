"""
Vorkommnisse-Datenbank
Eigene SQLite-Datenbank (vorkommnisse.db) – folgt dem Muster beschwerden_db.py

Backup : automatisch via Backup-Manager (glob *.db im DB-Ordner)
Turso  : über push_row() nach jedem Write (TABLE_MAP in turso_sync.py)

Datenmodell (alles in einer Tabelle, Sub-Daten als JSON):
  vorkommnisse  – ein Datensatz pro Vorkommnis
      id, flug, typ, datum, ort, offblock_plan, offblock_ist,
      erstellt_von, ursache, ergebnis,
      passagiere_json, personal_json, chronologie_json,
      erstellt_am, geaendert_am
"""
import json
import sqlite3
from datetime import datetime
from pathlib import Path

from config import VORKOMMNISSE_DB_PATH as _DB_PFAD_STR

_DB_PFAD = Path(_DB_PFAD_STR)

# ── Verbindung / Init ─────────────────────────────────────────────────────────

def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_DB_PFAD), timeout=5)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous  = NORMAL")
    conn.execute("PRAGMA busy_timeout  = 5000")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _push(row_id: int) -> None:
    """Async-Push einer einzelnen Zeile nach Turso."""
    try:
        from database.turso_sync import push_row
        conn = _connect()
        row = conn.execute(
            "SELECT * FROM vorkommnisse WHERE id = ?", (row_id,)
        ).fetchone()
        conn.close()
        if row:
            push_row(str(_DB_PFAD), "vorkommnisse", dict(row))
    except Exception:
        pass


def _init_db() -> None:
    _DB_PFAD.parent.mkdir(parents=True, exist_ok=True)
    with _connect() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS vorkommnisse (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            flug             TEXT    NOT NULL DEFAULT '',
            typ              TEXT    NOT NULL DEFAULT '',
            bereich          TEXT    NOT NULL DEFAULT '',
            datum            TEXT    NOT NULL DEFAULT '',
            ort              TEXT    NOT NULL DEFAULT '',
            offblock_plan    TEXT    NOT NULL DEFAULT '',
            offblock_ist     TEXT    NOT NULL DEFAULT '',
            erstellt_von     TEXT    NOT NULL DEFAULT '',
            ursache          TEXT    NOT NULL DEFAULT '',
            ergebnis         TEXT    NOT NULL DEFAULT '',
            passagiere_json  TEXT    NOT NULL DEFAULT '[]',
            personal_json    TEXT    NOT NULL DEFAULT '[]',
            chronologie_json TEXT    NOT NULL DEFAULT '[]',
            erstellt_am      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
            geaendert_am     TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
        );
        """)
        # Migration: bereich-Spalte für bestehende Datenbanken hinzufügen
        cols = [row[1] for row in conn.execute("PRAGMA table_info(vorkommnisse)").fetchall()]
        if "bereich" not in cols:
            conn.execute("ALTER TABLE vorkommnisse ADD COLUMN bereich TEXT NOT NULL DEFAULT ''")


_init_db()


# ── CRUD ──────────────────────────────────────────────────────────────────────

def speichern(daten: dict) -> int:
    """
    Speichert ein neues Vorkommnis.  daten entspricht dem Dict aus _sammle_daten().
    Gibt die neue id zurück.
    """
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _connect() as conn:
        cur = conn.execute(
            """INSERT INTO vorkommnisse
               (flug, typ, bereich, datum, ort, offblock_plan, offblock_ist,
                erstellt_von, ursache, ergebnis,
                passagiere_json, personal_json, chronologie_json,
                erstellt_am, geaendert_am)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                daten.get("flug", ""),
                daten.get("typ", ""),
                daten.get("bereich", ""),
                daten.get("datum", ""),
                daten.get("ort", ""),
                daten.get("offblock_plan", ""),
                daten.get("offblock_ist", ""),
                daten.get("erstellt_von", ""),
                daten.get("ursache", ""),
                daten.get("ergebnis", ""),
                json.dumps(daten.get("passagiere", []), ensure_ascii=False),
                json.dumps(daten.get("personal", []), ensure_ascii=False),
                json.dumps(daten.get("chronologie", []), ensure_ascii=False),
                jetzt,
                jetzt,
            ),
        )
        new_id = cur.lastrowid
    _push(new_id)
    return new_id


def aktualisieren(vorkommnis_id: int, daten: dict) -> None:
    """Überschreibt ein vorhandenes Vorkommnis komplett."""
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _connect() as conn:
        conn.execute(
            """UPDATE vorkommnisse SET
               flug=?, typ=?, bereich=?, datum=?, ort=?,
               offblock_plan=?, offblock_ist=?,
               erstellt_von=?, ursache=?, ergebnis=?,
               passagiere_json=?, personal_json=?, chronologie_json=?,
               geaendert_am=?
               WHERE id=?""",
            (
                daten.get("flug", ""),
                daten.get("typ", ""),
                daten.get("bereich", ""),
                daten.get("datum", ""),
                daten.get("ort", ""),
                daten.get("offblock_plan", ""),
                daten.get("offblock_ist", ""),
                daten.get("erstellt_von", ""),
                daten.get("ursache", ""),
                daten.get("ergebnis", ""),
                json.dumps(daten.get("passagiere", []), ensure_ascii=False),
                json.dumps(daten.get("personal", []), ensure_ascii=False),
                json.dumps(daten.get("chronologie", []), ensure_ascii=False),
                jetzt,
                vorkommnis_id,
            ),
        )
    _push(vorkommnis_id)


def loeschen(vorkommnis_id: int) -> None:
    with _connect() as conn:
        conn.execute("DELETE FROM vorkommnisse WHERE id = ?", (vorkommnis_id,))


def lade_alle() -> list[dict]:
    """Gibt alle Vorkommnisse als Liste (neueste zuerst) zurück."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM vorkommnisse ORDER BY erstellt_am DESC, id DESC"
        ).fetchall()
    result = []
    for row in rows:
        d = dict(row)
        d["passagiere"]  = json.loads(d.pop("passagiere_json",  "[]"))
        d["personal"]    = json.loads(d.pop("personal_json",    "[]"))
        d["chronologie"] = json.loads(d.pop("chronologie_json", "[]"))
        result.append(d)
    return result


def lade_ein(vorkommnis_id: int) -> dict | None:
    with _connect() as conn:
        row = conn.execute(
            "SELECT * FROM vorkommnisse WHERE id = ?", (vorkommnis_id,)
        ).fetchone()
    if row is None:
        return None
    d = dict(row)
    d["passagiere"]  = json.loads(d.pop("passagiere_json",  "[]"))
    d["personal"]    = json.loads(d.pop("personal_json",    "[]"))
    d["chronologie"] = json.loads(d.pop("chronologie_json", "[]"))
    return d
