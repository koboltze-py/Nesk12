"""
Sonderaufgaben-Datenbank
Speichert Snapshots der Sonderaufgaben-Eingaben in der SQLite-Datenbank (nesk3.db).
Die DB liegt im Ordner 'database SQL' und unterliegt den normalen Backup-Regeln.
"""
import json
import sqlite3
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import DB_PATH


# ── Interne Verbindung ────────────────────────────────────────────────────────

def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=5, check_same_thread=False)
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def _row_factory(cursor: sqlite3.Cursor, row: tuple) -> dict:
    return dict(zip([c[0] for c in cursor.description], row))


# ── Tabellen-Initialisierung ──────────────────────────────────────────────────

def init_table() -> None:
    """Erstellt die Tabelle sonderaufgaben_snapshots falls sie noch nicht existiert."""
    conn = _get_conn()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sonderaufgaben_snapshots (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                datum           TEXT    NOT NULL DEFAULT '',
                gespeichert_am  TEXT    NOT NULL DEFAULT '',
                aktion          TEXT    NOT NULL DEFAULT 'speichern',
                entries_json    TEXT    NOT NULL DEFAULT '{}',
                vorfeld_json    TEXT    NOT NULL DEFAULT '{}',
                bemerkung       TEXT    NOT NULL DEFAULT '',
                dienstplan_pfad TEXT    NOT NULL DEFAULT '',
                excel_datei     TEXT    NOT NULL DEFAULT ''
            )
        """)
        conn.commit()
    finally:
        conn.close()


# ── Schreiben ─────────────────────────────────────────────────────────────────

def save_snapshot(
    *,
    datum: str,
    entries: dict[str, str],
    vorfeld: dict[str, str],
    bemerkung: str,
    dienstplan_pfad: str = "",
    excel_datei: str = "",
    aktion: str = "speichern",
) -> int:
    """Speichert einen Snapshot und gibt die neue ID zurück."""
    init_table()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO sonderaufgaben_snapshots
                (datum, gespeichert_am, aktion, entries_json, vorfeld_json,
                 bemerkung, dienstplan_pfad, excel_datei)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datum,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                aktion,
                json.dumps(entries, ensure_ascii=False),
                json.dumps(vorfeld, ensure_ascii=False),
                bemerkung,
                dienstplan_pfad,
                excel_datei,
            ),
        )
        conn.commit()
        return cur.lastrowid  # type: ignore[return-value]
    finally:
        conn.close()


# ── Lesen ─────────────────────────────────────────────────────────────────────

def get_snapshots(limit: int = 300) -> list[dict]:
    """Gibt die letzten Snapshots zurück (neueste zuerst) ohne die JSON-Blobs."""
    init_table()
    conn = _get_conn()
    conn.row_factory = _row_factory
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, datum, gespeichert_am, aktion, bemerkung, excel_datei
            FROM   sonderaufgaben_snapshots
            ORDER  BY id DESC
            LIMIT  ?
            """,
            (limit,),
        )
        return cur.fetchall()
    finally:
        conn.close()


def get_snapshot_by_id(snap_id: int) -> dict | None:
    """Gibt einen vollständigen Snapshot inkl. JSON-Felder zurück."""
    init_table()
    conn = _get_conn()
    conn.row_factory = _row_factory
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM sonderaufgaben_snapshots WHERE id = ?",
            (snap_id,),
        )
        row = cur.fetchone()
        if row is None:
            return None
        # JSON-Felder deserialisieren
        try:
            row["entries"] = json.loads(row.get("entries_json") or "{}")
        except Exception:
            row["entries"] = {}
        try:
            row["vorfeld"] = json.loads(row.get("vorfeld_json") or "{}")
        except Exception:
            row["vorfeld"] = {}
        return row
    finally:
        conn.close()


def delete_snapshot(snap_id: int) -> None:
    """Löscht einen Snapshot aus der Datenbank."""
    init_table()
    conn = _get_conn()
    try:
        conn.execute(
            "DELETE FROM sonderaufgaben_snapshots WHERE id = ?", (snap_id,)
        )
        conn.commit()
    finally:
        conn.close()
