"""
Export-Historie: Speichert jeden Word-Export als Eintrag.
Datenbank: database SQL/export_historie.db
"""
import json
import os
import sqlite3
from datetime import datetime

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import _DB_DIR

_HIST_DB = os.path.join(_DB_DIR, "export_historie.db")

_SCHEMA = """
CREATE TABLE IF NOT EXISTS export_historie (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    von_datum       TEXT    NOT NULL,
    bis_datum       TEXT    NOT NULL,
    pax_zahl        INTEGER DEFAULT 0,
    bulmor_aktiv    INTEGER DEFAULT 5,
    einsaetze_zahl  INTEGER DEFAULT 0,
    sl_tag_name     TEXT    DEFAULT '',
    sl_nacht_name   TEXT    DEFAULT '',
    format          TEXT    DEFAULT 'dashboard',
    word_pfad       TEXT    DEFAULT '',
    excel_pfad      TEXT    DEFAULT '',
    params_json     TEXT    DEFAULT '{}',
    exportiert_am   TEXT    NOT NULL
);
"""


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(_HIST_DB, timeout=5, check_same_thread=False)
    conn.row_factory = lambda c, r: dict(zip([x[0] for x in c.description], r))
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 5000")
    conn.execute(_SCHEMA)
    conn.commit()
    return conn


def speichere_export(
    von_datum: str,
    bis_datum: str,
    pax_zahl: int,
    bulmor_aktiv: int,
    einsaetze_zahl: int,
    sl_tag_name: str,
    sl_nacht_name: str,
    format_: str,
    word_pfad: str,
    excel_pfad: str = "",
    params: dict | None = None,
) -> int:
    """Speichert einen Export-Eintrag und gibt die neue ID zurück."""
    params_json   = json.dumps(params or {}, ensure_ascii=False, default=str)
    exportiert_am = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = _get_conn()
    cur  = conn.execute(
        """INSERT INTO export_historie
               (von_datum, bis_datum, pax_zahl, bulmor_aktiv, einsaetze_zahl,
                sl_tag_name, sl_nacht_name, format, word_pfad, excel_pfad,
                params_json, exportiert_am)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (von_datum, bis_datum, pax_zahl, bulmor_aktiv, einsaetze_zahl,
         sl_tag_name, sl_nacht_name, format_, word_pfad, excel_pfad,
         params_json, exportiert_am),
    )
    conn.commit()
    return cur.lastrowid


def lade_alle_exporte(limit: int = 200) -> list[dict]:
    """Alle gespeicherten Exporte, neueste zuerst."""
    conn = _get_conn()
    return conn.execute(
        "SELECT * FROM export_historie ORDER BY exportiert_am DESC LIMIT ?",
        (limit,),
    ).fetchall()


def loesche_export(eintrag_id: int) -> None:
    """Löscht einen Export-Eintrag."""
    conn = _get_conn()
    conn.execute("DELETE FROM export_historie WHERE id = ?", (eintrag_id,))
    conn.commit()


def aktualisiere_word_pfad(eintrag_id: int, word_pfad: str) -> None:
    conn = _get_conn()
    conn.execute(
        "UPDATE export_historie SET word_pfad = ? WHERE id = ?",
        (word_pfad, eintrag_id),
    )
    conn.commit()
