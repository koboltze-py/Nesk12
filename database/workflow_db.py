"""
Workflow-Datenbank
==================
Speichert pro SM↔Dienstplan-Tag:
  - ob mit Carmen abgeglichen wurde (inkl. Zeitstempel)
  - eine freie Notiz

DB-Pfad: database SQL/workflow.db
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import WORKFLOW_DB_PATH


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(WORKFLOW_DB_PATH, timeout=5, check_same_thread=False)
    conn.row_factory = lambda c, r: dict(zip([col[0] for col in c.description], r))
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    return conn


def init_db() -> None:
    """Erstellt die Tabellen falls noch nicht vorhanden."""
    Path(WORKFLOW_DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS workflow_tag (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                datum                TEXT    NOT NULL,
                sm_datei             TEXT    NOT NULL,
                dp_datei             TEXT    NOT NULL DEFAULT '',
                abgeglichen_carmen   INTEGER NOT NULL DEFAULT 0,
                abgeglichen_carmen_am TEXT,
                notiz                TEXT    NOT NULL DEFAULT '',
                geaendert_am         TEXT,
                UNIQUE(datum, sm_datei)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS workflow_session (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                monat     TEXT    NOT NULL UNIQUE,
                sm_pfade  TEXT    NOT NULL DEFAULT '[]',
                dp_pfade  TEXT    NOT NULL DEFAULT '[]',
                last_used TEXT
            )
        """)
        conn.commit()


def _ensure_init() -> None:
    init_db()


def lade_eintrag(datum: str, sm_datei: str) -> dict:
    """
    Gibt den Datensatz für einen Tag zurück.
    Falls noch keiner existiert, ein leeres Default-Dict.
    """
    _ensure_init()
    with _get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM workflow_tag WHERE datum=? AND sm_datei=?",
            (datum, sm_datei),
        ).fetchone()
    return row or {
        "datum": datum,
        "sm_datei": sm_datei,
        "dp_datei": "",
        "abgeglichen_carmen": 0,
        "abgeglichen_carmen_am": None,
        "notiz": "",
        "geaendert_am": None,
    }


def speichere_eintrag(
    datum: str,
    sm_datei: str,
    dp_datei: str = "",
    abgeglichen_carmen: bool | None = None,
    notiz: str | None = None,
) -> None:
    """
    Speichert/aktualisiert einen Tages-Eintrag.
    Nur übergebene Felder (nicht None) werden aktualisiert.
    """
    _ensure_init()
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        # Existiert bereits?
        existing = conn.execute(
            "SELECT id, abgeglichen_carmen, notiz FROM workflow_tag WHERE datum=? AND sm_datei=?",
            (datum, sm_datei),
        ).fetchone()

        if existing:
            updates: list[str] = ["geaendert_am=?"]
            params: list = [jetzt]
            if dp_datei:
                updates.append("dp_datei=?");   params.append(dp_datei)
            if abgeglichen_carmen is not None:
                updates.append("abgeglichen_carmen=?")
                params.append(1 if abgeglichen_carmen else 0)
                if abgeglichen_carmen:
                    updates.append("abgeglichen_carmen_am=?")
                    params.append(jetzt)
                else:
                    updates.append("abgeglichen_carmen_am=?")
                    params.append(None)
            if notiz is not None:
                updates.append("notiz=?");  params.append(notiz)
            params += [datum, sm_datei]
            conn.execute(
                f"UPDATE workflow_tag SET {', '.join(updates)} WHERE datum=? AND sm_datei=?",
                params,
            )
        else:
            carmen_val = 1 if abgeglichen_carmen else 0
            carmen_am  = jetzt if abgeglichen_carmen else None
            conn.execute(
                """INSERT INTO workflow_tag
                   (datum, sm_datei, dp_datei, abgeglichen_carmen, abgeglichen_carmen_am, notiz, geaendert_am)
                   VALUES (?,?,?,?,?,?,?)""",
                (datum, sm_datei, dp_datei or "",
                 carmen_val, carmen_am,
                 notiz or "", jetzt),
            )
        conn.commit()


def alle_eintraege() -> list[dict]:
    """Gibt alle gespeicherten Einträge zurück."""
    _ensure_init()
    with _get_conn() as conn:
        return conn.execute(
            "SELECT * FROM workflow_tag ORDER BY datum, sm_datei"
        ).fetchall()


# ── Session-Persistenz (geladene Monats-Dateien) ─────────────────────────────

import json as _json  # noqa: E402  (nach den anderen imports)


def alle_monate() -> list[str]:
    """Gibt alle gespeicherten Monate zurück (YYYY-MM), neueste zuerst."""
    _ensure_init()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT monat FROM workflow_session ORDER BY monat DESC"
        ).fetchall()
    return [r["monat"] for r in rows]


def lade_session(monat: str) -> dict:
    """Gibt Session-Daten für einen Monat zurück oder Default-Dict."""
    _ensure_init()
    with _get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM workflow_session WHERE monat=?", (monat,)
        ).fetchone()
    return row or {"monat": monat, "sm_pfade": "[]", "dp_pfade": "[]", "last_used": None}


def speichere_session(monat: str, sm_pfade: list, dp_pfade: list) -> None:
    """Speichert/aktualisiert die Dateilisten für einen Monat."""
    _ensure_init()
    jetzt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        conn.execute(
            """
            INSERT INTO workflow_session (monat, sm_pfade, dp_pfade, last_used)
            VALUES (?,?,?,?)
            ON CONFLICT(monat) DO UPDATE SET
                sm_pfade  = excluded.sm_pfade,
                dp_pfade  = excluded.dp_pfade,
                last_used = excluded.last_used
            """,
            (
                monat,
                _json.dumps(sm_pfade, ensure_ascii=False),
                _json.dumps(dp_pfade, ensure_ascii=False),
                jetzt,
            ),
        )
        conn.commit()


def loesche_session(monat: str) -> None:
    """Löscht die Session eines Monats (Dateilisten, nicht Carmen/Notizen)."""
    _ensure_init()
    with _get_conn() as conn:
        conn.execute("DELETE FROM workflow_session WHERE monat=?", (monat,))
        conn.commit()
