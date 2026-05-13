# Turso Cloud-Sync (Stand: 13.05.2026)

## Überblick

Turso ist eine **libSQL-Cloud-Datenbank** (SQLite-kompatibel).  
Nesk3 verwendet Turso als Read-Replica / Sync-Schicht – die **lokalen SQLite-DBs bleiben die primäre Datenquelle**.

---

## Verbindungsdaten

```python
TURSO_URL   = "https://nesk-koboltze.aws-eu-west-1.turso.io"
TURSO_TOKEN = "<JWT-Token>"   # in config.py
TURSO_SYNC_INTERVAL = 30      # Sekunden
```

API-Endpunkt: `POST {TURSO_URL}/v2/pipeline`  
Auth: `Authorization: Bearer {TURSO_TOKEN}`

---

## Sync-Richtungen

```
App-Start:   Turso → lokal      (pull_all)
Jeder Write: lokal → Turso      (push_row)
Hintergrund: alle 30s lokal → Turso  (auto-sync Thread)
```

---

## TABLE_MAP (database/turso_sync.py)

Definiert welche lokalen Tabellen nach Turso synchronisiert werden:

| Lokale DB | Lokale Tabelle | Turso-Tabellenname |
|---|---|---|
| nesk3.db | mitarbeiter | `nesk3__mitarbeiter` |
| nesk3.db | abteilungen | `nesk3__abteilungen` |
| nesk3.db | positionen | `nesk3__positionen` |
| nesk3.db | dienstplan | `nesk3__dienstplan` |
| nesk3.db | fahrzeuge | `nesk3__fahrzeuge` |
| nesk3.db | fahrzeug_status | `nesk3__fahrzeug_status` |
| nesk3.db | fahrzeug_schaeden | `nesk3__fahrzeug_schaeden` |
| nesk3.db | fahrzeug_termine | `nesk3__fahrzeug_termine` |
| nesk3.db | uebergabe_protokolle | `nesk3__uebergabe_protokolle` |
| nesk3.db | uebergabe_fahrzeug_notizen | `nesk3__uebergabe_fahrzeug_notizen` |
| nesk3.db | uebergabe_handy_eintraege | `nesk3__uebergabe_handy_eintraege` |
| nesk3.db | uebergabe_verspaetungen | `nesk3__uebergabe_verspaetungen` |
| nesk3.db | settings | `nesk3__settings` |
| nesk3.db | backup_log | `nesk3__backup_log` |
| mitarbeiter.db | mitarbeiter | `ma__mitarbeiter` |
| mitarbeiter.db | positionen | `ma__positionen` |
| mitarbeiter.db | abteilungen | `ma__abteilungen` |
| einsaetze.db | einsaetze | `einsaetze__einsaetze` |
| verspaetungen.db | verspaetungen | `vers__verspaetungen` |
| telefonnummern.db | telefonnummern | `tel__telefonnummern` |
| telefonnummern.db | tel_import_log | `tel__import_log` |
| patienten_station.db | patienten | `pat__patienten` |
| patienten_station.db | medikamente | `pat__medikamente` |
| patienten_station.db | verbrauchsmaterial | `pat__verbrauchsmaterial` |
| call_transcription.db | call_logs | `call__call_logs` |
| call_transcription.db | textbausteine | `call__textbausteine` |
| psa.db | psa_verstoss | `psa__psa_verstoss` |
| stellungnahmen.db | stellungnahmen | `stelg__stellungnahmen` |
| beschwerden.db | beschwerden | `bschw__beschwerden` |
| beschwerden.db | beschwerde_antworten | `bschw__antworten` |
| vorkommnisse.db | vorkommnisse | `vork__vorkommnisse` |
| handys.db | handys | `handys__handys` |
| handys.db | handys_historie | `handys__historie` |

**Nicht synchronisiert:** `sqlite_sequence`, `backup_log`

---

## Kern-Funktionen

### `push_row(db_path, table, row_dict)`
```python
def push_row(db_path: str, table: str, row_dict: dict):
    """Schreibt einen Datensatz nach Turso (UPSERT)."""
    db_name = os.path.basename(db_path)
    key = (db_name, table)
    turso_table = TABLE_MAP.get(key)
    if not turso_table:
        return  # Tabelle nicht synchronisiert
    
    cols = list(row_dict.keys())
    vals = list(row_dict.values())
    placeholders = ", ".join(["?"] * len(cols))
    sql = f"INSERT OR REPLACE INTO {turso_table} ({', '.join(cols)}) VALUES ({placeholders})"
    _turso_request(sql, vals)
```

### `pull_all()`
```python
def pull_all():
    """Holt alle Daten aus Turso in die lokalen DBs."""
    for (db_name, local_table), turso_table in TABLE_MAP.items():
        rows = _turso_request(f"SELECT * FROM {turso_table}")
        # → schreibt in lokale SQLite-DB
```

### `_turso_request(sql, params)`
```python
def _turso_request(sql: str, params: list | None = None) -> dict:
    """Sendet SQL-Anfrage via HTTP POST an Turso /v2/pipeline."""
    body = {
        "requests": [
            {"type": "execute", "stmt": {"sql": sql, "args": [...]}},
            {"type": "close"}
        ]
    }
    # POST → TURSO_URL/v2/pipeline
    # Header: Authorization: Bearer <TOKEN>
```

---

## Auto-Sync Thread

```python
# In main.py oder turso_sync.py:
def _start_turso_sync():
    def _sync_loop():
        while True:
            time.sleep(TURSO_SYNC_INTERVAL)  # 30s
            try:
                pull_all()
            except Exception:
                pass  # offline → kein Fehler
    
    t = threading.Thread(target=_sync_loop, daemon=True)
    t.start()
```

---

## Fehlerverhalten

- Alle Turso-Operationen sind in `try/except` gewrappt
- Bei Netzwerkfehler: **kein Absturz**, App läuft weiter offline
- `push_row()` wird von `_push()` in jedem `_db.py`-Modul aufgerufen, ebenfalls `try/except`
- Kein Retry-Mechanismus – nächster Sync erfolgt nach 30s automatisch
