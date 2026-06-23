# Turso Cloud-Synchronisierung

## Übersicht

Turso ist eine libSQL-Cloud-Datenbank, die als Single Source of Truth (SSOT) für alle Nesk3-Installationen dient. Jeder PC hat seine eigene lokale SQLite-Kopie; Turso hält alle Instanzen synchron.

**Turso URL**: `https://nesk-koboltze.aws-eu-west-1.turso.io`

---

## Synchronisierungsablauf

```
App-Start
  └─→ pull_all()          ← Neuester Stand aus Turso in alle lokalen DBs
  └─→ Hintergrund-Thread
        └─→ alle 30s pull_all()

Bei jedem DB-Write
  └─→ push_row()          ← Datensatz parallel zu Turso senden
```

---

## Tabellennamensschema

Lokale Tabellen werden in Turso mit dem Dateiname-Prefix gespeichert:

| Lokale DB | Lokale Tabelle | Turso-Tabellenname |
|---|---|---|
| `nesk3.db` | `mitarbeiter` | `nesk3__mitarbeiter` |
| `nesk3.db` | `abteilungen` | `nesk3__abteilungen` |
| `nesk3.db` | `positionen` | `nesk3__positionen` |
| `nesk3.db` | `dienstplan` | `nesk3__dienstplan` |
| `nesk3.db` | `fahrzeuge` | `nesk3__fahrzeuge` |
| `nesk3.db` | `fahrzeug_status` | `nesk3__fahrzeug_status` |
| `nesk3.db` | `fahrzeug_schaeden` | `nesk3__fahrzeug_schaeden` |
| `nesk3.db` | `fahrzeug_termine` | `nesk3__fahrzeug_termine` |
| `nesk3.db` | `uebergabe_protokolle` | `nesk3__uebergabe_protokolle` |
| `nesk3.db` | `uebergabe_fahrzeug_notizen` | `nesk3__uebergabe_fahrzeug_notizen` |
| `nesk3.db` | `uebergabe_handy_eintraege` | `nesk3__uebergabe_handy_eintraege` |
| `nesk3.db` | `uebergabe_verspaetungen` | `nesk3__uebergabe_verspaetungen` |
| `nesk3.db` | `settings` | `nesk3__settings` |
| `mitarbeiter.db` | `mitarbeiter` | `ma__mitarbeiter` |
| `mitarbeiter.db` | `positionen` | `ma__positionen` |
| `mitarbeiter.db` | `abteilungen` | `ma__abteilungen` |
| `einsaetze.db` | `einsaetze` | `einsaetze__einsaetze` |
| `verspaetungen.db` | `verspaetungen` | `vers__verspaetungen` |
| `telefonnummern.db` | `telefonnummern` | `tel__telefonnummern` |
| `telefonnummern.db` | `tel_import_log` | `tel__import_log` |
| `patienten_station.db` | `patienten` | `pat__patienten` |
| `patienten_station.db` | `medikamente` | `pat__medikamente` |
| `patienten_station.db` | `verbrauchsmaterial` | `pat__verbrauchsmaterial` |
| `call_transcription.db` | `call_logs` | `call__call_logs` |
| `call_transcription.db` | `textbausteine` | `call__textbausteine` |
| `psa.db` | `psa_verstoss` | `psa__psa_verstoss` |
| `stellungnahmen.db` | `stellungnahmen` | `stelg__stellungnahmen` |
| `beschwerden.db` | `beschwerden` | `bschw__beschwerden` |
| `beschwerden.db` | `beschwerde_antworten` | `bschw__antworten` |
| `vorkommnisse.db` | `vorkommnisse` | `vork__vorkommnisse` |
| `handys.db` | `handys` | `handys__handys` |
| `handys.db` | `handys_historie` | `handys__historie` |

**Tabellen, die NICHT synchronisiert werden**: `sqlite_sequence`, `backup_log`

---

## Hauptfunktionen (`database/turso_sync.py`)

### `push_row(db_datei, tabelle, zeile_dict)`

Schreibt einen einzelnen Datensatz nach Turso:
```python
push_row("nesk3.db", "mitarbeiter", {"id": 5, "vorname": "Max", ...})
```

### `pull_all()`

Holt alle Turso-Tabellen und synchronisiert in die lokalen DBs:
```python
pull_all()  # Überschreibt lokale Daten mit Turso-Stand
```

### `_turso_request(sql, params=None)`

Interne HTTP-Anfrage an die Turso REST-API:
```python
# POST https://nesk-koboltze.aws-eu-west-1.turso.io
# Authorization: Bearer <TURSO_TOKEN>
# Content-Type: application/json
# Body: {"statements": [{"q": sql, "params": params}]}
```

---

## OneDrive-Artefaktbereinigung (`main.py`)

Beim App-Start werden automatisch zwei Artefakttypen bereinigt:

### 1. WAL/SHM-Dateien
```python
# Für jede *.db-wal Datei:
# 1. PRAGMA wal_checkpoint(TRUNCATE) ausführen
# 2. Datei löschen
# Verhindert: OneDrive-Dialog "568 Elemente löschen?"
```

### 2. OneDrive-Konfliktkopien
```python
# Muster: nesk3-W11-262013-8.db (PC-Name + Nummer)
# Diese entstehen wenn 2 PCs gleichzeitig die gleiche .db bearbeiten
# Werden automatisch gelöscht (Daten sind in nesk3.db + Turso gesichert)
```

---

## Backup-Sequenz beim App-Start

```python
# 1. WAL/SHM bereinigen
_cleanup_onedrive_artefakte()

# 2. Datenbankschema erstellen/migrieren
init_database()

# 3. Turso-Pull (neueste Daten holen)
from database.turso_sync import pull_all
pull_all()

# 4. Lokale DB-Backups erstellen
_erstelle_db_backup()   # → database SQL/db_backups/YYYY-MM-DD/

# 5. App starten
app = QApplication(...)
window = MainWindow()
window.show()

# 6. Hintergrund-Sync starten
threading.Thread(target=_background_sync, daemon=True).start()
```
