# Nesk3 – Datenbankübersicht

**Stand:** 28.04.2026 – v3.11.0  
**Alle Dateien liegen im Ordner:** `database SQL/` (relativ zu `BASE_DIR`)  
**Modus:** SQLite 3, WAL-Modus (Write-Ahead Logging) für parallele Lese-/Schreibzugriffe

---

## Übersicht aller Datenbanken

| Datei | Konstante in `config.py` | Modul | Zweck |
|-------|--------------------------|-------|-------|
| `nesk3.db` | `DB_PATH` | `database/connection.py` | Hauptdatenbank: Fahrzeuge, Übergabe, Einstellungen, Dienstplan-Status |
| `archiv.db` | `ARCHIV_DB_PATH` | `functions/archiv_functions.py` | Archivierte Einträge (gelöschte Protokolle etc.) |
| `mitarbeiter.db` | `MITARBEITER_DB_PATH` | `functions/mitarbeiter_functions.py` | Mitarbeiterstammdaten, PSA-Tracking, Schulungen |
| `beschwerden.db` | `BESCHWERDEN_DB_PATH` | `functions/beschwerden_db.py` | Beschwerdemanagement |
| `sanmat.db` | `SANMAT_DB_PATH` | `database/sanmat_db.py` | Sanitätsmaterial-Bestand und Verbrauch |
| `vorkommnisse.db` | `VORKOMMNISSE_DB_PATH` | `functions/vorkommnisse_db.py` | Vorkommnisberichte |
| `notizen.db` | `NOTIZEN_DB_PATH` | `functions/notizen_db.py` | Dashboard-Notizen und Termine |
| *(weitere)* | – | `functions/verspaetung_db.py` | Verspätungsmeldungen |
| *(weitere)* | – | `functions/stellungnahmen_db.py` | Stellungnahmen |
| *(weitere)* | – | `functions/telefonnummern_db.py` | Telefonnummern-Verzeichnis |
| *(weitere)* | – | `database/pax_db.py` | Patienten- und Einsatzzähler (PAX) |

---

## Datenbank-Details

### `nesk3.db` – Hauptdatenbank

Tabellen (verwaltet durch `database/migrations.py` beim App-Start):

| Tabelle | Wichtige Spalten | Beschreibung |
|---------|-----------------|--------------|
| `fahrzeuge` | id, kennzeichen, typ, hersteller, baujahr, notiz | Fahrzeugstammdaten |
| `fahrzeug_status` | id, fahrzeug_id, status, grund, von, bis, erstellt_am | Statushistorie je Fahrzeug (fahrbereit/defekt/werkstatt/ausser_dienst/sonstiges) |
| `uebergabe` | id, datum, von, bis, schichtleiter, notiz, … | Übergabe-Protokolle |
| `einstellungen` | key, value | Key-Value-Store für App-Einstellungen |
| `dienstplan_eintraege` | id, datum, name, dienst, krank_typ, … | Geladene Dienstplan-Daten aus Excel |
| `einsaetze` | id, datum, art, beschreibung, pax, sl_einsatz, … | Einsatz-/Veranstaltungsprotokoll |
| `patienten` | id, datum, name, diagnose, … | Patienten DRK Station |

### `mitarbeiter.db`

| Tabelle | Wichtige Spalten | Beschreibung |
|---------|-----------------|--------------|
| `mitarbeiter` | id, name, vorname, funktion, aktiv, export_aktiv | Mitarbeiterstammdaten |
| `schulungseintraege` | id, mitarbeiter_id, typ, datum, gueltig_bis, informiert, informiert_am | Schulungsdaten je Mitarbeiter |
| `psa_verstaesse` | id, mitarbeiter_id, datum, gesendet | PSA-Verstöße |
| `dokumente` | id, mitarbeiter_id, kategorie, titel, pfad, erstellt_am | Mitarbeiterdokumente |

### `sanmat.db`

| Tabelle | Wichtige Spalten | Beschreibung |
|---------|-----------------|--------------|
| `artikel` | id, name, kategorie, bestand, mindestbestand | Artikelstammdaten |
| `verbrauch` | id, artikel_id, menge, datum, notiz, erstellt_am | Verbrauchsbuchungen |

### `vorkommnisse.db`

| Tabelle | Wichtige Spalten | Beschreibung |
|---------|-----------------|--------------|
| `vorkommnisse` | id, datum, uhrzeit, ort, flugnummer, kategorie, beschreibung, massnahmen, offblock_plan, offblock_ist, verspaetung_min, personen (JSON), mitarbeiter, erstellt_am | Vollständige Vorkommnisberichte |

### `notizen.db`

| Tabelle | Wichtige Spalten | Beschreibung |
|---------|-----------------|--------------|
| `notizen` | id, datum, faellig_am, titel, text, gelesen, erledigt, erstellt_am | Dashboard-Notizen und Termine |

---

## Pfad-Logik (EXE vs. Script)

`config.py` ermittelt `BASE_DIR` automatisch:

- **Script-Modus** (`python main.py`): `BASE_DIR` = Ordner der `config.py`
- **EXE-Modus** (`Nesk3.exe`): `BASE_DIR` = OneDrive-Pfad aus `%OneDriveCommercial%` oder `%OneDrive%` + Unterordner `Dateien von Erste-Hilfe-Station-Flughafen.../Nesk/Nesk3`

Alle Datenbanken landen in `BASE_DIR/database SQL/` – sowohl im Script- als auch EXE-Modus auf dem gleichen physischen Pfad (OneDrive-Freigabe).

---

## Migrationen

`database/migrations.py` wird bei jedem App-Start ausgeführt und legt fehlende Tabellen/Spalten automatisch an (`CREATE TABLE IF NOT EXISTS`, `ALTER TABLE … ADD COLUMN IF NOT EXISTS`).  
Bestehende Daten bleiben erhalten.

---

## Turso (Cloud-Sync)

- **URL:** `https://nesk-koboltze.aws-eu-west-1.turso.io`
- **Sync-Intervall:** 30 Sekunden
- **Modul:** `database/turso_sync.py`
- Turso spiegelt die Hauptdatenbank (`nesk3.db`) in die Cloud – für Zugriff von mehreren PCs über OneDrive hinaus.
