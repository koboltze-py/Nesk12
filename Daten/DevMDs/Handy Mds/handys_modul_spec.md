# Nesk3 – Modul: Handys

## Übersicht

Das Modul **Handys** ersetzt den bisherigen Eintrag „Call Transcription" in der Sidebar.
Es dient der Verwaltung aller Diensthandys der DRK-Erste-Hilfe-Station am Flughafen Köln/Bonn,
ähnlich der bestehenden Fahrzeugverwaltung.

---

## 1. Sidebar-Änderung

| Vorher | Nachher |
|--------|---------|
| `Call Transcription` (Button + Icon) | `Handys` (Button + Icon: `phone` o. ä.) |

- Eintrag in `sidebar.py` (o. ä. Navigation-Konfiguration) entfernen: `Call Transcription`
- Neuen Eintrag hinzufügen:
  - **Label:** `Handys`
  - **Icon:** `ph.icons.PHONE` (oder passendes PySide6-Icon)
  - **Modul:** `handys_widget.py`

---

## 2. Datenbankdatei

| Eigenschaft | Wert |
|-------------|------|
| **Dateiname** | `handys.db` |
| **Speicherort** | Identisch mit anderen Nesk3-Datenbanken (z. B. `Daten/`) |
| **WAL-Modus** | `PRAGMA journal_mode=WAL;` |
| **Backup** | Vollständig in `BackupManager` registriert (gleiche Regeln wie `nesk3.db`) |

→ Detailliertes Schema: siehe [`handys_datenbank.md`](handys_datenbank.md)

---

## 3. Widget-Struktur (`handys_widget.py`)

Das Widget enthält **drei Tabs**:

| Tab | Inhalt |
|-----|--------|
| `Übersicht` | Tabelle aller Handys mit Status-Badges |
| `Details / Bearbeiten` | Formular zum Anlegen / Bearbeiten eines Geräts |
| `Historie` | Verlauf aller Zustandsänderungen |

---

## 4. Tab 1 – Übersicht

### Tabellenfelder (QTableWidget / QTreeView)

| Spalte | Quelle (DB-Feld) | Hinweis |
|--------|-----------------|---------|
| ID | `id` | Intern, ausgeblendet oder klein |
| Inventarnummer | `inventarnummer` | Pflichtfeld |
| Hersteller | `hersteller` | z. B. Samsung, Apple |
| Modell | `modell` | z. B. Galaxy A55 |
| Rufnummer | `rufnummer` | Dienstnummer |
| SIM-Karte | `sim_nummer` | ICCID oder Kurzbezeichnung |
| Standort / Ausgabe | `standort` | z. B. „Disposition", „Schichtleiter" |
| Zustand | `zustand` | Enum: `Aktiv`, `Defekt`, `Außer Betrieb`, `Reserve` |
| Defektbeschreibung | `defekt_beschreibung` | Freitext, nur wenn `zustand = Defekt` |
| Anschaffungsdatum | `anschaffungsdatum` | Datum |
| Zuletzt aktualisiert | `geaendert_am` | Auto-Timestamp |

### Aktions-Buttons (oberhalb der Tabelle)

```
[+ Neues Handy]   [✎ Bearbeiten]   [✕ Löschen]   [⟳ Aktualisieren]
[📤 Excel exportieren]   [📧 Per E-Mail senden]
```

### Status-Farbkodierung

| Zustand | Farbe |
|---------|-------|
| Aktiv | Grün |
| Defekt | Rot |
| Außer Betrieb | Orange |
| Reserve | Grau |

---

## 5. Tab 2 – Details / Bearbeiten

Formularfelder (QLineEdit / QComboBox / QDateEdit):

```
Inventarnummer *     [________________]
Hersteller           [________________]
Modell               [________________]
Rufnummer            [________________]
SIM-Nummer           [________________]
Standort / Ausgabe   [________________]
Zustand *            [Dropdown: Aktiv | Defekt | Außer Betrieb | Reserve]
Defektbeschreibung   [________________]  ← nur aktiv wenn Zustand = Defekt
Anschaffungsdatum    [Datum-Picker     ]
Notizen              [Mehrzeiliges Textfeld]

[Speichern]   [Abbrechen]
```

- `*` = Pflichtfeld mit Validierung
- Bei `Speichern`: Eintrag in `handys_historie` schreiben (wer, wann, was geändert)

---

## 6. Tab 3 – Historie

### Tabellenfelder

| Spalte | DB-Feld |
|--------|---------|
| Datum / Uhrzeit | `geaendert_am` |
| Handy (Inv.-Nr.) | `inventarnummer` (JOIN) |
| Geändertes Feld | `feld` |
| Alter Wert | `alter_wert` |
| Neuer Wert | `neuer_wert` |
| Benutzer | `benutzer` |

- Filterbar nach Handy (Dropdown) und Zeitraum
- Keine Bearbeitungsmöglichkeit (read-only)

---

## 7. Excel-Export

→ Details: siehe [`handys_excel_export.md`](handys_excel_export.md)

**Exportpfad (fest):**
```
C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\
Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\
Nesk\Nesk3\Daten\Handys\
```

**Dateiname:** `Handys_Uebersicht_YYYY-MM-DD.xlsx`

---

## 8. E-Mail-Funktion

→ Details: siehe [`handys_email.md`](handys_email.md)

---

## 9. Integration in bestehende Nesk3-Architektur

| Komponente | Änderung |
|------------|---------|
| `BackupManager` | `handys.db` in Backup-Liste aufnehmen |
| `config.py` / Pfadkonstanten | `HANDYS_EXPORT_PATH` als Konstante definieren |
| `cd.py` (Corporate Design) | Excel-Export nutzt bestehende CD-Funktionen |
| `sidebar.py` / Navigation | `Call Transcription` entfernen, `Handys` einfügen |
| `main_window.py` | `HandysWidget` instanziieren und registrieren |

---

## 10. Datei-Übersicht (neue Dateien)

```
nesk3/
├── modules/
│   └── handys/
│       ├── handys_widget.py          # Haupt-Widget (3 Tabs)
│       ├── handys_db.py              # DB-Zugriff (CRUD + Historie)
│       ├── handys_excel_export.py    # Excel-Exportlogik
│       └── handys_email.py           # E-Mail-Versand
├── data/
│   └── handys.db                     # SQLite-Datenbank
└── docs/
    ├── handys_modul_spec.md          # Diese Datei
    ├── handys_datenbank.md           # DB-Schema
    ├── handys_excel_export.md        # Export-Spezifikation
    └── handys_email.md               # E-Mail-Spezifikation
```
