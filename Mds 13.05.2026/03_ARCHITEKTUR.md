# Architektur – Module, Navigation, Datenfluss (Stand: 13.05.2026)

## Gesamtarchitektur

```
main.py
  └── MainWindow (gui/main_window.py)
        ├── Sidebar (220px, #354a5e)
        │     ├── _NeskLogoWidget (animiert, 33 FPS)
        │     └── SidebarButton × 15 (navigiert QStackedWidget)
        └── QStackedWidget (Content-Bereich)
              ├── [0]  DashboardWidget
              ├── [1]  MitarbeiterHauptWidget
              ├── [2]  DienstlichesWidget
              ├── [3]  AufgabenHauptWidget
              ├── [4]  DienstplanWidget
              ├── [5]  UebergabeWidget
              ├── [6]  FahrzeugeWidget
              ├── [7]  Code19Widget
              ├── [8]  BerichtWidget
              ├── [9]  PassagiereWidget
              ├── [10] TelefonnummernWidget
              ├── [11] HandysWidget
              ├── [12] BackupWidget
              ├── [13] EinstellungenWidget
              └── [14] VorkommnisseWidget
```

---

## Sidebar-Navigation

**15 Einträge** in `NAV_ITEMS`:

| Index | Icon | Label | Widget |
|---|---|---|---|
| 0 | 🏠 | Dashboard | `DashboardWidget` |
| 1 | 👥 | Mitarbeiter | `MitarbeiterHauptWidget` |
| 2 | ☕️ | Dienstliches | `DienstlichesWidget` |
| 3 | 📝 | Aufgaben | `AufgabenHauptWidget` |
| 4 | 📅 | Dienstplan | `DienstplanWidget` |
| 5 | 📋 | Übergabe | `UebergabeWidget` |
| 6 | 🚗 | Fahrzeuge | `FahrzeugeWidget` |
| 7 | 🕐 | Code 19 | `Code19Widget` |
| 8 | 📊 | Bericht | `BerichtWidget` |
| 9 | ✈️ | Passagiere | `PassagiereWidget` |
| 10 | 📞 | Telefonnummern | `TelefonnummernWidget` |
| 11 | 📱 | Handys | `HandysWidget` |
| 12 | 💾 | Backup | `BackupWidget` |
| 13 | ⚙️ | Einstellungen | `EinstellungenWidget` |
| 14 | ⚠️ | Vorkommnisse | `VorkommnisseWidget` |

---

## MainWindow-Klassen

### `_NeskLogoWidget` (200×170 px)
- Animiertes Logo in der Sidebar oben
- Zeichnet mit `QPainter`: Doppelring (Teal + Gold), pulsierender Glow, "NeSk"-Schriftzug mit Shimmer
- **Doppelklick**: Öffnet `SlotMachineDialog` (Easter Egg "Alice's Wunderrad")
- Timer: 30ms (~33 FPS)

### `SidebarButton(QPushButton)`
- Checkable, 48px hoch, klarer Hover-/Active-Stil
- Aktiv: `FIORI_BLUE` Hintergrund, weiße Schrift, fett
- Inaktiv: transparent, #cdd5e0 Schrift

### `MainWindow(QMainWindow)`
- Fenstergröße: 1280×800 (min. 900×600)
- `_navigate(index)`: Wechselt QStackedWidget + aktualisiert SidebarButton-Status
- `_check_termine_startup()`: 800ms nach Start: prüft bald fällige Fahrzeug-Termine

---

## Dashboard-Widget

```
DashboardWidget
  ├── StatCards-Reihe (Kacheln)
  │     ├── Mitarbeiter aktiv
  │     ├── Fahrzeuge aktiv
  │     ├── Dienstplan heute
  │     └── Passagierzahl heute
  ├── _SkyWidget (animiertes Flugzeug, 72px hoch)
  ├── _TerminKalender (QCalendarWidget)
  │     ├── Blauer Punkt = Fahrzeug-Termin
  │     └── Grüner Punkt = Notiz
  └── Notizen-Panel (rechts)
        ├── Button "➕ Neue Notiz"
        ├── Button "🗄️ Archiv"
        └── _zeige_notizen() – scrollbare Liste
```

**Notizen-Fenster:** `lade_fenster()` → heute-5 bis heute+10  
Jede Notiz hat Buttons: 👁 Gelesen, ↩️ Rückgängig (für erledigte), ✅ Erledigt, 🗑 Löschen

---

## HandysWidget – Tab-Struktur

```
HandysWidget
  ├── Tab "Geräteübersicht"   ← Tabelle aller Handys, Filter, Suche
  ├── Tab "Historien"         ← Änderungsprotokoll
  ├── Tab "Berichte"          ← Word-Berichte, Excel-Export
  └── Tab "Einstellungen"     ← Export-Pfad, Druckeinstellungen
```

---

## VorkommnisseWidget

- Linke Seite: Liste aller Vorkommnisse (Tabelle)
- Rechte Seite: Detailformular
  - Felder: Flugnummer, Typ, Datum, Ort, Offblock-Plan/-Ist, Erstellt-von
  - Sub-Tabellen: Passagiere (JSON), Personal (JSON), Chronologie (JSON)
  - **Auto-Save beim E-Mail-Dialog**: wenn `_current_id is None` und Flugnummer vorhanden → automatisch speichern vor E-Mail-Entwurf

---

## Aufgaben-Module

```
AufgabenHauptWidget
  ├── Reiter "Tag"     → AufgabenTagWidget (Tagesaufgaben, Checklisten)
  └── Reiter "Haupt"   → AufgabenWidget (Aufgaben Nacht, Sonderaufgaben)
```

---

## Datenlayer-Prinzip

```
GUI-Widget
  │
  ├── functions/<modul>_db.py    → direkte SQLite-Calls (Notizen, Vorkommnisse, Beschwerden, Handys, ...)
  │
  └── database/models.py         → ORM-ähnliche Abfragen für nesk3.db (Mitarbeiter, Dienstplan, Fahrzeuge, ...)
        └── database/connection.py  → get_connection() für nesk3.db
```

---

## Turso Sync-Architektur

```
Jeder Write-CRUD in functions/*_db.py
  └── _push(row_id) → database.turso_sync.push_row(db_path, table_name, row_dict)
        └── HTTP POST an Turso-Endpoint (JWT-Auth)
              └── TABLE_MAP in turso_sync.py definiert welche Tabellen/DBs repliziert werden
```

**Außerdem:** Auto-Sync-Thread alle 30 Sekunden (via `TURSO_SYNC_INTERVAL`)

---

## SAP Fiori Design-Konstanten

| Konstante | Wert | Verwendung |
|---|---|---|
| `FIORI_BLUE` | `#0a6ed1` | Primärfarbe, aktive Sidebar-Buttons |
| `FIORI_LIGHT_BLUE` | `#eef4fa` | Hintergrundstreifen |
| `FIORI_TEXT` | `#32363a` | Standard-Textfarbe |
| `FIORI_BORDER` | `#d9d9d9` | Rahmen, Trennlinien |
| `FIORI_SUCCESS` | `#107e3e` | Erfolgs-Badges |
| `FIORI_WARNING` | `#e9730c` | Warnungen |
| `FIORI_ERROR` | `#bb0000` | Fehler-Anzeigen |
| `FIORI_SIDEBAR_BG` | `#354a5e` | Sidebar-Hintergrund |
