# Nesk3 – Excel-Export: Handys

## Übersicht

Der Excel-Export erzeugt eine formatierte Übersichtsdatei aller Diensthandys
im Nesk3-Corporate-Design (`cd.py`) und speichert sie lokal.
Der Export wird sowohl manuell (Button) als auch automatisch beim E-Mail-Versand ausgelöst.

---

## Exportpfad

```
C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\
Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\
Nesk\Nesk3\Daten\Handys\
```

Als Konstante in `config.py`:

```python
HANDYS_EXPORT_PATH = (
    r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - "
    r"Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - "
    r"DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3\Daten\Handys"
)
```

> Der Ordner wird beim ersten Export automatisch angelegt (`os.makedirs(..., exist_ok=True)`).

---

## Dateiname

```
Handys_Uebersicht_YYYY-MM-DD.xlsx
```

Beispiel: `Handys_Uebersicht_2025-07-15.xlsx`

Bei mehrfachem Export am selben Tag wird die Datei **überschrieben** (kein Suffix).

---

## Tabellenblatt-Struktur

### Blatt 1: `Übersicht`

#### Kopfzeile (Zeile 1)

Zellbereich `A1:J1` verbunden, Inhalt:

```
Diensthandy-Übersicht – DRK Erste-Hilfe-Station Flughafen Köln/Bonn
Erstellt: DD.MM.YYYY HH:MM Uhr
```

Formatierung: DRK-Rot Hintergrund, Weiß fett, Schriftgröße 12.

#### Spaltenüberschriften (Zeile 2)

| Spalte | Überschrift |
|--------|------------|
| A | Inventarnummer |
| B | Hersteller |
| C | Modell |
| D | Rufnummer |
| E | SIM-Nummer |
| F | Standort / Ausgabe |
| G | Zustand |
| H | Defektbeschreibung |
| I | Anschaffungsdatum |
| J | Zuletzt geändert |

Formatierung: DRK-Dunkelgrau Hintergrund, Weiß fett, Rahmen.

#### Datenzeilen (ab Zeile 3)

- Wechselnde Zeilenfärbung (weiß / hellgrau, wie in `cd.py`)
- Spalte G (`Zustand`): farbige Zellhinterlegung nach Status:

| Zustand | Hintergrundfarbe | Schriftfarbe |
|---------|-----------------|-------------|
| Aktiv | `#C6EFCE` (Hellgrün) | `#276221` |
| Defekt | `#FFC7CE` (Hellrot) | `#9C0006` |
| Außer Betrieb | `#FFEB9C` (Gelb) | `#9C6500` |
| Reserve | `#D9D9D9` (Grau) | `#404040` |

#### Spaltenbreiten

| Spalte | Breite (Zeichen) |
|--------|-----------------|
| A | 16 |
| B | 14 |
| C | 18 |
| D | 16 |
| E | 20 |
| F | 20 |
| G | 14 |
| H | 35 |
| I | 16 |
| J | 20 |

#### Fußzeile (letzte Zeile + 2)

```
Gesamt: XX Geräte  |  Aktiv: X  |  Defekt: X  |  Außer Betrieb: X  |  Reserve: X
```

Berechnung aus den Daten, keine Formeln.

---

### Blatt 2: `Historie`

Exportiert die letzten **90 Tage** der `handys_historie`-Tabelle.

#### Spaltenüberschriften

| Spalte | Überschrift |
|--------|------------|
| A | Datum / Uhrzeit |
| B | Inventarnummer |
| C | Geändertes Feld |
| D | Alter Wert |
| E | Neuer Wert |
| F | Benutzer |

- Keine Bearbeitbarkeit; rein informativ
- Sortierung: neueste Einträge oben

---

## Implementierungshinweise (`handys_excel_export.py`)

```python
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime
from config import HANDYS_EXPORT_PATH
from modules.handys.handys_db import get_all_handys, get_historie
# from cd import DRK_ROT, DRK_DUNKELGRAU, ZEILENFARBE_GERADE, ZEILENFARBE_UNGERADE


ZUSTAND_FARBEN = {
    'Aktiv':           {'bg': 'C6EFCE', 'fg': '276221'},
    'Defekt':          {'bg': 'FFC7CE', 'fg': '9C0006'},
    'Außer Betrieb':   {'bg': 'FFEB9C', 'fg': '9C6500'},
    'Reserve':         {'bg': 'D9D9D9', 'fg': '404040'},
}


def export_handys_excel(open_after: bool = False) -> str:
    """
    Exportiert alle Handys als Excel-Datei in HANDYS_EXPORT_PATH.
    Gibt den vollständigen Dateipfad zurück.
    """
    os.makedirs(HANDYS_EXPORT_PATH, exist_ok=True)

    dateiname = f"Handys_Uebersicht_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    pfad = os.path.join(HANDYS_EXPORT_PATH, dateiname)

    wb = openpyxl.Workbook()

    # --- Blatt 1: Übersicht ---
    ws1 = wb.active
    ws1.title = "Übersicht"
    _erstelle_uebersicht_blatt(ws1)

    # --- Blatt 2: Historie ---
    ws2 = wb.create_sheet("Historie")
    _erstelle_historie_blatt(ws2)

    wb.save(pfad)

    if open_after:
        os.startfile(pfad)   # Windows only

    return pfad


def _erstelle_uebersicht_blatt(ws):
    handys = get_all_handys()
    now_str = datetime.now().strftime('%d.%m.%Y %H:%M Uhr')

    # Kopfzeile
    ws.merge_cells('A1:J1')
    ws['A1'] = (
        f"Diensthandy-Übersicht – DRK Erste-Hilfe-Station Flughafen Köln/Bonn"
        f"\nErstellt: {now_str}"
    )
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=12)
    ws['A1'].fill = PatternFill('solid', fgColor='C0292B')  # DRK-Rot aus cd.py
    ws['A1'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[1].height = 30

    # Spaltenüberschriften
    headers = [
        'Inventarnummer', 'Hersteller', 'Modell', 'Rufnummer', 'SIM-Nummer',
        'Standort / Ausgabe', 'Zustand', 'Defektbeschreibung',
        'Anschaffungsdatum', 'Zuletzt geändert'
    ]
    col_widths = [16, 14, 18, 16, 20, 20, 14, 35, 16, 20]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='404040')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[cell.column_letter].width = width

    # Datenzeilen
    db_felder = [
        'inventarnummer', 'hersteller', 'modell', 'rufnummer', 'sim_nummer',
        'standort', 'zustand', 'defekt_beschreibung', 'anschaffungsdatum', 'geaendert_am'
    ]
    for row_idx, handy in enumerate(handys, start=3):
        fill_bg = 'FFFFFF' if row_idx % 2 == 0 else 'F2F2F2'
        for col_idx, feld in enumerate(db_felder, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=handy[feld] or '')
            cell.fill = PatternFill('solid', fgColor=fill_bg)
            # Zustand-Spalte (G = col 7) farbig hervorheben
            if col_idx == 7:
                zustand = handy['zustand'] or 'Aktiv'
                farbe = ZUSTAND_FARBEN.get(zustand, ZUSTAND_FARBEN['Aktiv'])
                cell.fill = PatternFill('solid', fgColor=farbe['bg'])
                cell.font = Font(color=farbe['fg'], bold=True)

    # Zusammenfassung
    zaehler = {'Aktiv': 0, 'Defekt': 0, 'Außer Betrieb': 0, 'Reserve': 0}
    for h in handys:
        zaehler[h['zustand']] = zaehler.get(h['zustand'], 0) + 1
    summary_row = len(handys) + 4
    ws.cell(row=summary_row, column=1,
            value=(f"Gesamt: {len(handys)} Geräte  |  "
                   f"Aktiv: {zaehler['Aktiv']}  |  "
                   f"Defekt: {zaehler['Defekt']}  |  "
                   f"Außer Betrieb: {zaehler['Außer Betrieb']}  |  "
                   f"Reserve: {zaehler['Reserve']}"))
    ws.merge_cells(f'A{summary_row}:J{summary_row}')
    ws.cell(row=summary_row, column=1).font = Font(bold=True, italic=True)


def _erstelle_historie_blatt(ws):
    from datetime import timedelta
    cutoff = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
    eintraege = [e for e in get_historie() if (e['geaendert_am'] or '') >= cutoff]

    headers = ['Datum / Uhrzeit', 'Inventarnummer', 'Geändertes Feld',
               'Alter Wert', 'Neuer Wert', 'Benutzer']
    col_widths = [20, 18, 22, 30, 30, 16]

    for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='404040')
        ws.column_dimensions[cell.column_letter].width = w

    for row_idx, eintrag in enumerate(reversed(eintraege), start=2):
        ws.cell(row=row_idx, column=1, value=eintrag['geaendert_am'])
        ws.cell(row=row_idx, column=2, value=eintrag['inventarnummer'])
        ws.cell(row=row_idx, column=3, value=eintrag['feld'])
        ws.cell(row=row_idx, column=4, value=eintrag['alter_wert'])
        ws.cell(row=row_idx, column=5, value=eintrag['neuer_wert'])
        ws.cell(row=row_idx, column=6, value=eintrag['benutzer'])
```

---

## Aufruf aus dem Widget

```python
# Button-Handler im HandysWidget
def on_excel_export_clicked(self):
    try:
        pfad = export_handys_excel(open_after=True)
        QMessageBox.information(self, "Export erfolgreich",
                                f"Datei gespeichert:\n{pfad}")
    except Exception as e:
        QMessageBox.critical(self, "Exportfehler", str(e))
```
