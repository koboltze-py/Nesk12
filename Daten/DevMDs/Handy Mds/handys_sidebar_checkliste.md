# Nesk3 – Sidebar-Änderung & Integrations-Checkliste

## Sidebar: Call Transcription → Handys

### Zu entfernen

Suche in der Navigation / Sidebar-Konfiguration nach folgendem Eintrag
und **entferne ihn vollständig**:

```python
# Beispiel – exakter Code je nach Implementierung anpassen:
{"label": "Call Transcription", "icon": ..., "widget": CallTranscriptionWidget}
```

Betroffene Dateien (typisch in Nesk3):
- `sidebar.py` oder `navigation.py`
- `main_window.py` (Widget-Instanziierung)
- ggf. `__init__.py` des Call-Transcription-Moduls

> Die Modul-Dateien des Call-Transcription-Features können **archiviert** (nicht gelöscht)
> werden, falls eine spätere Reaktivierung gewünscht ist.

---

### Hinzuzufügen

```python
# In der Sidebar-Konfiguration (an der gleichen Position wie "Call Transcription"):
{"label": "Handys", "icon": ph.icons.PHONE, "widget": HandysWidget}

# In main_window.py:
from modules.handys.handys_widget import HandysWidget
```

---

## Integrations-Checkliste

### Datenbank

- [ ] `handys.db` Initialisierungsfunktion aufrufen beim App-Start (`init_db()`)
- [ ] `handys.db` in `BackupManager`-Dateiliste aufnehmen
- [ ] Sicherstellen, dass `Daten/`-Ordner in `.gitignore` (DB-Datei nicht ins Repo)

### Konfiguration (`config.py`)

- [ ] `HANDYS_EXPORT_PATH` Konstante hinzufügen
- [ ] `HANDYS_EMAIL_EMPFAENGER` Konstante hinzufügen
- [ ] `HANDYS_EMAIL_ABSENDER` Konstante hinzufügen (oder aus Benutzereinstellungen)

### Neue Dateien anlegen

- [ ] `modules/handys/__init__.py`
- [ ] `modules/handys/handys_widget.py` – Haupt-Widget (3 Tabs)
- [ ] `modules/handys/handys_db.py` – CRUD + Historien-Logik
- [ ] `modules/handys/handys_excel_export.py` – Excel-Exportfunktion
- [ ] `modules/handys/handys_email.py` – Outlook-E-Mail-Funktion

### Sidebar / Navigation

- [ ] `Call Transcription` Eintrag entfernen
- [ ] `Handys` Eintrag an gleicher Position hinzufügen
- [ ] Icon prüfen / anpassen (PySide6-kompatibel)

### Widget

- [ ] Tab 1 – Übersichtstabelle mit Statusfarben
- [ ] Tab 2 – Formular Anlegen / Bearbeiten mit Pflichtfeldvalidierung
- [ ] Tab 3 – Historie (read-only, filterbar)
- [ ] Button „+ Neues Handy"
- [ ] Button „✎ Bearbeiten"
- [ ] Button „✕ Löschen" (mit Bestätigungsdialog)
- [ ] Button „⟳ Aktualisieren"
- [ ] Button „📤 Excel exportieren"
- [ ] Button „📧 Per E-Mail senden"

### Excel-Export

- [ ] Exportordner wird automatisch erstellt (`os.makedirs(..., exist_ok=True)`)
- [ ] Blatt 1 „Übersicht" mit Status-Farbcodierung
- [ ] Blatt 2 „Historie" mit 90-Tage-Filter
- [ ] Fußzeile mit Geräte-Zusammenfassung
- [ ] DRK Corporate Design aus `cd.py` angewendet

### E-Mail

- [ ] `pywin32` installiert auf allen Nesk3-PCs: `pip install pywin32`
- [ ] Outlook auf Arbeits-PCs verfügbar und eingerichtet
- [ ] Fallback-Dialog bei fehlendem Outlook / pywin32 implementiert
- [ ] Excel wird **immer** gespeichert, auch wenn E-Mail-Versand fehlschlägt

### Abhängigkeiten

```
pywin32       # Outlook MAPI
openpyxl      # Excel (bereits vorhanden)
```

---

## Empfohlene Implementierungsreihenfolge

1. **DB-Schicht** (`handys_db.py`) – CRUD + `init_db()` + Historien-Trigger
2. **Backup-Integration** – `handys.db` in BackupManager eintragen
3. **Basis-Widget** (`handys_widget.py`) – Tab 1 Übersicht + Tab 2 Formular
4. **Historien-Tab** (`handys_widget.py`) – Tab 3 lesen und anzeigen
5. **Excel-Export** (`handys_excel_export.py`)
6. **E-Mail-Funktion** (`handys_email.py`)
7. **Sidebar** – Call Transcription raus, Handys rein
8. **Test** auf allen Nesk3-PCs (inkl. Turso-Sync prüfen, falls Handys-DB ebenfalls synchronisiert werden soll)

---

## Hinweis: Turso-Sync für `handys.db`

Aktuell ist unklar, ob `handys.db` ebenfalls über Turso (Cloud-Sync) zwischen
den ~4 Nesk3-Workstations synchronisiert werden soll.

**Empfehlung:** Ja – gleiche Sync-Strategie wie `nesk3.db`, damit Änderungen
an einem PC sofort auf allen anderen sichtbar sind.

Falls ja: in `turso_sync.py` (o. ä.) die `handys.db` analog zur Haupt-DB
als embedded replica registrieren.
