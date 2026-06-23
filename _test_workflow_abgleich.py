"""Temporäres Test-Skript – kein Produktionscode, wird nicht committet"""
import sys
sys.path.insert(0, r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\Nesk\Nesk3")

from pathlib import Path
from gui.workflow import _parse_staerkemeldung, _parse_dienstplan_fuer_abgleich, _abgleichen, _datum_aus_dateiname

SM_DIR  = Path(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\06_Stärkemeldung\05_Mai")
DP_DIR  = Path(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\04_Tagesdienstpläne\05_Mai")

sm_files = sorted(SM_DIR.glob("*.docx"))
dp_files = sorted(DP_DIR.glob("*.xlsx"))

print(f"Stärkemeldungen gefunden : {len(sm_files)}")
print(f"Tagesdienstpläne gefunden: {len(dp_files)}")
print()

# ── Daten laden ───────────────────────────────────────────────────────────────
sm_data = [_parse_staerkemeldung(str(f)) for f in sm_files]
dp_data = [_parse_dienstplan_fuer_abgleich(str(f)) for f in dp_files]

# Datum-Index DP aufbauen (genau wie in _LadeThread)
dp_by_datum: dict[str, list[dict]] = {}
for dp in dp_data:
    d = dp.get("datum", "")
    dp_by_datum.setdefault(d, []).append(dp)

matched_dp: set[int] = set()
paarungen: list[tuple[dict, dict]] = []

for sm in sm_data:
    sm_datum = sm.get("datum", "")
    candidates = dp_by_datum.get(sm_datum, []) if sm_datum else []
    unmatch = [c for c in candidates if id(c) not in matched_dp]
    if unmatch:
        dp_match = unmatch[0]
        matched_dp.add(id(dp_match))
    elif not sm_datum:
        dp_match = {"datei": "(kein Datum)", "datum": "", "personen": []}
    else:
        dp_match = {"datei": f"(kein DP für {sm_datum})", "datum": sm_datum, "personen": []}
    paarungen.append((sm, dp_match))

# ÜbrigeDP ohne SM
for dp in dp_data:
    if id(dp) not in matched_dp:
        paarungen.append(({"datei": "–", "datum": "", "personen": []}, dp))

# ── Einzeltest: erster Tag ─────────────────────────────────────────────────────
print("=" * 70)
print(f"PAARUNGEN ({len(paarungen)} Einträge):")
print("=" * 70)
for sm, dp in paarungen[:5]:
    print(f"  SM: {sm['datei']:<50}  DP: {dp['datei']}")
if len(paarungen) > 5:
    print(f"  ... ({len(paarungen)-5} weitere)")

print()
print("=" * 70)
print("EINZELTEST: erste Paarung")
print("=" * 70)
sm1, dp1 = paarungen[0]
print(f"\nStärkemeldung '{sm1['datei']}' – Datum: '{sm1['datum']}'")
print(f"  Personen ({len(sm1['personen'])}):")
for p in sm1["personen"][:5]:
    print(f"    {p['vollname']:<30}  Dienst={p['dienst']:<10}  {p['beginn']}-{p['ende']}")
if len(sm1["personen"]) > 5:
    print(f"    ... ({len(sm1['personen'])-5} weitere)")

print(f"\nDienstplan '{dp1['datei']}' – Datum: '{dp1['datum']}'")
print(f"  Personen ({len(dp1['personen'])}):")
for p in dp1["personen"][:5]:
    print(f"    {p['vollname']:<30}  Dienst={p['dienst']:<6}  {p['beginn']}-{p['ende']}")
if len(dp1["personen"]) > 5:
    print(f"    ... ({len(dp1['personen'])-5} weitere)")

erg1 = _abgleichen(sm1, dp1)
print(f"\n  ✅  Identisch   : {len(erg1.ok)}")
print(f"  ⚠️  Abweichungen: {len(erg1.abweichung)}")
for e in erg1.abweichung[:5]:
    print(f"      {e['name']}: {'; '.join(e['unterschiede'])}")
print(f"  📄  Nur Stärke  : {len(erg1.nur_staerke)}")
print(f"  📋  Nur Dienstpl: {len(erg1.nur_dienstplan)}")

# ── Gesamtabgleich ─────────────────────────────────────────────────────────────
print()
print("=" * 70)
print("GESAMTABGLEICH MAI")
print("=" * 70)

gesamt_ok   = 0
gesamt_abw  = 0
gesamt_ns   = 0
gesamt_nd   = 0
fehler_tage = []
kein_match  = []

for sm, dp in paarungen:
    if not dp.get("personen") and not sm.get("personen"):
        continue
    erg = _abgleichen(sm, dp)
    gesamt_ok  += len(erg.ok)
    gesamt_abw += len(erg.abweichung)
    gesamt_ns  += len(erg.nur_staerke)
    gesamt_nd  += len(erg.nur_dienstplan)
    if "(kein DP" in dp.get("datei", ""):
        kein_match.append(sm["datei"])
    elif erg.hat_fehler:
        fehler_tage.append((sm["datei"], dp["datei"], erg))

print(f"\nGESAMT über {len(paarungen)} Paarungen:")
print(f"  ✅  Identisch   : {gesamt_ok}")
print(f"  ⚠️  Abweichungen: {gesamt_abw}")
print(f"  📄  Nur Stärke  : {gesamt_ns}")
print(f"  📋  Nur Dienstpl: {gesamt_nd}")
print(f"  ❌  Tage mit Diff: {len(fehler_tage)}")

if kein_match:
    print(f"\nSM ohne passenden Dienstplan ({len(kein_match)}):")
    for f in kein_match:
        print(f"  ⚠️  {f}")

print()
print("Tage mit Abweichungen:")
for sm_datei, dp_datei, erg in fehler_tage[:8]:
    print(f"\n  [{sm_datei}  ↔  {dp_datei}]")
    for e in erg.abweichung[:5]:
        print(f"    ⚠️ {e['name']}: {'; '.join(e['unterschiede'])}")
    if len(erg.abweichung) > 5:
        print(f"    ... ({len(erg.abweichung)-5} weitere Abweichungen)")
    for e in erg.nur_staerke[:3]:
        print(f"    📄 NUR STÄRKE  {e['name']} ({e.get('sm_dienst','')})")
    if len(erg.nur_staerke) > 3:
        print(f"    ... ({len(erg.nur_staerke)-3} weitere nur Stärke)")
    for e in erg.nur_dienstplan[:3]:
        print(f"    📋 NUR DIENSTPL {e['name']} ({e.get('dp_dienst','')})")
    if len(erg.nur_dienstplan) > 3:
        print(f"    ... ({len(erg.nur_dienstplan)-3} weitere nur Dienstplan)")

if len(fehler_tage) > 8:
    print(f"\n... und {len(fehler_tage)-8} weitere Tage mit Abweichungen")

from pathlib import Path
from gui.workflow import _parse_staerkemeldung, _parse_dienstplan_fuer_abgleich, _abgleichen

SM_DIR  = Path(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\06_Stärkemeldung\05_Mai")
DP_DIR  = Path(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\04_Tagesdienstpläne\05_Mai")

sm_files = sorted(SM_DIR.glob("*.docx"))
dp_files = sorted(DP_DIR.glob("*.xlsx"))

print(f"Stärkemeldungen gefunden : {len(sm_files)}")
print(f"Tagesdienstpläne gefunden: {len(dp_files)}")
print()

# ── Einzeltest: erster Tag ────────────────────────────────────────────────────
print("=" * 70)
print("EINZELTEST: 01.05.2026")
print("=" * 70)

sm1 = _parse_staerkemeldung(str(sm_files[0]))
dp1 = _parse_dienstplan_fuer_abgleich(str(dp_files[0]))

print(f"\nStärkemeldung '{sm1['datei']}' – Datum: '{sm1['datum']}' – Fehler: {sm1.get('fehler')}")
print(f"  Personen ({len(sm1['personen'])}):")
for p in sm1["personen"]:
    print(f"    {p['vollname']:<30}  Dienst={p['dienst']:<6}  {p['beginn']}-{p['ende']}")

print(f"\nDienstplan '{dp1['datei']}' – Datum: '{dp1['datum']}' – Fehler: {dp1.get('fehler')}")
print(f"  Personen ({len(dp1['personen'])}):")
for p in dp1["personen"]:
    print(f"    {p['vollname']:<30}  Dienst={p['dienst']:<6}  {p['beginn']}-{p['ende']}")

erg1 = _abgleichen(sm1, dp1)
print(f"\n  ✅  Identisch   : {len(erg1.ok)}")
print(f"  ⚠️  Abweichungen: {len(erg1.abweichung)}")
for e in erg1.abweichung:
    print(f"      {e['name']}: {'; '.join(e['unterschiede'])}")
print(f"  📄  Nur Stärke  : {len(erg1.nur_staerke)}")
for e in erg1.nur_staerke:
    print(f"      {e['name']} ({e.get('sm_dienst','')})")
print(f"  📋  Nur Dienstpl: {len(erg1.nur_dienstplan)}")
for e in erg1.nur_dienstplan:
    print(f"      {e['name']} ({e.get('dp_dienst','')})")

# ── Gesamtabgleich aller Mai-Tage ─────────────────────────────────────────────
print()
print("=" * 70)
print("GESAMTABGLEICH MAI (ersten 5 Tage ausführlich, dann Zusammenfassung)")
print("=" * 70)

gesamt_ok   = 0
gesamt_abw  = 0
gesamt_ns   = 0
gesamt_nd   = 0
fehler_tage = []

for sm_f, dp_f in zip(sm_files, dp_files):
    sm = _parse_staerkemeldung(str(sm_f))
    dp = _parse_dienstplan_fuer_abgleich(str(dp_f))
    erg = _abgleichen(sm, dp)

    gesamt_ok  += len(erg.ok)
    gesamt_abw += len(erg.abweichung)
    gesamt_ns  += len(erg.nur_staerke)
    gesamt_nd  += len(erg.nur_dienstplan)

    # Tage mit Abweichungen merken
    if erg.hat_fehler:
        fehler_tage.append((sm_f.name, erg))

print()
print(f"GESAMT über {len(sm_files)} Tage:")
print(f"  ✅  Identisch   : {gesamt_ok}")
print(f"  ⚠️  Abweichungen: {gesamt_abw}")
print(f"  📄  Nur Stärke  : {gesamt_ns}")
print(f"  📋  Nur Dienstpl: {gesamt_nd}")
print(f"  ❌  Tage mit Diff: {len(fehler_tage)}")

print()
print("Tage mit Abweichungen / fehlenden Einträgen:")
for tag, erg in fehler_tage[:10]:  # max 10 anzeigen
    print(f"\n  [{tag}]")
    for e in erg.abweichung:
        print(f"    ⚠️ ABWEICHUNG {e['name']}: {'; '.join(e['unterschiede'])}")
    for e in erg.nur_staerke:
        print(f"    📄 NUR STÄRKE  {e['name']} (Dienst={e.get('sm_dienst','')})")
    for e in erg.nur_dienstplan:
        print(f"    📋 NUR DIENSTPL {e['name']} (Dienst={e.get('dp_dienst','')})")

if len(fehler_tage) > 10:
    print(f"\n  … und {len(fehler_tage)-10} weitere Tage mit Abweichungen")
