"""Tiefe Analyse: Dienstplan-Parser Keys + Stärkemeldung Struktur"""
import sys
sys.path.insert(0, '.')
from functions.dienstplan_parser import DienstplanParser

pfad = r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\04_Tagesdienstpläne\05_Mai\01.05.2026.xlsx"
result = DienstplanParser(pfad, alle_anzeigen=True, round_dispo=False).parse()

print("=== DIENSTPLAN KEYS (parse() result) ===")
print(f"  Top-Level Keys: {list(result.keys())}")
print(f"  betreuer: {len(result.get('betreuer', []))} Einträge")
print(f"  dispo:    {len(result.get('dispo', []))} Einträge")

if result.get('betreuer'):
    p = result['betreuer'][0]
    print(f"\n  Erster Betreuer – alle Keys: {list(p.keys())}")
    print(f"  {p}")

if result.get('dispo'):
    p = result['dispo'][0]
    print(f"\n  Erster Dispo – alle Keys: {list(p.keys())}")
    print(f"  {p}")

print()
print("=== ERSTE 5 BETREUER ===")
for p in result.get('betreuer', [])[:5]:
    print(f"  {p.get('vollname','?'):<30}  dienst={p.get('dienst','')!r}  start={p.get('start_zeit','')!r}  end={p.get('end_zeit','')!r}")

print()
print("=== ERSTE 5 DISPO ===")
for p in result.get('dispo', [])[:5]:
    print(f"  {p.get('vollname','?'):<30}  dienst={p.get('dienst','')!r}  start={p.get('start_zeit','')!r}  end={p.get('end_zeit','')!r}")
