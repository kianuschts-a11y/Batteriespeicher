# check_battery_params.py
# Führe dieses Skript aus, um die geladenen Werte zu sehen

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_import import load_battery_tech_params
from config import DEFAULT_BATTERY_MAX_CHARGE_KW, DEFAULT_BATTERY_MAX_DISCHARGE_KW

# Lade die technischen Parameter
print("Lade technische Parameter aus Batteriespeicherkosten.xlsm...")
print("-" * 70)
battery_tech_params = load_battery_tech_params("Batteriespeicherkosten.xlsm")

if battery_tech_params is None:
    print("FEHLER: Konnte Parameter nicht laden!")
    sys.exit(1)

# Funktion zum Abrufen der Werte für eine Kapazität (wie im Code verwendet)
def resolve_power_for_capacity(capacity_kwh: float, battery_tech_params: dict):
    cap_int = int(round(capacity_kwh))
    if battery_tech_params and isinstance(battery_tech_params, dict):
        # exakte Kapazität
        if cap_int in battery_tech_params:
            m = battery_tech_params[cap_int]
            return (
                float(m.get("max_charge_kw", DEFAULT_BATTERY_MAX_CHARGE_KW)),
                float(m.get("max_discharge_kw", DEFAULT_BATTERY_MAX_DISCHARGE_KW)),
                cap_int  # exakte Kapazität gefunden
            )
        # nächstliegende Kapazität suchen
        keys = sorted(battery_tech_params.keys())
        if keys:
            nearest = min(keys, key=lambda k: abs(k - cap_int))
            m = battery_tech_params.get(nearest, {})
            nearest_cap = nearest
            return (
                float(m.get("max_charge_kw", DEFAULT_BATTERY_MAX_CHARGE_KW)),
                float(m.get("max_discharge_kw", DEFAULT_BATTERY_MAX_DISCHARGE_KW)),
                nearest_cap
            )
    # Fallback auf feste Default-Werte
    return (DEFAULT_BATTERY_MAX_CHARGE_KW, DEFAULT_BATTERY_MAX_DISCHARGE_KW, None)

print("\n" + "=" * 70)
print("LADE- UND ENTLADELEISTUNGEN FÜR KAPAZITÄTEN 1-10 KWH")
print("=" * 70)
print(f"{'Kapazität':<12} {'Max. Ladeleistung':<20} {'Max. Entladeleistung':<25} {'Quelle'}")
print("-" * 70)

for capacity in range(1, 11):
    result = resolve_power_for_capacity(float(capacity), battery_tech_params)
    charge_kw, discharge_kw, nearest = result
    
    if nearest is None:
        source = "(Fallback: Default-Werte)"
    elif nearest == capacity:
        source = "(exakt in Excel)"
    else:
        source = f"(nächstliegende: {nearest} kWh)"
    
    print(f"{capacity:>2} kWh       {charge_kw:>6.2f} kW              {discharge_kw:>6.2f} kW                 {source}")

print("\n" + "-" * 70)
print(f"Standard-Fallback-Werte (falls nicht in Excel gefunden):")
print(f"  Ladeleistung: {DEFAULT_BATTERY_MAX_CHARGE_KW} kW")
print(f"  Entladeleistung: {DEFAULT_BATTERY_MAX_DISCHARGE_KW} kW")
print("\n" + "=" * 70)

# Zeige auch alle verfügbaren Kapazitäten aus der Excel-Datei
print("\nAlle verfügbaren Kapazitäten in der Excel-Datei:")
print("-" * 70)
sorted_caps = sorted(battery_tech_params.keys())
for cap in sorted_caps[:20]:  # Erste 20 anzeigen
    params = battery_tech_params[cap]
    print(f"{cap:>3} kWh: Lade={params.get('max_charge_kw', 'N/A'):>6.2f} kW, "
          f"Entlade={params.get('max_discharge_kw', 'N/A'):>6.2f} kW")
if len(sorted_caps) > 20:
    print(f"... und {len(sorted_caps) - 20} weitere")

