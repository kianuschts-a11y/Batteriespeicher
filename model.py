import numpy as np
import pandas as pd

def get_supported_data_resolutions():
    """
    Gibt die unterstützten Datenauflösungen zurück.
    
    Returns:
        dict: Dictionary mit unterstützten Auflösungen und deren Periodenanzahl
    """
    return {
        "1h": {"periods_normal": 8760, "periods_leap": 8784, "interval_hours": 1.0},
        "30min": {"periods_normal": 17520, "periods_leap": 17568, "interval_hours": 0.5},
        "15min": {"periods_normal": 35040, "periods_leap": 35136, "interval_hours": 0.25},
        "10min": {"periods_normal": 52560, "periods_leap": 52704, "interval_hours": 1.0/6.0},
        "5min": {"periods_normal": 105120, "periods_leap": 105408, "interval_hours": 1.0/12.0},
        "1min": {"periods_normal": 525600, "periods_leap": 527040, "interval_hours": 1.0/60.0}
    }

def validate_energy_balance(results_df, battery_efficiency_charge, battery_efficiency_discharge, tolerance_percent=0.01):
    """
    Validiert die Energiebilanz einer Simulation.
    Prüft ob Energieerhaltung gilt und Verluste plausibel sind.
    
    Args:
        results_df (pd.DataFrame): DataFrame mit Simulationsergebnissen
        battery_efficiency_charge (float): Lade-Wirkungsgrad (0-1)
        battery_efficiency_discharge (float): Entlade-Wirkungsgrad (0-1)
        tolerance_percent (float): Maximale erlaubte Abweichung in %
    
    Returns:
        dict: Validierungsergebnisse mit Status und Fehlerwerten
    """
    
    # 1. PV-BILANZ: PV = Direktverbrauch + Ladung + Einspeisung
    pv_total = results_df['PV_Generation_kWh'].sum()
    pv_used = (results_df['Direct_Self_Consumption_kWh'].sum() + 
               results_df['Battery_Charge_kWh'].sum() + 
               results_df['Grid_Export_kWh'].sum())
    
    pv_balance_error = abs(pv_total - pv_used) / pv_total * 100 if pv_total > 0 else 0
    
    # 2. VERBRAUCHER-BILANZ: Verbrauch = Direktverbrauch + Batterieentladung + Netzbezug
    consumption_total = results_df['Consumption_kWh'].sum()
    consumption_covered = (results_df['Direct_Self_Consumption_kWh'].sum() + 
                          results_df['Battery_Discharge_kWh'].sum() + 
                          results_df['Grid_Import_kWh'].sum())
    
    consumption_balance_error = abs(consumption_total - consumption_covered) / consumption_total * 100 if consumption_total > 0 else 0
    
    # 3. BATTERIE-VERLUSTE: Plausibilitätsprüfung
    total_charge = results_df['Battery_Charge_kWh'].sum()
    total_discharge_netto = results_df['Battery_Discharge_kWh'].sum()
    total_charge_losses = results_df['Battery_Charge_Losses_kWh'].sum()
    total_discharge_losses = results_df['Battery_Discharge_Losses_kWh'].sum()
    
    # Erwartete Verluste berechnen
    expected_charge_losses = total_charge * (1 - battery_efficiency_charge)
    # Bei Entladung: Netto-Entladung / η gibt Brutto-Entladung, Verluste = Brutto - Netto
    expected_discharge_brutto = total_discharge_netto / battery_efficiency_discharge if battery_efficiency_discharge > 0 else 0
    expected_discharge_losses = expected_discharge_brutto - total_discharge_netto
    
    # Abweichungen berechnen
    charge_loss_error = abs(total_charge_losses - expected_charge_losses) / expected_charge_losses * 100 if expected_charge_losses > 0 else 0
    discharge_loss_error = abs(total_discharge_losses - expected_discharge_losses) / expected_discharge_losses * 100 if expected_discharge_losses > 0 else 0
    
    # Validierung
    pv_ok = pv_balance_error <= tolerance_percent
    consumption_ok = consumption_balance_error <= tolerance_percent
    charge_loss_ok = charge_loss_error <= tolerance_percent or total_charge == 0
    discharge_loss_ok = discharge_loss_error <= tolerance_percent or total_discharge_netto == 0
    
    all_ok = pv_ok and consumption_ok and charge_loss_ok and discharge_loss_ok
    
    # Ausgabe bei Problemen
    if not pv_ok:
        print(f"⚠️ WARNUNG: PV-Energiebilanz nicht ausgeglichen!")
        print(f"  PV gesamt: {pv_total:.2f} kWh")
        print(f"  PV verwendet: {pv_used:.2f} kWh")
        print(f"  Fehler: {pv_balance_error:.4f}%")
    
    if not consumption_ok:
        print(f"⚠️ WARNUNG: Verbraucher-Energiebilanz nicht ausgeglichen!")
        print(f"  Verbrauch gesamt: {consumption_total:.2f} kWh")
        print(f"  Verbrauch gedeckt: {consumption_covered:.2f} kWh")
        print(f"  Fehler: {consumption_balance_error:.4f}%")
    
    if not charge_loss_ok and total_charge > 0:
        print(f"⚠️ WARNUNG: Lade-Verluste unplausibel!")
        print(f"  Tatsächlich: {total_charge_losses:.2f} kWh")
        print(f"  Erwartet: {expected_charge_losses:.2f} kWh")
        print(f"  Fehler: {charge_loss_error:.4f}%")
    
    if not discharge_loss_ok and total_discharge_netto > 0:
        print(f"⚠️ WARNUNG: Entlade-Verluste unplausibel!")
        print(f"  Tatsächlich: {total_discharge_losses:.2f} kWh")
        print(f"  Erwartet: {expected_discharge_losses:.2f} kWh")
        print(f"  Fehler: {discharge_loss_error:.4f}%")
    
    if all_ok:
        # Optional: Debug-Ausgabe bei Erfolg (kann auskommentiert werden bei Bedarf)
        print(f"✅ Energiebilanz-Validierung OK (PV: {pv_balance_error:.4f}%, Verbrauch: {consumption_balance_error:.4f}%)")
    
    return {
        'all_ok': all_ok,
        'pv_balance_ok': pv_ok,
        'consumption_balance_ok': consumption_ok,
        'charge_loss_ok': charge_loss_ok,
        'discharge_loss_ok': discharge_loss_ok,
        'pv_balance_error_percent': pv_balance_error,
        'consumption_balance_error_percent': consumption_balance_error,
        'charge_loss_error_percent': charge_loss_error,
        'discharge_loss_error_percent': discharge_loss_error
    }

def simulate_one_year(
    consumption_series: pd.Series, 
    pv_generation_series: pd.Series, 
    battery_capacity_kwh: float, 
    battery_efficiency_charge: float, # Wirkungsgrad beim Laden
    battery_efficiency_discharge: float, # Wirkungsgrad beim Entladen
    battery_max_charge_kw: float, # Maximale Ladeleistung in kW
    battery_max_discharge_kw: float, # Maximale Entladeleistung in kW
    price_grid_per_kwh: float, # Preis für Netzbezug in Euro/kWh (kann auch ein pd.Series für variable Tarife sein)
    price_feed_in_per_kwh: float, # Preis für Netzeinspeisung in Euro/kWh (kann auch ein pd.Series für variable Tarife sein)
    initial_soc_percent: float = 50.0, # Anfangsladezustand der Batterie in %
    min_soc_percent: float = 10.0, # Minimaler Ladezustand in %
    max_soc_percent: float = 90.0, # Maximaler Ladezustand in %
    annual_capacity_loss_percent: float = 2.0, # Jährlicher Kapazitätsverlust in %
    simulation_year: int = 1 # Jahr der Simulation (für Kapazitätsalterung)
) -> dict:
    """
    Simuliert die Energieflüsse für ein Jahr mit automatischer Erkennung der Datenauflösung.

    Args:
        consumption_series (pd.Series): Stromverbrauch in kWh (1h, 30min, 15min, 10min, 5min oder 1min).
        pv_generation_series (pd.Series): PV-Erzeugung in kWh (1h, 30min, 15min, 10min, 5min oder 1min).
        battery_capacity_kwh (float): Speicherkapazität der Batterie in kWh.
        battery_efficiency_charge (float): Wirkungsgrad der Batterie beim Laden (0-1).
        battery_efficiency_discharge (float): Wirkungsgrad der Batterie beim Entladen (0-1).
        battery_max_charge_kw (float): Maximale Ladeleistung der Batterie in kW.
        battery_max_discharge_kw (float): Maximale Entladeleistung der Batterie in kW.
        price_grid_per_kwh (float | pd.Series): Preis für Netzbezug in Euro/kWh. Kann eine Konstante oder eine Zeitreihe sein.
        price_feed_in_per_kwh (float | pd.Series): Preis für Netzeinspeisung in Euro/kWh. Kann eine Konstante oder eine Zeitreihe sein.
        initial_soc_percent (float): Anfangsladezustand der Batterie in % (0-100).
        min_soc_percent (float): Minimaler Ladezustand der Batterie in % (0-100).
        max_soc_percent (float): Maximaler Ladezustand der Batterie in % (0-100).
        annual_capacity_loss_percent (float): Jährlicher Kapazitätsverlust in % (0-10).
        simulation_year (int): Jahr der Simulation (für Kapazitätsalterung).

    Returns:
        dict: Ein Dictionary mit Zeitreihen der Energieflüsse, KPIs und Simulationsmetadaten.
        Die Lade-/Entladegeschwindigkeiten werden automatisch basierend auf der Datenauflösung angepasst.
    """

    num_periods = len(consumption_series)
    
    # Bestimme die Zeitauflösung basierend auf der Länge der Daten
    # Unterstützte Auflösungen: 1h, 30min, 15min, 10min, 5min, 1min
    if num_periods == 8760:  # Stündliche Daten (normales Jahr)
        time_interval_hours = 1.0
        data_resolution = "hourly"
    elif num_periods == 8784:  # Stündliche Daten (Schaltjahr)
        time_interval_hours = 1.0
        data_resolution = "hourly_leap_year"
    elif num_periods == 17520:  # 30-Minuten-Daten (normales Jahr)
        time_interval_hours = 0.5
        data_resolution = "30min"
    elif num_periods == 17568:  # 30-Minuten-Daten (Schaltjahr)
        time_interval_hours = 0.5
        data_resolution = "30min_leap_year"
    elif num_periods == 35040:  # 15-Minuten-Daten (normales Jahr)
        time_interval_hours = 0.25
        data_resolution = "15min"
    elif num_periods == 35136:  # 15-Minuten-Daten (Schaltjahr)
        time_interval_hours = 0.25
        data_resolution = "15min_leap_year"
    elif num_periods == 52560:  # 10-Minuten-Daten (normales Jahr)
        time_interval_hours = 1.0/6.0  # 10 Minuten = 1/6 Stunde
        data_resolution = "10min"
    elif num_periods == 52704:  # 10-Minuten-Daten (Schaltjahr)
        time_interval_hours = 1.0/6.0
        data_resolution = "10min_leap_year"
    elif num_periods == 105120:  # 5-Minuten-Daten (normales Jahr)
        time_interval_hours = 1.0/12.0  # 5 Minuten = 1/12 Stunde
        data_resolution = "5min"
    elif num_periods == 105408:  # 5-Minuten-Daten (Schaltjahr)
        time_interval_hours = 1.0/12.0
        data_resolution = "5min_leap_year"
    elif num_periods == 525600:  # 1-Minuten-Daten (normales Jahr)
        time_interval_hours = 1.0/60.0  # 1 Minute = 1/60 Stunde
        data_resolution = "1min"
    elif num_periods == 527040:  # 1-Minuten-Daten (Schaltjahr)
        time_interval_hours = 1.0/60.0
        data_resolution = "1min_leap_year"
    else:
        # ERFORDERT: Bekannte Datenauflösung für korrekte Simulation
        raise ValueError(f"Unbekannte Datenauflösung: {num_periods} Perioden. "
                       f"Unterstützte Auflösungen: 1h (8760), 30min (17520), 15min (35040), "
                       f"10min (52560), 5min (105120), 1min (525600), 1min Schaltjahr (527040). "
                       f"Bitte überprüfen Sie Ihre Zeitreihen-Daten.")

    # Initialisierung der Zeitreihen für die Ergebnisse
    soc_series = np.zeros(num_periods) # State of Charge
    grid_import_series = np.zeros(num_periods) # Netzbezug
    grid_export_series = np.zeros(num_periods) # Netzeinspeisung
    battery_charge_series = np.zeros(num_periods) # Ladung der Batterie (Brutto - inkl. Verluste)
    battery_discharge_series = np.zeros(num_periods) # Entladung der Batterie (Netto - nach Verlusten)
    battery_charge_losses_series = np.zeros(num_periods) # Ladeverluste
    battery_discharge_losses_series = np.zeros(num_periods) # Entladeverluste
    direct_self_consumption_series = np.zeros(num_periods) # Direkter Eigenverbrauch

    # Kapazitätsalterung berechnen
    capacity_loss_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (simulation_year - 1)
    current_battery_capacity_kwh = battery_capacity_kwh * capacity_loss_factor
    
    # Aktueller Ladezustand der Batterie (in kWh) - begrenzt auf erlaubten Bereich
    initial_soc_kwh = (initial_soc_percent / 100.0) * current_battery_capacity_kwh
    min_soc_kwh = (min_soc_percent / 100.0) * current_battery_capacity_kwh
    max_soc_kwh = (max_soc_percent / 100.0) * current_battery_capacity_kwh
    current_soc_kwh = np.clip(initial_soc_kwh, min_soc_kwh, max_soc_kwh)

    for i in range(num_periods):
        pv_gen = pv_generation_series.iloc[i]
        consumption = consumption_series.iloc[i]

        # 1. Direkter Eigenverbrauch (PV direkt an Last)
        direct_self_consumption = min(pv_gen, consumption)
        direct_self_consumption_series[i] = direct_self_consumption  # Hier fehlte die Speicherung!
        remaining_pv = pv_gen - direct_self_consumption
        remaining_consumption = consumption - direct_self_consumption

        # 2. Überschüssige PV-Energie in Batterie laden
        if remaining_pv > 0:
            # Maximal mögliche Ladung basierend auf SOC-Grenzen und max. Ladeleistung
            # WICHTIG: Berücksichtige die Zeitauflösung für korrekte Ladegeschwindigkeit
            max_charge_possible_kwh = min(max_soc_kwh - current_soc_kwh, battery_max_charge_kw * time_interval_hours)
            charge_from_pv = min(remaining_pv, max_charge_possible_kwh)
            
            # Berechnung: Verluste beim Laden
            charge_to_battery = charge_from_pv * battery_efficiency_charge
            charge_losses = charge_from_pv - charge_to_battery
            current_soc_kwh += charge_to_battery
            battery_charge_series[i] = charge_from_pv  # Brutto-Ladung (inkl. Verluste)
            battery_charge_losses_series[i] = charge_losses  # Ladeverluste
            remaining_pv -= charge_from_pv

        # 3. Fehlende Energie aus Batterie entladen oder aus Netz beziehen
        # KORREKTUR 2025-11-03: Batterie-Entladungslogik überarbeitet
        # Problem (alt): Division durch η und Multiplikation mit η hoben sich mathematisch auf
        # Lösung (neu): Physikalisch korrekte Implementierung - erst maximale Entladung berechnen,
        #               dann Verluste anwenden, dann auf Bedarf begrenzen
        if remaining_consumption > 0:
            # SCHRITT 1: Berechne maximale Entladung aus Batterie (physikalisch möglich)
            # Berücksichtige: SOC-Grenzen und max. Entladeleistung
            # WICHTIG: Berücksichtige die Zeitauflösung für korrekte Entladegeschwindigkeit
            max_discharge_from_battery_kwh = min(
                current_soc_kwh - min_soc_kwh,  # SOC-Grenze
                battery_max_discharge_kw * time_interval_hours  # Leistungsgrenze
            )
            
            # SCHRITT 2: Berechne wieviel beim Verbraucher ankommt (nach Verlusten)
            # Aus Batterie entnommen → beim Verbraucher kommt weniger an (Wirkungsgrad!)
            max_discharge_to_consumption_kwh = max_discharge_from_battery_kwh * battery_efficiency_discharge
            
            # SCHRITT 3: Wie viel wird tatsächlich verbraucht?
            # Nicht mehr als benötigt!
            actual_discharge_to_consumption = min(max_discharge_to_consumption_kwh, remaining_consumption)
            
            # SCHRITT 4: Rückrechnung: Wieviel wurde tatsächlich aus der Batterie entnommen?
            # Bei Wirkungsgrad 95%: Um X kWh zu liefern, muss X/0.95 aus Batterie entnommen werden
            actual_discharge_from_battery = actual_discharge_to_consumption / battery_efficiency_discharge
            
            # SCHRITT 5: Berechne Verluste
            discharge_losses = actual_discharge_from_battery - actual_discharge_to_consumption
            
            # SCHRITT 6: Aktualisiere Zustandsvariablen
            current_soc_kwh -= actual_discharge_from_battery
            battery_discharge_series[i] = actual_discharge_to_consumption  # Netto-Entladung (nach Verlusten)
            battery_discharge_losses_series[i] = discharge_losses  # Entladeverluste
            remaining_consumption -= actual_discharge_to_consumption

        # 4. Restlicher PV-Überschuss ins Netz einspeisen
        if remaining_pv > 0:
            grid_export_series[i] = remaining_pv

        # 5. Restlicher Verbrauch aus dem Netz beziehen
        if remaining_consumption > 0:
            grid_import_series[i] = remaining_consumption

        # Sicherstellen, dass SoC innerhalb der erlaubten Grenzen bleibt (kleine Rundungsfehler)
        current_soc_kwh = np.clip(current_soc_kwh, min_soc_kwh, max_soc_kwh)
        soc_series[i] = current_soc_kwh

    # Erstellung eines Pandas DataFrame für die Zeitreihen-Ergebnisse
    time_index = consumption_series.index
    results_df = pd.DataFrame({
        'PV_Generation_kWh': pv_generation_series.values,
        'Consumption_kWh': consumption_series.values,
        'SOC_kWh': soc_series,
        'Battery_Charge_kWh': battery_charge_series,
        'Battery_Discharge_kWh': battery_discharge_series,
        'Battery_Charge_Losses_kWh': battery_charge_losses_series,
        'Battery_Discharge_Losses_kWh': battery_discharge_losses_series,
        'Grid_Import_kWh': grid_import_series,
        'Grid_Export_kWh': grid_export_series,
        'Direct_Self_Consumption_kWh': direct_self_consumption_series
    }, index=time_index)

    # Berechnung einfacher KPIs (können später in analysis.py verfeinert werden)
    total_pv_generation = results_df['PV_Generation_kWh'].sum()
    total_consumption = results_df['Consumption_kWh'].sum()
    total_grid_import = results_df['Grid_Import_kWh'].sum()
    total_grid_export = results_df['Grid_Export_kWh'].sum()
    total_direct_self_consumption = results_df['Direct_Self_Consumption_kWh'].sum()
    total_battery_charge = results_df['Battery_Charge_kWh'].sum()
    total_battery_discharge = results_df['Battery_Discharge_kWh'].sum()
    total_battery_charge_losses = results_df['Battery_Charge_Losses_kWh'].sum()
    total_battery_discharge_losses = results_df['Battery_Discharge_Losses_kWh'].sum()
    
    # Kapazitätsalterung KPIs
    original_capacity_kwh = battery_capacity_kwh
    current_capacity_kwh = current_battery_capacity_kwh
    if original_capacity_kwh > 0:
        capacity_loss_percent = ((original_capacity_kwh - current_capacity_kwh) / original_capacity_kwh) * 100
    else:
        capacity_loss_percent = 0.0  # Keine Kapazität = kein Verlust

    # Autarkiegrad: Anteil des Verbrauchs, der nicht aus dem Netz bezogen wird
    autarky_rate = (total_consumption - total_grid_import) / total_consumption if total_consumption > 0 else 0

    # Eigenverbrauchsquote: Anteil der PV-Erzeugung, der selbst verbraucht wird
    self_consumption_rate = (total_direct_self_consumption + total_battery_discharge) / total_pv_generation if total_pv_generation > 0 else 0

    # Kosten und Ersparnisse
    if isinstance(price_grid_per_kwh, pd.Series):
        total_import_cost = (results_df['Grid_Import_kWh'] * price_grid_per_kwh).sum()
    else:
        total_import_cost = total_grid_import * price_grid_per_kwh

    if isinstance(price_feed_in_per_kwh, pd.Series):
        total_export_revenue = (results_df['Grid_Export_kWh'] * price_feed_in_per_kwh).sum()
    else:
        total_export_revenue = total_grid_export * price_feed_in_per_kwh

    # Jährliche Stromkosten mit Speicher
    effective_annual_energy_cost = total_import_cost - total_export_revenue
    
    # HINWEIS 2025-11-03: Referenzkosten OHNE Batterie werden NICHT hier berechnet!
    # Grund: Bei variablen Tarifen wäre Durchschnittspreisberechnung (.mean()) ungenau.
    # Die zeitliche Verteilung der Tarife würde ignoriert (z.B. wann PV einspeist, wann Strom gekauft wird).
    # Korrekte Berechnung erfolgt in analysis.py durch echte Simulation ohne Batterie.
    # Separation of Concerns: model.py simuliert, analysis.py bewertet wirtschaftlich.
    
    # Berechne KPIs für die finanzielle Analyse
    # Verwende die korrekte Netto-Ersparnis für die KPI-Berechnung
    # Die finanziellen Parameter werden später in der UI überschrieben
    financial_kpis = {
        'payback_period_years': np.nan,  # Wird später in der UI berechnet
        'irr_percentage': np.nan  # Wird später in der UI berechnet
    }

    # Validiere Energiebilanz (Qualitätssicherung)
    energy_balance_validation = validate_energy_balance(
        results_df, 
        battery_efficiency_charge, 
        battery_efficiency_discharge,
        tolerance_percent=0.01  # 0.01% Toleranz für Rundungsfehler
    )

    return {
        'time_series_data': results_df,
        'kpis': {
            'autarky_rate': autarky_rate,
            'self_consumption_rate': self_consumption_rate,
            'total_grid_import_kwh': total_grid_import,
            'total_grid_export_kwh': total_grid_export,
            'annual_energy_cost': effective_annual_energy_cost,  # Rückwärtskompatibilität
            'effective_annual_energy_cost': effective_annual_energy_cost,
            'grid_import_cost': total_import_cost,
            'grid_export_revenue': total_export_revenue,
            # 'annual_energy_cost_savings' entfernt - wird korrekt in analysis.py berechnet (2025-11-03)
            'payback_period_years': financial_kpis['payback_period_years'],
            'irr_percentage': financial_kpis['irr_percentage'],
            'total_pv_generation_kwh': total_pv_generation,
            'total_consumption_kwh': total_consumption,
            'total_direct_self_consumption_kwh': total_direct_self_consumption,
            'total_battery_charge_kwh': total_battery_charge,
            'total_battery_discharge_kwh': total_battery_discharge,
            'total_battery_charge_losses_kwh': total_battery_charge_losses,
            'total_battery_discharge_losses_kwh': total_battery_discharge_losses,
            'battery_efficiency_charge': battery_efficiency_charge,
            'battery_efficiency_discharge': battery_efficiency_discharge,
            'original_capacity_kwh': original_capacity_kwh,
            'current_capacity_kwh': current_capacity_kwh,
            'capacity_loss_percent': capacity_loss_percent,
            'simulation_year': simulation_year
        },
        'simulation_metadata': {
            'data_resolution': data_resolution,
            'time_interval_hours': time_interval_hours,
            'num_periods': num_periods,
            'battery_max_charge_kw': battery_max_charge_kw,
            'battery_max_discharge_kw': battery_max_discharge_kw,
            'energy_balance_validation': energy_balance_validation
        }
    } 