import pandas as pd
import numpy as np
from model import simulate_one_year
from analysis import calculate_financial_kpis
from config import DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT

def find_optimal_size(
    consumption_series: pd.Series,
    pv_generation_series: pd.Series,
    min_capacity_kwh: float,
    max_capacity_kwh: float,
    step_kwh: float,
    battery_efficiency_charge: float,
    battery_efficiency_discharge: float,
    battery_max_charge_kw: float,
    battery_max_discharge_kw: float,
    price_grid_per_kwh,
    price_feed_in_per_kwh,
    battery_cost_curve: dict,  # Dictionary mit Kostenkurven-Daten
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05, # Diskontierungsrate für NPV
    initial_soc_percent: float = 50.0, # Anfangsladezustand der Batterie in %
    min_soc_percent: float = 10.0, # Minimaler Ladezustand in %
    max_soc_percent: float = 90.0, # Maximaler Ladezustand in %
    annual_capacity_loss_percent: float = 2.0, # Jährlicher Kapazitätsverlust in %
    battery_tech_params: dict | None = None,
    project_interest_rate_db: float = 0.03,  # Zinssatz für DB-Rechnung
) -> list:
    """
    Findet die wirtschaftlich optimale Speichergröße durch Iteration über verschiedene Kapazitäten.
    """
    results = []
    
    # OPTIMIERUNG: Simulation ohne Batterie nur EINMAL ausführen
    print("🔄 Führe Simulation ohne Batterie aus (einmalig)...")
    no_battery_sim = simulate_one_year(
        consumption_series=consumption_series,
        pv_generation_series=pv_generation_series,
        battery_capacity_kwh=0.0, # Keine Batterie
        battery_efficiency_charge=1.0, # Irrelevant
        battery_efficiency_discharge=1.0, # Irrelevant
        battery_max_charge_kw=0.0, # Keine Ladung
        battery_max_discharge_kw=0.0, # Keine Entladung
        price_grid_per_kwh=price_grid_per_kwh,
        price_feed_in_per_kwh=price_feed_in_per_kwh,
        initial_soc_percent=0.0,
        min_soc_percent=0.0,
        max_soc_percent=0.0,
        annual_capacity_loss_percent=0.0,
        simulation_year=1
    )
    print("✅ Simulation ohne Batterie abgeschlossen!")
    
    # Helper: pro Kapazität passende Lade-/Entladeleistung bestimmen
    def resolve_power_for_capacity(capacity_kwh: float):
        cap_int = int(round(capacity_kwh))
        if battery_tech_params and isinstance(battery_tech_params, dict):
            # exakte Kapazität
            if cap_int in battery_tech_params:
                m = battery_tech_params[cap_int]
                return (
                    float(m.get("max_charge_kw", battery_max_charge_kw)),
                    float(m.get("max_discharge_kw", battery_max_discharge_kw)),
                )
            # nächstliegende Kapazität suchen
            keys = sorted(battery_tech_params.keys())
            if keys:
                nearest = min(keys, key=lambda k: abs(k - cap_int))
                m = battery_tech_params.get(nearest, {})
                return (
                    float(m.get("max_charge_kw", battery_max_charge_kw)),
                    float(m.get("max_discharge_kw", battery_max_discharge_kw)),
                )
        # Fallback auf feste UI/Default-Werte
        return (battery_max_charge_kw, battery_max_discharge_kw)
    
    for capacity in np.arange(min_capacity_kwh, max_capacity_kwh + step_kwh, step_kwh):
        if capacity == 0: # Szenario ohne Speicher
            # Verwende die bereits berechnete Simulation ohne Batterie
            sim_result = no_battery_sim
        else:
            cap_charge_kw, cap_discharge_kw = resolve_power_for_capacity(capacity)
            sim_result = simulate_one_year(
                consumption_series=consumption_series,
                pv_generation_series=pv_generation_series,
                battery_capacity_kwh=capacity,
                battery_efficiency_charge=battery_efficiency_charge,
                battery_efficiency_discharge=battery_efficiency_discharge,
                battery_max_charge_kw=cap_charge_kw,
                battery_max_discharge_kw=cap_discharge_kw,
                price_grid_per_kwh=price_grid_per_kwh,
                price_feed_in_per_kwh=price_feed_in_per_kwh,
                initial_soc_percent=initial_soc_percent,
                min_soc_percent=min_soc_percent,
                max_soc_percent=max_soc_percent,
                annual_capacity_loss_percent=annual_capacity_loss_percent,
                simulation_year=1  # Optimierung basiert auf erstem Jahr
            )
        
        # Finanzielle KPIs berechnen - mit tatsächlichen Simulationsdaten
        # OPTIMIERUNG: Verwende die bereits berechnete Simulation ohne Batterie
        cap_charge_kw, cap_discharge_kw = resolve_power_for_capacity(capacity) if capacity > 0 else (0.0, 0.0)
        financial_kpis = calculate_financial_kpis(
            sim_result['kpis']['annual_energy_cost'], # Jährliche Kosten mit Speicher
            total_consumption=sim_result['kpis']['total_consumption_kwh'],
            total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
            battery_capacity_kwh=capacity,
            battery_cost_curve=battery_cost_curve,
            project_lifetime_years=project_lifetime_years,
            discount_rate=discount_rate,
            price_grid_per_kwh=price_grid_per_kwh, # Für Referenzkosten ohne Speicher
            price_feed_in_per_kwh=price_feed_in_per_kwh, # Für Referenzkosten ohne Speicher
            annual_capacity_loss_percent=annual_capacity_loss_percent,
            # Neue Parameter für tatsächliche Simulation
            consumption_series=consumption_series,
            pv_generation_series=pv_generation_series,
            battery_efficiency_charge=battery_efficiency_charge,
            battery_efficiency_discharge=battery_efficiency_discharge,
            battery_max_charge_kw=cap_charge_kw,
            battery_max_discharge_kw=cap_discharge_kw,
            initial_soc_percent=initial_soc_percent,
            min_soc_percent=min_soc_percent,
            max_soc_percent=max_soc_percent,
            grid_import_with_battery=sim_result['kpis']['total_grid_import_kwh'],
            grid_export_with_battery=sim_result['kpis']['total_grid_export_kwh'],
            # OPTIMIERUNG: Übergebe die bereits berechnete Simulation ohne Batterie
            no_battery_sim_result=no_battery_sim
        )
        
        # Deckungsbeitrags-KPIs berechnen - neue Funktion
        from analysis import calculate_contribution_margin_kpis
        contribution_margin_kpis = calculate_contribution_margin_kpis(
            annual_energy_cost_with_battery=sim_result['kpis']['annual_energy_cost'],
            total_consumption=sim_result['kpis']['total_consumption_kwh'],
            total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
            battery_capacity_kwh=capacity,
            battery_cost_curve=battery_cost_curve,
            price_grid_per_kwh=price_grid_per_kwh,
            price_feed_in_per_kwh=price_feed_in_per_kwh,
            project_lifetime_years=project_lifetime_years,  # Wird aus config.py geladen wenn None
            discount_rate=discount_rate,  # UI-Wert nutzen
            annual_capacity_loss_percent=annual_capacity_loss_percent,  # UI-Wert nutzen
            project_interest_rate_db=project_interest_rate_db,
            grid_import_with_battery=sim_result['kpis']['total_grid_import_kwh'],
            grid_export_with_battery=sim_result['kpis']['total_grid_export_kwh'],
            grid_import_without_battery=None,  # Wird intern berechnet
            grid_export_without_battery=None,  # Wird intern berechnet
            consumption_series=consumption_series,  # Für echte Simulation ohne Batterie
            pv_generation_series=pv_generation_series  # Für echte Simulation ohne Batterie
        )

        # Berechne zusätzliche Kennzahlen für die erweiterte Tabelle
        total_pv_generation = sim_result['kpis']['total_pv_generation_kwh']
        total_consumption = sim_result['kpis']['total_consumption_kwh']
        total_grid_import = sim_result['kpis']['total_grid_import_kwh']
        total_grid_export = sim_result['kpis']['total_grid_export_kwh']
        total_battery_charge = sim_result['kpis']['total_battery_charge_kwh']
        total_battery_discharge = sim_result['kpis']['total_battery_discharge_kwh']
        total_battery_charge_losses = sim_result['kpis']['total_battery_charge_losses_kwh']
        total_battery_discharge_losses = sim_result['kpis']['total_battery_discharge_losses_kwh']
        
        # Selbstgenutzter Strom = Direkter Eigenverbrauch + Batterieentladung
        total_self_consumption = sim_result['kpis']['total_direct_self_consumption_kwh'] + total_battery_discharge
        
        # Gesamte Speicherverluste
        total_battery_losses = total_battery_charge_losses + total_battery_discharge_losses
        
        results.append({
            'battery_capacity_kwh': capacity,
            'autarky_rate': sim_result['kpis']['autarky_rate'],
            'self_consumption_rate': sim_result['kpis']['self_consumption_rate'],
            'effective_annual_energy_cost': sim_result['kpis'].get('effective_annual_energy_cost', sim_result['kpis']['annual_energy_cost']),
            'grid_import_cost': sim_result['kpis'].get('grid_import_cost', 0),
            'grid_export_revenue': sim_result['kpis'].get('grid_export_revenue', 0),
            # Neue Spalten für erweiterte Analyse
            'self_consumption_kwh': total_self_consumption,
            'grid_export_kwh': total_grid_export,
            'pv_generation_kwh': total_pv_generation,
            'total_consumption_kwh': total_consumption,
            'grid_import_kwh': total_grid_import,
            'battery_losses_kwh': total_battery_losses,
            **financial_kpis, # Füge finanzielle KPIs hinzu (Cash Flow-basierte Amortisationszeit)
            **contribution_margin_kpis, # Füge Deckungsbeitrags-KPIs hinzu (DB3 für Rentabilität)
            # KORREKTUR: Verwende Cash Flow-basierte Amortisationszeit (wirtschaftlich üblich)
            'payback_period_years': financial_kpis['payback_period_years']
        })
    return results

def run_variable_tariff_scenario(
    consumption_series: pd.Series,
    pv_generation_series: pd.Series,
    battery_capacity_kwh: float,
    battery_efficiency_charge: float,
    battery_efficiency_discharge: float,
    battery_max_charge_kw: float,
    battery_max_discharge_kw: float,
    price_grid_series: pd.Series, # Stündliche Bezugspreise
    price_feed_in_series: pd.Series, # Stündliche Einspeisepreise
    battery_cost_curve: dict,  # Dictionary mit Kostenkurven-Daten
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05,
    initial_soc_percent: float = 50.0, # Anfangsladezustand der Batterie in %
    min_soc_percent: float = 10.0, # Minimaler Ladezustand in %
    max_soc_percent: float = 90.0, # Maximaler Ladezustand in %
    annual_capacity_loss_percent: float = 2.0, # Jährlicher Kapazitätsverlust in %
    battery_tech_params: dict | None = None
) -> dict:
    """
    Führt eine Simulation mit variablen Stromtarifen durch.
    """
    # Pro Kapazität Lade-/Entladeleistung bestimmen (Excel-basiert, falls verfügbar)
    def resolve_power_for_capacity(capacity_kwh: float):
        cap_int = int(round(capacity_kwh))
        if battery_tech_params and isinstance(battery_tech_params, dict):
            if cap_int in battery_tech_params:
                m = battery_tech_params[cap_int]
                return (
                    float(m.get("max_charge_kw", battery_max_charge_kw)),
                    float(m.get("max_discharge_kw", battery_max_discharge_kw)),
                )
            keys = sorted(battery_tech_params.keys())
            if keys:
                nearest = min(keys, key=lambda k: abs(k - cap_int))
                m = battery_tech_params.get(nearest, {})
                return (
                    float(m.get("max_charge_kw", battery_max_charge_kw)),
                    float(m.get("max_discharge_kw", battery_max_discharge_kw)),
                )
        return (battery_max_charge_kw, battery_max_discharge_kw)

    cap_charge_kw, cap_discharge_kw = resolve_power_for_capacity(battery_capacity_kwh) if battery_capacity_kwh > 0 else (0.0, 0.0)
    sim_result = simulate_one_year(
        consumption_series=consumption_series,
        pv_generation_series=pv_generation_series,
        battery_capacity_kwh=battery_capacity_kwh,
        battery_efficiency_charge=battery_efficiency_charge,
        battery_efficiency_discharge=battery_efficiency_discharge,
        battery_max_charge_kw=cap_charge_kw,
        battery_max_discharge_kw=cap_discharge_kw,
        price_grid_per_kwh=price_grid_series, # Übergabe der Zeitreihe
        price_feed_in_per_kwh=price_feed_in_series, # Übergabe der Zeitreihe
        initial_soc_percent=initial_soc_percent,
        min_soc_percent=min_soc_percent,
        max_soc_percent=max_soc_percent,
        annual_capacity_loss_percent=annual_capacity_loss_percent,
        simulation_year=1  # Variable Tarife basieren auf erstem Jahr
    )

    financial_kpis = calculate_financial_kpis(
        sim_result['kpis']['annual_energy_cost'],
        total_consumption=sim_result['kpis']['total_consumption_kwh'],
        total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
        battery_capacity_kwh=battery_capacity_kwh,
        battery_cost_curve=battery_cost_curve,
        project_lifetime_years=project_lifetime_years,
        discount_rate=discount_rate,
        price_grid_per_kwh=price_grid_series, # Für Referenzkosten ohne Speicher
        price_feed_in_per_kwh=price_feed_in_series, # Für Referenzkosten ohne Speicher
        annual_capacity_loss_percent=annual_capacity_loss_percent,
        consumption_series=consumption_series,
        pv_generation_series=pv_generation_series,
        battery_efficiency_charge=battery_efficiency_charge,
        battery_efficiency_discharge=battery_efficiency_discharge,
        battery_max_charge_kw=cap_charge_kw,
        battery_max_discharge_kw=cap_discharge_kw,
        initial_soc_percent=initial_soc_percent,
        min_soc_percent=min_soc_percent,
        max_soc_percent=max_soc_percent,
        grid_import_with_battery=sim_result['kpis']['total_grid_import_kwh'],
        grid_export_with_battery=sim_result['kpis']['total_grid_export_kwh']
    )

    return {
        'scenario_name': 'Variable Tarife',
        'time_series_data': sim_result['time_series_data'],
        'kpis': {**sim_result['kpis'], **financial_kpis}
    }

def compare_household_scenarios(
    household_profiles: dict, # Dictionary mit {Name: {'consumption': Series, 'pv_gen': Series, ...}}
    battery_capacity_kwh: float,
    battery_efficiency_charge: float,
    battery_efficiency_discharge: float,
    battery_max_charge_kw: float,
    battery_max_discharge_kw: float,
    price_grid_per_kwh,
    price_feed_in_per_kwh,
    battery_cost_curve: dict,  # Dictionary mit Kostenkurven-Daten
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05,
    initial_soc_percent: float = 50.0, # Anfangsladezustand der Batterie in %
    min_soc_percent: float = 10.0, # Minimaler Ladezustand in %
    max_soc_percent: float = 90.0 # Maximaler Ladezustand in %
) -> dict:
    """
    Vergleicht die Simulation für verschiedene Haushaltstypen.
    """
    all_scenario_results = {}
    for name, profile_data in household_profiles.items():
        sim_result = simulate_one_year(
            consumption_series=profile_data['consumption'],
            pv_generation_series=profile_data['pv_gen'],
            battery_capacity_kwh=battery_capacity_kwh,
            battery_efficiency_charge=battery_efficiency_charge,
            battery_efficiency_discharge=battery_efficiency_discharge,
            battery_max_charge_kw=battery_max_charge_kw,
            battery_max_discharge_kw=battery_max_discharge_kw,
            price_grid_per_kwh=price_grid_per_kwh,
            price_feed_in_per_kwh=price_feed_in_per_kwh,
            initial_soc_percent=initial_soc_percent,
            min_soc_percent=min_soc_percent,
            max_soc_percent=max_soc_percent,
            annual_capacity_loss_percent=DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT,  # Aus config.py (vereinheitlicht)
            simulation_year=1
        )

        financial_kpis = calculate_financial_kpis(
            sim_result['kpis']['annual_energy_cost'],
            total_consumption=sim_result['kpis']['total_consumption_kwh'],
            total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
            battery_capacity_kwh=battery_capacity_kwh,
            battery_cost_curve=battery_cost_curve,
            project_lifetime_years=project_lifetime_years,
            discount_rate=discount_rate,
            price_grid_per_kwh=price_grid_per_kwh, # Für Referenzkosten ohne Speicher
            price_feed_in_per_kwh=price_feed_in_per_kwh, # Für Referenzkosten ohne Speicher
            annual_capacity_loss_percent=DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT,  # Aus config.py (vereinheitlicht)
            consumption_series=profile_data['consumption'],  # Für tatsächliche Simulation ohne Batterie
            pv_generation_series=profile_data['pv_gen'],  # Für tatsächliche Simulation ohne Batterie
            battery_efficiency_charge=battery_efficiency_charge,
            battery_efficiency_discharge=battery_efficiency_discharge,
            battery_max_charge_kw=battery_max_charge_kw,
            battery_max_discharge_kw=battery_max_discharge_kw,
            initial_soc_percent=initial_soc_percent,
            min_soc_percent=min_soc_percent,
            max_soc_percent=max_soc_percent,
            grid_import_with_battery=sim_result['kpis']['total_grid_import_kwh'],
            grid_export_with_battery=sim_result['kpis']['total_grid_export_kwh']
        )

        all_scenario_results[name] = {
            'time_series_data': sim_result['time_series_data'],
            'kpis': {**sim_result['kpis'], **financial_kpis}
        }
    return all_scenario_results 