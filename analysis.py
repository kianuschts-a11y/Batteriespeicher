import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def apply_modern_plotly_theme(fig):
    """
    Wendet ein modernes, professionelles Design auf Plotly-Graphen an.
    """
    # Moderne Farbpalette (Orange-Theme)
    colors = {
        'primary': '#FF8C42',      # Hauptorange
        'secondary': '#FF6B35',    # Sekundärorange  
        'accent': '#FFB366',       # Akzentorange
        'success': '#28A745',      # Grün für positive Werte
        'danger': '#DC3545',       # Rot für negative Werte
        'warning': '#FFC107',      # Gelb für Warnungen
        'info': '#17A2B8',         # Blau für Informationen
        'light': '#F8F9FA',        # Hellgrau
        'dark': '#212529',         # Dunkelgrau
        'white': '#FFFFFF',        # Weiß
        'gray': '#6C757D'         # Grau
    }
    
    # Moderne Layout-Einstellungen
    fig.update_layout(
        # Dunkles Theme für bessere Lesbarkeit
        template='plotly_dark',
        
        # Moderne Schriftarten
        font=dict(
            family="'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif",
            size=12,
            color=colors['white']
        ),
        
        # Titel-Styling
        title=dict(
            font=dict(
                size=18,
                color=colors['primary'],
                family="'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif"
            ),
            x=0.5,  # Zentriert
            xanchor='center'
        ),
        
        # Achsen-Styling
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(255, 255, 255, 0.1)',
            gridwidth=1,
            zeroline=False,
            showline=True,
            linecolor=colors['gray'],
            linewidth=1,
            tickfont=dict(color=colors['white'], size=11),
            title_font=dict(color=colors['white'], size=13)
        ),
        
        yaxis=dict(
            showgrid=True,
            gridcolor='rgba(255, 255, 255, 0.1)',
            gridwidth=1,
            zeroline=False,
            showline=True,
            linecolor=colors['gray'],
            linewidth=1,
            tickfont=dict(color=colors['white'], size=11),
            title_font=dict(color=colors['white'], size=13)
        ),
        
        # Hintergrund
        plot_bgcolor='rgba(0, 0, 0, 0)',  # Transparent
        paper_bgcolor='rgba(0, 0, 0, 0)',  # Transparent
        
        # Legende
        legend=dict(
            bgcolor='rgba(0, 0, 0, 0.8)',
            bordercolor=colors['gray'],
            borderwidth=1,
            font=dict(color=colors['white'], size=11),
            x=1.02,  # Rechts außen
            y=1,
            xanchor='left',
            yanchor='top'
        ),
        
        # Hover-Label
        hoverlabel=dict(
            bgcolor='rgba(0, 0, 0, 0.9)',
            bordercolor=colors['primary'],
            font_size=12,
            font_color=colors['white']
        ),
        
        # Moderne Animationen
        transition=dict(
            duration=0,
            easing='cubic-in-out'
        ),
        
        # Responsive Design
        autosize=True,
        margin=dict(l=60, r=60, t=80, b=60),
        
        # Moderne Toolbar
        modebar=dict(
            bgcolor='rgba(0, 0, 0, 0.8)',
            color=colors['white'],
            activecolor=colors['primary']
        )
    )
    
    # Moderne Linien-Styling für alle Traces
    for trace in fig.data:
        if trace.type == 'scatter':
            # Glatte Linien mit Schatten-Effekt
            trace.update(
                line=dict(
                    width=3,
                    shape='spline',  # Glatte Kurven
                    smoothing=0.3
                ),
                marker=dict(
                    size=6,
                    line=dict(width=2, color=colors['white'])
                ),
                # Schatten-Effekt durch zusätzliche Linie
                opacity=0.9
            )
        elif trace.type == 'bar':
            # Moderne Balken mit Gradient
            trace.update(
                marker=dict(
                    line=dict(width=1, color=colors['white']),
                    opacity=0.8
                )
            )
    
    return fig

def create_modern_color_palette(n_colors):
    """
    Erstellt eine moderne Farbpalette basierend auf dem Orange-Theme.
    """
    base_colors = [
        '#FF8C42',  # Hauptorange
        '#28A745',  # Grün
        '#17A2B8',  # Blau
        '#FFC107',  # Gelb
        '#DC3545',  # Rot
        '#6F42C1',  # Lila
        '#20C997',  # Türkis
        '#FD7E14',  # Orange
        '#E83E8C',  # Pink
        '#6C757D'   # Grau
    ]
    
    # Wiederhole Farben falls mehr benötigt werden
    return [base_colors[i % len(base_colors)] for i in range(n_colors)]

def add_modern_annotations(fig, annotations_data):
    """
    Fügt moderne Annotationen zu Plotly-Graphen hinzu.
    """
    colors = {
        'primary': '#FF8C42',
        'success': '#28A745',
        'danger': '#DC3545',
        'warning': '#FFC107',
        'info': '#17A2B8'
    }
    
    for annotation in annotations_data:
        fig.add_annotation(
            x=annotation['x'],
            y=annotation['y'],
            text=annotation['text'],
            showarrow=annotation.get('showarrow', True),
            arrowhead=annotation.get('arrowhead', 2),
            arrowcolor=colors.get(annotation.get('color', 'primary'), colors['primary']),
            arrowwidth=annotation.get('arrowwidth', 2),
            bgcolor=f"rgba({int(colors[annotation.get('color', 'primary')][1:3], 16)}, {int(colors[annotation.get('color', 'primary')][3:5], 16)}, {int(colors[annotation.get('color', 'primary')][5:7], 16)}, 0.8)",
            bordercolor=colors.get(annotation.get('color', 'primary'), colors['primary']),
            borderwidth=annotation.get('borderwidth', 1),
            font=dict(
                color='white',
                size=annotation.get('fontsize', 12),
                family="'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif"
            ),
            xanchor=annotation.get('xanchor', 'center'),
            yanchor=annotation.get('yanchor', 'middle')
        )
    
    return fig

def add_modern_shapes(fig, shapes_data):
    """
    Fügt moderne Formen (Linien, Rechtecke) zu Plotly-Graphen hinzu.
    """
    colors = {
        'primary': '#FF8C42',
        'success': '#28A745',
        'danger': '#DC3545',
        'warning': '#FFC107',
        'info': '#17A2B8'
    }
    
    for shape in shapes_data:
        fig.add_shape(
            type=shape['type'],
            x0=shape['x0'],
            y0=shape['y0'],
            x1=shape['x1'],
            y1=shape['y1'],
            line=dict(
                color=colors.get(shape.get('color', 'primary'), colors['primary']),
                width=shape.get('width', 2),
                dash=shape.get('dash', 'solid')
            ),
            fillcolor=shape.get('fillcolor', 'rgba(0,0,0,0)'),
            opacity=shape.get('opacity', 1.0)
        )
    
    return fig

def calculate_irr(cash_flows, max_iterations=100, tolerance=1e-6):
    """
    Berechnet den Internal Rate of Return (IRR) für eine Reihe von Cash Flows.
    
    Args:
        cash_flows (list): Liste der Cash Flows (erster Wert ist die Initialinvestition, negativ)
        max_iterations (int): Maximale Anzahl Iterationen für die Newton-Raphson Methode
        tolerance (float): Toleranz für die Konvergenz
        
    Returns:
        float: IRR als Dezimalzahl (z.B. 0.05 für 5%)
    """
    if len(cash_flows) < 2:
        return np.nan
    
    # Initial guess: 10%
    rate = 0.10
    
    for iteration in range(max_iterations):
        npv = 0
        npv_derivative = 0
        
        for i, cash_flow in enumerate(cash_flows):
            if i == 0:  # Initialinvestition
                npv += cash_flow
                npv_derivative += 0  # Ableitung der Konstante ist 0
            else:
                discount_factor = (1 + rate) ** i
                npv += cash_flow / discount_factor
                npv_derivative += -i * cash_flow / (discount_factor * (1 + rate))
        
        # Newton-Raphson Update
        if abs(npv_derivative) < tolerance:
            break
            
        rate_new = rate - npv / npv_derivative
        
        # Prüfe auf Konvergenz
        if abs(rate_new - rate) < tolerance:
            rate = rate_new
            break
            
        rate = rate_new
        
        # Prüfe auf unrealistische Werte
        if rate < -0.99 or rate > 10:  # IRR zwischen -99% und 1000%
            return np.nan
    
    return rate

def calculate_financial_kpis(
    annual_energy_cost_with_battery: float, # Jährliche Energiekosten mit Batteriespeicher
    total_consumption: float, # Gesamter jährlicher Verbrauch
    total_pv_generation: float, # Gesamte jährliche PV-Erzeugung
    battery_capacity_kwh: float,
    battery_cost_curve: dict,  # Dictionary mit Kostenkurven-Daten
    project_lifetime_years: int,
    discount_rate: float,
    price_grid_per_kwh, # Kann float oder pd.Series sein
    price_feed_in_per_kwh, # Kann float oder pd.Series sein
    annual_capacity_loss_percent: float, # Jährlicher Kapazitätsverlust in Prozent
    consumption_series: pd.Series = None, # Für tatsächliche Simulation
    pv_generation_series: pd.Series = None, # Für tatsächliche Simulation
    battery_efficiency_charge: float = 0.95, # Für tatsächliche Simulation
    battery_efficiency_discharge: float = 0.95, # Für tatsächliche Simulation
    battery_max_charge_kw: float = 10.0, # Für tatsächliche Simulation
    battery_max_discharge_kw: float = 10.0, # Für tatsächliche Simulation
    initial_soc_percent: float = 50.0, # Für tatsächliche Simulation
    min_soc_percent: float = 10.0, # Für tatsächliche Simulation
    max_soc_percent: float = 90.0, # Für tatsächliche Simulation
    grid_import_with_battery: float = None, # Netzbezug mit Batterie
    grid_export_with_battery: float = None, # Netzeinspeisung mit Batterie
    no_battery_sim_result: dict = None  # OPTIMIERUNG: Bereits berechnete Simulation ohne Batterie
) -> dict:
    """
    Berechnet finanzielle KPIs wie Amortisationszeit und Net Present Value (NPV).
    """
    # Spezialfall: 0 kWh Batteriekapazität (Referenzfall ohne Batterie)
    if battery_capacity_kwh == 0:
        return {
            'investment_cost': 0,  # Keine Investition ohne Batterie
            'annual_savings': 0,   # Keine Einsparungen ohne Batterie
            'payback_period_years': np.nan,  # Keine Amortisationszeit
            'npv': 0,  # Kein NPV
            'irr_percentage': np.nan  # Kein IRR
        }
    
    # Investitionskosten (nur für Kapazitäten > 0) - aus Kostenkurve
    from data_import import get_battery_cost
    investment_cost = get_battery_cost(battery_capacity_kwh, battery_cost_curve)

    # Berechne Ersparnis nach der korrekten Formel:
    # Ersparnis = (verringerter Netzbezug × Strompreis) - (verringerte Einspeisung × Einspeisevergütung)
    
    if consumption_series is not None and pv_generation_series is not None and grid_import_with_battery is not None and grid_export_with_battery is not None:
        # OPTIMIERUNG: Verwende bereits berechnete Simulation ohne Batterie falls verfügbar
        if no_battery_sim_result is not None:
            print(f"🚀 OPTIMIERUNG: Verwende bereits berechnete Simulation ohne Batterie!")
            grid_import_no_battery = no_battery_sim_result["kpis"]["total_grid_import_kwh"]
            grid_export_no_battery = no_battery_sim_result["kpis"]["total_grid_export_kwh"]
        else:
            # Fallback: Führe Simulation ohne Batterie durch (falls nicht bereits berechnet)
            print(f"⚠️ FALLBACK: Führe Simulation ohne Batterie durch...")
            total_consumption = consumption_series.sum()
            total_pv_generation = pv_generation_series.sum()
            
            try:
                # Führe echte Simulation ohne Batterie durch
                from model import simulate_one_year
                no_battery_sim = simulate_one_year(
                    consumption_series=consumption_series,
                    pv_generation_series=pv_generation_series,
                    battery_capacity_kwh=0.0,  # KEINE Batterie
                    battery_efficiency_charge=0.95,  # Standard-Wirkungsgrad
                    battery_efficiency_discharge=0.95,  # Standard-Wirkungsgrad
                    battery_max_charge_kw=0.0,  # KEINE Ladeleistung
                    battery_max_discharge_kw=0.0,  # KEINE Entladeleistung
                    price_grid_per_kwh=price_grid_per_kwh,
                    price_feed_in_per_kwh=price_feed_in_per_kwh,
                    max_soc_percent=0.0,
                    annual_capacity_loss_percent=0.0,
                    simulation_year=1
                )
                grid_import_no_battery = no_battery_sim["kpis"]["total_grid_import_kwh"]
                grid_export_no_battery = no_battery_sim["kpis"]["total_grid_export_kwh"]
                
                # VALIDIERUNG: Prüfe ob die Simulation ohne Batterie plausibel ist
                if grid_import_no_battery < 0 or grid_export_no_battery < 0:
                    print(f"⚠️ WARNUNG: Simulation ohne Batterie liefert negative Werte!")
                    print(f"  grid_import_no_battery: {grid_import_no_battery}")
                    print(f"  grid_export_no_battery: {grid_export_no_battery}")
                    raise ValueError("Negative Werte in Simulation ohne Batterie")
                
            except Exception as e:
                print(f"❌ FEHLER in Simulation ohne Batterie: {e}")
                print(f"🔄 Fallback auf vereinfachte Berechnung...")
                # Fallback auf vereinfachte Berechnung
                if isinstance(price_grid_per_kwh, pd.Series):
                    annual_cost_no_battery = (total_consumption * price_grid_per_kwh.mean()) - (total_pv_generation * price_feed_in_per_kwh.mean())
                else:
                    annual_cost_no_battery = (total_consumption * price_grid_per_kwh) - (total_pv_generation * price_feed_in_per_kwh)
                
                annual_savings = annual_cost_no_battery - annual_energy_cost_with_battery
                savings_from_reduced_import = None
                loss_from_reduced_export = None
        
        # Berechne verringerten Netzbezug und verringerte Einspeisung
        reduced_grid_import = grid_import_no_battery - grid_import_with_battery
        reduced_grid_export = grid_export_no_battery - grid_export_with_battery  # Ohne Batterie - mit Batterie (mehr Einspeisung ohne Batterie)
        
        # KORREKTE Ersparnis-Berechnung nach der Formel
        # Ersparnis = Weniger Netzbezug (Kosteneinsparung) - Weniger Einspeisung (Kostenverlust)
        # Da reduced_grid_export positiv ist (mehr Einspeisung ohne Batterie), ist das ein Kostenverlust
        if isinstance(price_grid_per_kwh, pd.Series):
            annual_savings = (reduced_grid_import * price_grid_per_kwh.mean()) - (reduced_grid_export * price_feed_in_per_kwh.mean())
        else:
            annual_savings = (reduced_grid_import * price_grid_per_kwh) - (reduced_grid_export * price_feed_in_per_kwh)
        
        # Berechne die einzelnen Komponenten der Ersparnis für transparente Darstellung
        if isinstance(price_grid_per_kwh, pd.Series):
            savings_from_reduced_import = reduced_grid_import * price_grid_per_kwh.mean()
            loss_from_reduced_export = reduced_grid_export * price_feed_in_per_kwh.mean()
        else:
            savings_from_reduced_import = reduced_grid_import * price_grid_per_kwh
            loss_from_reduced_export = reduced_grid_export * price_feed_in_per_kwh
        
        # Debug-Ausgabe
        print(f"DEBUG calculate_financial_kpis:")
        print(f"  total_consumption: {total_consumption:.2f} kWh")
        print(f"  total_pv_generation: {total_pv_generation:.2f} kWh")
        print(f"  grid_import_no_battery: {grid_import_no_battery:.2f} kWh")
        print(f"  grid_import_with_battery: {grid_import_with_battery:.2f} kWh")
        print(f"  reduced_grid_import: {reduced_grid_import:.2f} kWh")
        print(f"  grid_export_no_battery: {grid_export_no_battery:.2f} kWh")
        print(f"  grid_export_with_battery: {grid_export_with_battery:.2f} kWh")
        print(f"  reduced_grid_export: {reduced_grid_export:.2f} kWh")
        print(f"  savings_from_reduced_import: {savings_from_reduced_import:.2f} EUR")
        print(f"  loss_from_reduced_export: {loss_from_reduced_export:.2f} EUR")
        print(f"  annual_savings: {annual_savings:.2f} EUR")
        print(f"  price_grid_per_kwh: {price_grid_per_kwh}")
        print(f"  price_feed_in_per_kwh: {price_feed_in_per_kwh}")
    else:
        # Fallback: Verwende die ursprüngliche Berechnung wenn keine Zeitreihen verfügbar
        # Setze die Komponenten auf None für den Fallback
        savings_from_reduced_import = None
        loss_from_reduced_export = None
        
        if isinstance(price_grid_per_kwh, pd.Series):
            annual_cost_no_battery = (total_consumption * price_grid_per_kwh.mean()) - (total_pv_generation * price_feed_in_per_kwh.mean())
        else:
            annual_cost_no_battery = (total_consumption * price_grid_per_kwh) - (total_pv_generation * price_feed_in_per_kwh)
        
        annual_savings = annual_cost_no_battery - annual_energy_cost_with_battery

    # Amortisationszeit (Payback Period) - KORRIGIERT: Degradation der Ersparnisse statt Kapazität
    payback_period = np.nan # Initialisiere mit NaN
    if annual_savings > 0:
        # Berechne kumulierte Ersparnisse über die Zeit mit Degradation
        cumulative_savings = 0
        for year in range(1, project_lifetime_years + 1):
            # Degradation der Ersparnisse über die Jahre (nicht der Kapazität!)
            # Wenn die Batteriekapazität um 1% abnimmt, nehmen auch die Ersparnisse um ~1% ab
            degradation_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (year - 1)
            year_savings = annual_savings * degradation_factor
            
            cumulative_savings += year_savings
            
            # Prüfe Break-even
            if cumulative_savings >= investment_cost:
                # Berechne genaue Payback Period mit Bruchteilen
                if year == 1:
                    # Im ersten Jahr: Einfache Division
                    payback_period = investment_cost / year_savings
                else:
                    # In späteren Jahren: Interpolation zwischen den Jahren
                    previous_cumulative = cumulative_savings - year_savings
                    remaining_investment = investment_cost - previous_cumulative
                    fraction_of_year = remaining_investment / year_savings
                    payback_period = (year - 1) + fraction_of_year
                break
        else:
            payback_period = 999  # Kein Break-even gefunden
    else:
        # Wenn keine Ersparnis, setze Amortisationszeit auf einen hohen Wert
        payback_period = 999  # Unrealistisch hoher Wert für nicht wirtschaftliche Investitionen

    # Net Present Value (NPV) mit Degradation - KORRIGIERT: Degradation der Ersparnisse
    cash_flows = [-investment_cost] # Initialinvestition
    
    # Berechne jährliche Ersparnisse mit Degradation
    for year in range(1, project_lifetime_years + 1):
        # Degradation der Ersparnisse über die Jahre (nicht der Kapazität!)
        degradation_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (year - 1)
        year_savings = annual_savings * degradation_factor
        
        cash_flows.append(year_savings)
    
    # Manuelle NPV-Berechnung (da np.npv in neueren NumPy-Versionen entfernt wurde)
    npv = 0
    for i, cash_flow in enumerate(cash_flows):
        npv += cash_flow / ((1 + discount_rate) ** i)

    # IRR-Berechnung
    irr = calculate_irr(cash_flows)
    irr_percentage = irr * 100 if not np.isnan(irr) else np.nan

    # Debug-Ausgabe für Cash Flow-basierte Amortisationszeit
    print(f"🔍 DIAGNOSE Cash Flow-Amortisation für {battery_capacity_kwh} kWh Batterie:")
    print(f"  💰 Investition: {investment_cost:.0f} EUR")
    print(f"  💵 Jährliche Ersparnis: {annual_savings:.0f} EUR")
    print(f"  📊 Einfache Amortisation: {investment_cost/annual_savings:.1f} Jahre" if annual_savings > 0 else "  ❌ Keine Ersparnis!")
    if 'reduced_grid_import' in locals():
        print(f"  🔋 Reduzierter Netzbezug: {reduced_grid_import:.0f} kWh")
        print(f"  ⚡ Reduzierte Einspeisung: {reduced_grid_export:.0f} kWh")
        print(f"  💲 Strompreis: {price_grid_per_kwh.mean() if isinstance(price_grid_per_kwh, pd.Series) else price_grid_per_kwh:.3f} EUR/kWh")
        print(f"  💲 Einspeisevergütung: {price_feed_in_per_kwh.mean() if isinstance(price_feed_in_per_kwh, pd.Series) else price_feed_in_per_kwh:.3f} EUR/kWh")
        print(f"  📈 Ersparnis durch weniger Netzbezug: {savings_from_reduced_import:.0f} EUR")
        print(f"  📉 Verlust durch weniger Einspeisung: {loss_from_reduced_export:.0f} EUR")
    print(f"  🎯 Netto-Ersparnis: {annual_savings:.0f} EUR")
    print(f"  ⏰ Amortisationszeit: {payback_period:.1f} Jahre")
    print(f"  {'✅' if payback_period < 50 else '❌'} {'OK' if payback_period < 50 else 'PROBLEM!'}")
    print(f"  ---")

    return {
        'investment_cost': investment_cost,
        'annual_savings': annual_savings,
        'savings_from_reduced_import': savings_from_reduced_import,  # Ersparnis durch reduzierten Netzbezug
        'loss_from_reduced_export': loss_from_reduced_export,  # Verlust durch reduzierte Einspeisung
        'payback_period_years': payback_period,
        'npv': npv,
        'irr_percentage': irr_percentage
    }

def calculate_contribution_margin_kpis(
    annual_energy_cost_with_battery: float,
    total_consumption: float,
    total_pv_generation: float,
    battery_capacity_kwh: float,
    battery_cost_curve: dict,
    price_grid_per_kwh,  # Kann float oder pd.Series sein (konsistent mit calculate_financial_kpis)
    price_feed_in_per_kwh,  # Kann float oder pd.Series sein (konsistent mit calculate_financial_kpis)
    project_lifetime_years: int = None,  # Wird aus config.py geladen
    discount_rate: float = None,  # Wird aus config.py geladen
    annual_capacity_loss_percent: float = None,  # Wird aus config.py geladen
    project_interest_rate_db: float = 0.03,  # 3% Zinsen für DB
    grid_import_with_battery: float = None,
    grid_export_with_battery: float = None,
    grid_import_without_battery: float = None,
    grid_export_without_battery: float = None,
    consumption_series: pd.Series = None,  # Für echte Simulation ohne Batterie
    pv_generation_series: pd.Series = None  # Für echte Simulation ohne Batterie
) -> dict:
    """
    Berechnet Deckungsbeitrags-KPIs basierend auf der bestehenden Excel-Struktur.
    
    Args:
        annual_energy_cost_with_battery: Jährliche Energiekosten mit Batteriespeicher
        total_consumption: Gesamter jährlicher Verbrauch
        total_pv_generation: Gesamte jährliche PV-Erzeugung
        battery_capacity_kwh: Batteriekapazität in kWh
        battery_cost_curve: Dictionary mit Kostenkurven-Daten
        price_grid_per_kwh: Preis für Netzbezug in €/kWh (float oder pd.Series)
        price_feed_in_per_kwh: Preis für Netzeinspeisung in €/kWh (float oder pd.Series)
        project_lifetime_years: Projektlaufzeit in Jahren (optional, Standard aus config.py)
        discount_rate: Abzinsungssatz (optional, Standard aus config.py)
        annual_capacity_loss_percent: Jährliche Degradation in % (optional, Standard aus config.py)
        project_interest_rate_db: Zinssatz für Deckungsbeitragsrechnung in %
        grid_import_with_battery: Netzbezug mit Batterie (optional)
        grid_export_with_battery: Netzeinspeisung mit Batterie (optional)
        grid_import_without_battery: Netzbezug ohne Batterie (optional)
        grid_export_without_battery: Netzeinspeisung ohne Batterie (optional)
        consumption_series: Verbrauchszeitreihe für Simulation ohne Batterie
        pv_generation_series: PV-Erzeugungszeitreihe für Simulation ohne Batterie
    
    Returns:
        dict: Deckungsbeitrags-KPIs
    """
    
    # Lade Standardwerte aus config.py (konsistent mit calculate_financial_kpis)
    from config import DEFAULT_PROJECT_LIFETIME_YEARS, DEFAULT_DISCOUNT_RATE, DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT
    
    if project_lifetime_years is None:
        project_lifetime_years = DEFAULT_PROJECT_LIFETIME_YEARS
    if discount_rate is None:
        discount_rate = DEFAULT_DISCOUNT_RATE
    if annual_capacity_loss_percent is None:
        annual_capacity_loss_percent = DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT
    # Spezialfall: 0 kWh Batteriekapazität (Referenzfall ohne Batterie)
    if battery_capacity_kwh == 0:
        return {
            'investment_cost': 0,
            'annual_revenue': 0,
            'annual_variable_costs': 0,
            'contribution_margin_1': 0,
            'annual_depreciation': 0,
            'contribution_margin_2': 0,
            'annual_interest': 0,
            'contribution_margin_3': 0,
            'total_db3_nominal': 0,
            'total_db3_present_value': 0,
            'payback_period_years': np.nan,
            'roi_percentage': np.nan
        }
    
    # Investitionskosten aus Kostenkurve
    from data_import import get_battery_cost
    investment_cost = get_battery_cost(battery_capacity_kwh, battery_cost_curve)
    
    # Berechne Netzbezug und -einspeisung ohne Batterie (Referenzszenario)
    if grid_import_without_battery is None or grid_export_without_battery is None:
        # KORREKT: Echte Simulation ohne Batterie (wie in calculate_financial_kpis)
        # Verwende die gleiche Logik wie in der finanziellen Analyse
        from model import simulate_one_year
        
        # ERFORDERT: Zeitreihen müssen verfügbar sein für korrekte Berechnung
        if consumption_series is None or pv_generation_series is None:
            raise ValueError("consumption_series und pv_generation_series sind erforderlich für korrekte DB-Berechnung. "
                           "Vereinfachte Annahmen führen zu falschen Ergebnissen.")
        
        # ECHTE Simulation ohne Batterie
        no_battery_sim = simulate_one_year(
            consumption_series=consumption_series,
            pv_generation_series=pv_generation_series,
            battery_capacity_kwh=0.0,  # KEINE Batterie
            battery_efficiency_charge=0.95,  # Standard-Wirkungsgrad
            battery_efficiency_discharge=0.95,  # Standard-Wirkungsgrad
            battery_max_charge_kw=0.0,  # KEINE Ladeleistung
            battery_max_discharge_kw=0.0,  # KEINE Entladeleistung
            price_grid_per_kwh=price_grid_per_kwh,
            price_feed_in_per_kwh=price_feed_in_per_kwh,
            max_soc_percent=0.0,
            annual_capacity_loss_percent=0.0,
            simulation_year=1
        )
        grid_import_without_battery = no_battery_sim["kpis"]["total_grid_import_kwh"]
        grid_export_without_battery = no_battery_sim["kpis"]["total_grid_export_kwh"]
    
    # Berechne die Differenzen (Ersparnisse/Verluste durch Batterie)
    reduced_grid_import = grid_import_without_battery - (grid_import_with_battery or 0)
    reduced_grid_export = grid_export_without_battery - (grid_export_with_battery or 0)
    
    # Debug-Ausgabe für DB-Berechnung
    print(f"DEBUG calculate_contribution_margin_kpis:")
    print(f"  battery_capacity_kwh: {battery_capacity_kwh}")
    print(f"  grid_import_without_battery: {grid_import_without_battery:.2f} kWh")
    print(f"  grid_import_with_battery: {grid_import_with_battery or 0:.2f} kWh")
    print(f"  reduced_grid_import: {reduced_grid_import:.2f} kWh")
    print(f"  grid_export_without_battery: {grid_export_without_battery:.2f} kWh")
    print(f"  grid_export_with_battery: {grid_export_with_battery or 0:.2f} kWh")
    print(f"  reduced_grid_export: {reduced_grid_export:.2f} kWh")
    
    # Umsatz = Ersparnis durch reduzierten Netzbezug
    # Konsistente Behandlung von float und pd.Series (wie in calculate_financial_kpis)
    if isinstance(price_grid_per_kwh, pd.Series):
        annual_revenue = reduced_grid_import * price_grid_per_kwh.mean()
    else:
        annual_revenue = reduced_grid_import * price_grid_per_kwh
    
    # Variable Kosten = Verlust durch reduzierte Einspeisung
    if isinstance(price_feed_in_per_kwh, pd.Series):
        annual_variable_costs = reduced_grid_export * price_feed_in_per_kwh.mean()
    else:
        annual_variable_costs = reduced_grid_export * price_feed_in_per_kwh
    
    # DB I = Umsatz - Variable Kosten
    contribution_margin_1 = annual_revenue - annual_variable_costs
    
    # Abschreibung = Investition / Projektlaufzeit (konstant)
    annual_depreciation = investment_cost / project_lifetime_years
    
    # DB II = DB I - Abschreibung
    contribution_margin_2 = contribution_margin_1 - annual_depreciation
    
    # Zinsen = Durchschnittliches gebundenes Kapital × Zinssatz
    average_capital = investment_cost / 2
    annual_interest = average_capital * project_interest_rate_db
    
    # DB III = DB II - Zinsen
    contribution_margin_3 = contribution_margin_2 - annual_interest
    
    # Debug-Ausgabe für DB-Berechnung
    print(f"  annual_revenue: {annual_revenue:.2f} EUR")
    print(f"  annual_variable_costs: {annual_variable_costs:.2f} EUR")
    print(f"  contribution_margin_1: {contribution_margin_1:.2f} EUR")
    print(f"  annual_depreciation: {annual_depreciation:.2f} EUR")
    print(f"  contribution_margin_2: {contribution_margin_2:.2f} EUR")
    print(f"  annual_interest: {annual_interest:.2f} EUR")
    print(f"  contribution_margin_3: {contribution_margin_3:.2f} EUR")
    
    # Berechne Summen über Projektlaufzeit mit Degradation
    total_db3_nominal = 0
    total_db3_present_value = 0
    
    for year in range(1, project_lifetime_years + 1):
        # Degradationsfaktor für dieses Jahr
        degradation_factor = (1 - annual_capacity_loss_percent / 100.0) ** (year - 1)
        
        # Nominal-Werte (mit Degradation)
        year_db3_nominal = contribution_margin_3 * degradation_factor
        total_db3_nominal += year_db3_nominal
        
        # Barwert-Werte (mit Degradation und Abzinsung)
        discount_factor = 1 / ((1 + discount_rate) ** year)
        year_db3_pv = year_db3_nominal * discount_factor
        total_db3_present_value += year_db3_pv
    
    # DB III pro kWh Speicherkapazität - KORRIGIERTE BERECHNUNG
    # DB3 pro kWh Berechnungen entfernt - redundant da gleiche Optima wie DB3 gesamt
    
    # Amortisationszeit berechnen
    payback_period = np.nan
    if contribution_margin_3 > 0:
        cumulative_db3 = 0
        for year in range(1, project_lifetime_years + 1):
            degradation_factor = (1 - annual_capacity_loss_percent / 100.0) ** (year - 1)
            year_db3 = contribution_margin_3 * degradation_factor
            cumulative_db3 += year_db3
            
            if cumulative_db3 >= investment_cost:
                if year == 1:
                    payback_period = investment_cost / year_db3
                else:
                    previous_cumulative = cumulative_db3 - year_db3
                    remaining_investment = investment_cost - previous_cumulative
                    fraction_of_year = remaining_investment / year_db3
                    payback_period = (year - 1) + fraction_of_year
                break
        else:
            payback_period = 999  # Kein Break-even gefunden
    else:
        payback_period = 999  # Negative Ersparnisse
    
    # ROI = Durchschnittlicher DB III / Investition × 100
    average_db3 = total_db3_nominal / project_lifetime_years
    roi_percentage = (average_db3 / investment_cost * 100) if investment_cost > 0 else 0
    
    # Debug-Ausgabe deaktiviert nach Korrektur der Amortisationszeit-Berechnung
    # print(f"🔍 DIAGNOSE DB-Berechnung für {battery_capacity_kwh} kWh Batterie:")
    # print(f"  💰 Investition: {investment_cost:.0f} EUR")
    # print(f"  💵 DB III (jährlich): {contribution_margin_3:.0f} EUR")
    # print(f"  ⏰ Amortisationszeit: {payback_period:.1f} Jahre")
    
    return {
        'investment_cost': investment_cost,
        'annual_revenue': annual_revenue,
        'annual_variable_costs': annual_variable_costs,
        'contribution_margin_1': contribution_margin_1,
        'annual_depreciation': annual_depreciation,
        'contribution_margin_2': contribution_margin_2,
        'annual_interest': annual_interest,
        'contribution_margin_3': contribution_margin_3,
        'total_db3_nominal': total_db3_nominal,
        'total_db3_present_value': total_db3_present_value,
        'payback_period_years': payback_period,
        'roi_percentage': roi_percentage
    }

def compute_energy_axis_range(data: pd.DataFrame, padding_ratio: float = 0.1, min_padding: float = 0.5) -> tuple[float, float]:
    """
    Bestimmt einen geeigneten Wertebereich für die Energie-Achse basierend auf den sichtbaren Daten.
    """
    if data is None or data.empty:
        return (-1.0, 1.0)

    positive_cols = ['Consumption_kWh', 'PV_Generation_kWh', 'Grid_Import_kWh', 'Battery_Charge_kWh']
    negative_cols = ['Grid_Export_kWh', 'Battery_Discharge_kWh']

    pos_values = [data[col].max() for col in positive_cols if col in data.columns]
    neg_values = [-(data[col].max()) for col in negative_cols if col in data.columns]

    pos_max = max([v for v in pos_values if pd.notna(v)] + [0])
    neg_min = min([v for v in neg_values if pd.notna(v)] + [0])

    if pos_max == neg_min:
        pos_max += 1.0
        neg_min -= 1.0

    span = max(pos_max - neg_min, 1.0)
    padding = max(span * padding_ratio, min_padding)

    return (neg_min - padding, pos_max + padding)


def plot_energy_flows_for_period(time_series_data: pd.DataFrame, start_date: str, end_date: str, battery_capacity_kwh: float = 10.0):
    """
    Visualisiert die stündlichen Energieflüsse mit interaktiven Zeitfenster-Optionen.

    Args:
        time_series_data (pd.DataFrame): DataFrame mit den Zeitreihen-Ergebnissen der Simulation.
        start_date (str): Startdatum im Format 'YYYY-MM-DD'.
        end_date (str): Enddatum im Format 'YYYY-MM-DD'.
    """
    period_data = time_series_data.loc[start_date:end_date]
    full_y_range = compute_energy_axis_range(period_data)
    
    # Verwende die Daten direkt ohne problematische Zeitstempel-Konvertierung
    print(f"Debug: Verwende {len(period_data)} Datenpunkte direkt")

    fig = make_subplots(rows=2, cols=1, 
                        shared_xaxes=True, 
                        vertical_spacing=0.15,  # Mehr Abstand zwischen den Subplots
                        row_heights=[0.7, 0.3],
                        subplot_titles=('Energieflüsse', 'Batterie Ladezustand'))

    # Energieflüsse (obere Grafik)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['Consumption_kWh'], 
                             mode='lines', name='Verbrauch', line=dict(color='red', width=2)),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['PV_Generation_kWh'], 
                             mode='lines', name='PV-Erzeugung', line=dict(color='green', width=2)),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['Grid_Import_kWh'], 
                             mode='lines', name='Netzbezug', fill='tozeroy', 
                             line=dict(color='blue', width=0), fillcolor='rgba(0,0,255,0.3)'),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=-period_data['Grid_Export_kWh'], # Negativ für Einspeisung
                             mode='lines', name='Netzeinspeisung', fill='tozeroy', 
                             line=dict(color='orange', width=0), fillcolor='rgba(255,165,0,0.3)'),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['Battery_Charge_kWh'], 
                             mode='lines', name='Batterie Ladung', line=dict(color='purple', dash='dot', width=2)),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=-period_data['Battery_Discharge_kWh'], # Negativ für Entladung
                             mode='lines', name='Batterie Entladung', line=dict(color='brown', dash='dot', width=2)),
                  row=1, col=1)

    # Batterie Ladezustand (untere Grafik) - Konvertiere zu Prozent
    soc_percent = (period_data['SOC_kWh'] / battery_capacity_kwh * 100).clip(0, 100)
    
    fig.add_trace(go.Scatter(x=period_data.index, y=soc_percent, 
                             mode='lines', name='Batterie SoC (%)', line=dict(color='darkgreen', width=2)),
                  row=2, col=1)

    # Erstelle Monatsansichten für Dropdown
    monthly_buttons = []
    start_date_dt = pd.to_datetime(start_date)
    end_date_dt = pd.to_datetime(end_date)
    
    # Gesamter Zeitraum
    full_range_args = {
        "xaxis.range": [start_date_dt, end_date_dt],
        "xaxis2.range": [start_date_dt, end_date_dt],
        "yaxis.range": [full_y_range[0], full_y_range[1]],
        "yaxis.autorange": False,
        "yaxis2.range": [0, 100],
        "yaxis2.autorange": False,
    }
    monthly_buttons.append(dict(
        args=[full_range_args, {"xaxis.rangeslider.range": [start_date_dt, end_date_dt]}],
        label="📅 Gesamter Zeitraum",
        method="relayout"
    ))
    
    # Einzelne Monate
    current_date = start_date_dt.replace(day=1)
    while current_date <= end_date_dt:
        month_end = (current_date + pd.DateOffset(months=1)) - pd.DateOffset(days=1)
        if month_end > end_date_dt:
            month_end = end_date_dt
            
        month_data = period_data.loc[current_date:month_end]
        if month_data.empty:
            current_date += pd.DateOffset(months=1)
            continue

        month_y_range = compute_energy_axis_range(month_data)
        month_range_args = {
            "xaxis.range": [current_date, month_end],
            "xaxis2.range": [current_date, month_end],
            "yaxis.range": [month_y_range[0], month_y_range[1]],
            "yaxis.autorange": False,
            "yaxis2.range": [0, 100],
            "yaxis2.autorange": False,
        }
        monthly_buttons.append(dict(
            args=[month_range_args, {"xaxis.rangeslider.range": [current_date, month_end]}],
            label=f"📅 {current_date.strftime('%B %Y')}",
            method="relayout"
        ))
        current_date += pd.DateOffset(months=1)

    # Interaktive Zeitfenster-Optionen (Plotly Range Slider + Dropdown)
    fig.update_layout(
        title_text='Energieflüsse - Interaktive Zeitfenster-Auswahl', 
        height=800,
        autosize=True,
        margin=dict(l=60, r=40, t=90, b=80),
        showlegend=True,
        xaxis=dict(
            rangeslider=dict(
                visible=True,
                thickness=0.08,
                range=[start_date_dt, end_date_dt]
            ),
            type="date"
        ),
        uirevision="energy_flow_period"
    )
    
    if monthly_buttons:
        fig.update_layout(
            updatemenus=[
                dict(
                    type="dropdown",
                    direction="down",
                    buttons=monthly_buttons,
                    x=0.0,
                    xanchor="left",
                    y=1.12,
                    yanchor="bottom",
                    showactive=True
                )
            ]
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    # Y-Achsen Labels
    fig.update_yaxes(title_text="Energie (kWh)", row=1, col=1, range=[full_y_range[0], full_y_range[1]])
    fig.update_yaxes(title_text="Batterie SoC (%)", row=2, col=1, range=[0, 100])
    
    # X-Achsen Labels (Plotly formatiert automatisch)
    fig.update_xaxes(title_text="Datum", row=2, col=1)
    fig.update_xaxes(title_text="Datum", row=1, col=1)

    return fig

def plot_monthly_energy_flows(time_series_data: pd.DataFrame, year: int, battery_capacity_kwh: float = 10.0):
    """
    Erstellt separate Grafiken für jeden Monat des Jahres.
    
    Args:
        time_series_data (pd.DataFrame): DataFrame mit den Zeitreihen-Ergebnissen der Simulation.
        year (int): Jahr für die Monatsgrafiken.
    
    Returns:
        dict: Dictionary mit Monatsgrafiken, Schlüssel sind Monatsnamen.
    """
    monthly_figures = {}
    
    for month in range(1, 13):
        # Start- und Enddatum für den Monat
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year}-{month:02d}-31"
        else:
            # Verwende den letzten Tag des aktuellen Monats
            next_month = month + 1
            if next_month > 12:
                next_year = year + 1
                next_month = 1
            else:
                next_year = year
            end_date = f"{next_year}-{next_month:02d}-01"
        
        # Monatsdaten filtern - verwende exklusive Filterung für end_date
        try:
            month_data = time_series_data.loc[start_date:end_date]
            # Entferne den ersten Tag des nächsten Monats falls er enthalten ist
            month_data = month_data[month_data.index.month == month]
        except:
            # Fallback: Filtere nach Monat und Jahr
            month_data = time_series_data[
                (time_series_data.index.year == year) & 
                (time_series_data.index.month == month)
            ]
        
        # Verwende die Daten direkt ohne problematische Zeitstempel-Konvertierung
        print(f"Debug: Monat {month} - Verwende {len(month_data)} Datenpunkte direkt")
        
        if len(month_data) > 0:  # Nur erstellen wenn Daten vorhanden
            print(f"Debug: Monat {month} - {len(month_data)} Datenpunkte gefunden")
            print(f"Debug: Index-Typ: {type(month_data.index)}")
            print(f"Debug: Erste 3 Zeitstempel: {month_data.index[:3]}")
            print(f"Debug: Jahr-Bereich: {month_data.index.year.min()} - {month_data.index.year.max()}")
            # Erstelle Grafik für diesen Monat
            fig = make_subplots(rows=2, cols=1, 
                               shared_xaxes=True, 
                               vertical_spacing=0.15,
                               row_heights=[0.7, 0.3],
                               subplot_titles=(f'Energieflüsse - {pd.to_datetime(start_date).strftime("%B %Y")}', 
                                             'Batterie Ladezustand'))

            # Energieflüsse (obere Grafik)
            fig.add_trace(go.Scatter(x=month_data.index, y=month_data['Consumption_kWh'], 
                                     mode='lines', name='Verbrauch', line=dict(color='red', width=2)),
                          row=1, col=1)
            fig.add_trace(go.Scatter(x=month_data.index, y=month_data['PV_Generation_kWh'], 
                                     mode='lines', name='PV-Erzeugung', line=dict(color='green', width=2)),
                          row=1, col=1)
            fig.add_trace(go.Scatter(x=month_data.index, y=month_data['Grid_Import_kWh'], 
                                     mode='lines', name='Netzbezug', fill='tozeroy', 
                                     line=dict(color='blue', width=0), fillcolor='rgba(0,0,255,0.3)'),
                          row=1, col=1)
            fig.add_trace(go.Scatter(x=month_data.index, y=-month_data['Grid_Export_kWh'], 
                                     mode='lines', name='Netzeinspeisung', fill='tozeroy', 
                                     line=dict(color='orange', width=0), fillcolor='rgba(255,165,0,0.3)'),
                          row=1, col=1)
            fig.add_trace(go.Scatter(x=month_data.index, y=month_data['Battery_Charge_kWh'], 
                                     mode='lines', name='Batterie Ladung', line=dict(color='purple', dash='dot', width=2)),
                          row=1, col=1)
            fig.add_trace(go.Scatter(x=month_data.index, y=-month_data['Battery_Discharge_kWh'], 
                                     mode='lines', name='Batterie Entladung', line=dict(color='brown', dash='dot', width=2)),
                          row=1, col=1)

            # Batterie Ladezustand (untere Grafik) - Konvertiere zu Prozent
            soc_percent = (month_data['SOC_kWh'] / battery_capacity_kwh * 100).clip(0, 100)
            
            fig.add_trace(go.Scatter(x=month_data.index, y=soc_percent, 
                                     mode='lines', name='Batterie SoC (%)', line=dict(color='darkgreen', width=2)),
                          row=2, col=1)

            # Modernes Layout für Monatsgrafik
            fig.update_layout(
                title_text=f'Energieflüsse - {pd.to_datetime(start_date).strftime("%B %Y")}', 
                height=500,
                showlegend=True
            )
            
            # Modernes Theme anwenden
            fig = apply_modern_plotly_theme(fig)

            # Y-Achsen Labels
            fig.update_yaxes(title_text="Energie (kWh)", row=1, col=1)
            fig.update_yaxes(title_text="Batterie SoC (%)", row=2, col=1, range=[0, 100])
            
            # X-Achsen Labels (Plotly formatiert automatisch)
            fig.update_xaxes(title_text="Datum", row=2, col=1)
            fig.update_xaxes(title_text="Datum", row=1, col=1)

            # Speichere Grafik mit Monatsnamen als Schlüssel
            month_name = pd.to_datetime(start_date).strftime("%B")
            monthly_figures[month_name] = fig
    
    return monthly_figures

def plot_technical_optimization_curve(optimization_results: list):
    """
    Visualisiert die technischen Kennzahlen der Optimierungskurve.

    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)

    # Erstelle einzelne Grafik für technische Kennzahlen
    fig = go.Figure()
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(2)

    # Autarkiegrad
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["autarky_rate"] * 100, 
        mode='lines+markers', 
        name='Autarkiegrad (%)', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ))
    
    # Eigenverbrauchsquote
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["self_consumption_rate"] * 100, 
        mode='lines+markers', 
        name='Eigenverbrauchsquote (%)', 
        line=dict(color=colors[1], dash='dot', width=2),
        marker=dict(size=6)
    ))

    # Modernes Layout
    fig.update_layout(
        title_text='Technische Kennzahlen - Optimierung der Batteriespeichergröße',
        xaxis_title='Batteriekapazität (kWh)',
        yaxis_title='Prozent (%)',
        height=600,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

    # Layout für technische Kennzahlen
    fig.update_layout(
        title_text='Technische Kennzahlen - Optimierung der Batteriespeichergröße',
        hovermode='x unified',
        height=600,
        autosize=True,
        margin=dict(l=20, r=20, t=80, b=20),
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(color='black'),
        title_font=dict(color='black'),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color='black')
        )
    )
    
    # Achsenbeschriftungen
    fig.update_xaxes(title_text='Batteriekapazität (kWh)', title_font=dict(color='black'), tickfont=dict(color='black'))
    fig.update_yaxes(title_text='Prozent (%)', title_font=dict(color='black'), tickfont=dict(color='black'))
    
    # Y-Achsen-Bereiche anpassen - feinere Skalierung
    fig.update_yaxes(range=[0, 100])
    fig.update_yaxes(dtick=10)  # Feine Skalierung alle 10%
    
    # Download-Button hinzufügen
    fig.update_layout(
        modebar_add=[
            "downloadImage"
        ]
    )
    
    return fig

def plot_economic_optimization_curve(optimization_results: list):
    """
    Visualisiert die Deckungsbeitrag III (DB3) Kennzahlen der Optimierungskurve.

    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)

    # Erstelle einzelne Grafik für Deckungsbeitrag III Kennzahlen
    fig = go.Figure()
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(4)
    
    # DB3 Nominal
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["total_db3_nominal"], 
        mode='lines+markers', 
        name='DB III gesamt (Nominal)', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ))
    
    # DB3 Barwert
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["total_db3_present_value"], 
        mode='lines+markers', 
        name='DB III gesamt (Barwert)', 
        line=dict(color=colors[1], width=3, dash='dash'),
        marker=dict(size=6)
    ))
    
    # Markiere nicht wirtschaftliche Investitionen (basierend auf Cash Flow-Amortisationszeit)
    non_economic = df_results[df_results['payback_period_years'] >= 999]
    if not non_economic.empty:
        fig.add_trace(go.Scatter(
            x=non_economic["battery_capacity_kwh"], 
            y=[-1000] * len(non_economic),  # Zeige bei -1000 € um Überlagerung zu vermeiden
            mode='markers', 
            name='Nicht wirtschaftlich', 
            marker=dict(color=colors[2], size=8, symbol='x'),
            showlegend=True
        ))

    # === BREAK-EVEN-PUNKT MARKIEREN ===
    # Finde den Punkt, wo DB3 Barwert = 0 (Break-even)
    positive_db3 = df_results[df_results['total_db3_present_value'] > 0]
    if not positive_db3.empty:
        # Erste positive DB3 (Break-even-Punkt)
        break_even_idx = positive_db3.index[0]
        break_even_capacity = df_results.loc[break_even_idx, 'battery_capacity_kwh']
        break_even_db3 = df_results.loc[break_even_idx, 'total_db3_present_value']
        
        # Markiere Break-even-Punkt
        fig.add_vline(
            x=break_even_capacity, 
            line_dash="dash", 
            line_color="orange",
            line_width=2
        )
        
        # Annotation für Break-even-Punkt
        fig.add_annotation(
            x=break_even_capacity,
            y=break_even_db3,
            text=f"Break-even:<br>{break_even_capacity:.1f} kWh",
            showarrow=True,
            arrowhead=2,
            arrowcolor="orange",
            arrowwidth=2,
            bgcolor="rgba(255,165,0,0.8)",
            bordercolor="orange",
            borderwidth=1
        )

    # === OPTIMALER PUNKT MARKIEREN ===
    best_db3_idx = df_results['total_db3_present_value'].idxmax()
    best_capacity = df_results.loc[best_db3_idx, 'battery_capacity_kwh']
    best_db3 = df_results.loc[best_db3_idx, 'total_db3_present_value']
    
    # Markiere optimalen Punkt
    fig.add_vline(
        x=best_capacity, 
        line_dash="dot", 
        line_color="purple",
        line_width=2
    )
    
    # Annotation für optimalen Punkt
    fig.add_annotation(
        x=best_capacity,
        y=best_db3,
        text=f"Optimal:<br>{best_capacity:.1f} kWh<br>DB3: {best_db3:.0f}€",
        showarrow=True,
        arrowhead=2,
        arrowcolor="purple",
        arrowwidth=2,
        bgcolor="rgba(128,0,128,0.8)",
        bordercolor="purple",
        borderwidth=1
    )

    # === MODERNES LAYOUT ===
    fig.update_layout(
        title_text='Deckungsbeitrag III (DB3) - Optimierung der Batteriespeichergröße',
        xaxis_title='Batteriekapazität (kWh)',
        yaxis_title='Deckungsbeitrag III (€)',
        height=600,
        showlegend=True
    )
    
    # Moderne Erklärung hinzufügen
    fig.add_annotation(
        text="💡 Break-even: Punkt, wo DB3 Barwert = 0<br>💡 Optimal: Punkt mit höchstem DB3 Barwert<br>💡 DB3: Deckungsbeitrag III nach Abzug aller Kosten",
        xref="paper", yref="paper",
        x=0.02, y=0.98,
        xanchor="left", yanchor="top",
        showarrow=False,
        bgcolor="rgba(0,0,0,0.8)",
        bordercolor="#FF8C42",
        borderwidth=1,
        font=dict(color="white", size=10)
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_optimization_curve(optimization_results: list):
    """
    Visualisiert die komplette Optimierungskurve für die Speichergröße (kombiniert).

    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)

    # Erstelle zwei Subplots: einen für technische Kennzahlen, einen für wirtschaftliche
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=('Technische Kennzahlen', 'Wirtschaftliche Kennzahlen'),
        vertical_spacing=0.15,
        row_heights=[0.5, 0.5]
    )
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(6)

    # === OBERER GRAPH: Technische Kennzahlen ===
    # Autarkiegrad
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["autarky_rate"] * 100, 
        mode='lines+markers', 
        name='Autarkiegrad (%)', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ), row=1, col=1)
    
    # Eigenverbrauchsquote
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["self_consumption_rate"] * 100, 
        mode='lines+markers', 
        name='Eigenverbrauchsquote (%)', 
        line=dict(color=colors[1], dash='dot', width=2),
        marker=dict(size=6)
    ), row=1, col=1)

    # === UNTERER GRAPH: Wirtschaftliche Kennzahlen ===
    # NPV
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["npv"], 
        mode='lines+markers', 
        name='NPV (€)', 
        line=dict(color=colors[2], width=3),
        marker=dict(size=6)
    ), row=2, col=1)
    
    # Amortisationszeit - Filtere unrealistische Werte
    valid_payback = df_results[df_results['payback_period_years'] < 100]  # Filtere unrealistische Werte
    if not valid_payback.empty:
        fig.add_trace(go.Scatter(
            x=valid_payback["battery_capacity_kwh"], 
            y=valid_payback["payback_period_years"], 
            mode='lines+markers', 
            name='Amortisationszeit (Jahre)', 
            line=dict(color=colors[3], width=2),
            marker=dict(size=6)
        ), row=2, col=1)
    
    # Markiere nicht wirtschaftliche Investitionen (basierend auf Cash Flow-Amortisationszeit)
    non_economic = df_results[df_results['payback_period_years'] >= 999]
    if not non_economic.empty:
        fig.add_trace(go.Scatter(
            x=non_economic["battery_capacity_kwh"], 
            y=[-0.5] * len(non_economic),  # Zeige bei -0.5 statt 0 um Überlagerung zu vermeiden
            mode='markers', 
            name='Nicht wirtschaftlich', 
            marker=dict(color=colors[4], size=8, symbol='x'),
            showlegend=True
        ), row=2, col=1)

    # === BREAK-EVEN-PUNKT MARKIEREN ===
    # Finde den Punkt, wo NPV = 0 (Break-even)
    positive_npv = df_results[df_results['npv'] > 0]
    if not positive_npv.empty:
        # Erste positive NPV (Break-even-Punkt)
        break_even_idx = positive_npv.index[0]
        break_even_capacity = df_results.loc[break_even_idx, 'battery_capacity_kwh']
        break_even_npv = df_results.loc[break_even_idx, 'npv']
        
        # Markiere Break-even-Punkt
        fig.add_vline(
            x=break_even_capacity, 
            line_dash="dash", 
            line_color="orange",
            line_width=2,
            row=2, col=1
        )
        
        # Annotation für Break-even-Punkt
        fig.add_annotation(
            x=break_even_capacity,
            y=break_even_npv,
            text=f"Break-even:<br>{break_even_capacity:.1f} kWh",
            showarrow=True,
            arrowhead=2,
            arrowcolor="orange",
            arrowwidth=2,
            bgcolor="rgba(255,165,0,0.8)",
            bordercolor="orange",
            borderwidth=1,
            row=2, col=1
        )

    # === OPTIMALER PUNKT MARKIEREN ===
    best_npv_idx = df_results['npv'].idxmax()
    best_capacity = df_results.loc[best_npv_idx, 'battery_capacity_kwh']
    best_npv = df_results.loc[best_npv_idx, 'npv']
    
    # Markiere optimalen Punkt
    fig.add_vline(
        x=best_capacity, 
        line_dash="dot", 
        line_color="purple",
        line_width=2,
        row=2, col=1
    )
    
    # Annotation für optimalen Punkt
    fig.add_annotation(
        x=best_capacity,
        y=best_npv,
        text=f"Optimal:<br>{best_capacity:.1f} kWh<br>NPV: {best_npv:.0f}€",
        showarrow=True,
        arrowhead=2,
        arrowcolor="purple",
        arrowwidth=2,
        bgcolor="rgba(128,0,128,0.8)",
        bordercolor="purple",
        borderwidth=1,
        row=2, col=1
    )

    # === MODERNES LAYOUT ===
    fig.update_layout(
        title_text='Optimierung der Batteriespeichergröße - Break-even & Optimaler Punkt',
        height=800,
        showlegend=True
    )
    
    # Achsenbeschriftungen
    fig.update_xaxes(title_text='Batteriekapazität (kWh)', row=2, col=1)
    fig.update_yaxes(title_text='Autarkiegrad / Eigenverbrauchsquote (%)', row=1, col=1)
    fig.update_yaxes(title_text='NPV (€) / Amortisationszeit (Jahre)', row=2, col=1)
    
    # Y-Achsen-Bereiche anpassen
    fig.update_yaxes(range=[0, 100], row=1, col=1)  # Prozent für technische Kennzahlen
    
    # Moderne Erklärung hinzufügen
    fig.add_annotation(
        text="💡 Break-even: Punkt, wo sich die Anlage selbst finanziert hat<br>💡 Optimal: Punkt mit höchstem NPV<br>💡 Amortisationszeit: Zeit bis zur Selbstfinanzierung",
        xref="paper", yref="paper",
        x=0.02, y=0.98,
        xanchor="left", yanchor="top",
        showarrow=False,
        bgcolor="rgba(0,0,0,0.8)",
        bordercolor="#FF8C42",
        borderwidth=1,
        font=dict(color="white", size=10)
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_energy_flows_optimization(optimization_results: list):
    """
    Erstellt eine separate Grafik für Energieflüsse in Abhängigkeit der Speichergröße.
    
    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)
    
    # Erstelle einzelne Grafik für Energieflüsse
    fig = go.Figure()
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(5)
    
    # PV-Erzeugung (konstant)
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["pv_generation_kwh"], 
        mode='lines+markers', 
        name='PV-Erzeugung', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ))
    
    # Stromverbrauch (konstant)
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["total_consumption_kwh"], 
        mode='lines+markers', 
        name='Stromverbrauch', 
        line=dict(color=colors[1], width=3),
        marker=dict(size=6)
    ))
    
    # Selbstgenutzter Strom
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["self_consumption_kwh"], 
        mode='lines+markers', 
        name='Selbstgenutzter Strom', 
        line=dict(color=colors[2], width=3),
        marker=dict(size=6)
    ))
    
    # Eingespeister Strom
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["grid_export_kwh"], 
        mode='lines+markers', 
        name='Eingespeister Strom', 
        line=dict(color=colors[3], width=3),
        marker=dict(size=6)
    ))
    
    # Netzbezug
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["grid_import_kwh"], 
        mode='lines+markers', 
        name='Netzbezug', 
        line=dict(color=colors[4], width=3),
        marker=dict(size=6)
    ))
    
    # Modernes Layout
    fig.update_layout(
        title_text='Energieflüsse in Abhängigkeit der Speichergröße',
        xaxis_title='Batteriekapazität (kWh)',
        yaxis_title='Energie (kWh/Jahr)',
        height=600,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_battery_losses_optimization(optimization_results: list):
    """
    Erstellt eine separate Grafik für Speicherverluste in Abhängigkeit der Speichergröße.
    
    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)
    
    # Erstelle einzelne Grafik für Speicherverluste
    fig = go.Figure()
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(1)
    
    # Speicherverluste
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["battery_losses_kwh"], 
        mode='lines+markers', 
        name='Speicherverluste', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ))
    
    # Modernes Layout
    fig.update_layout(
        title_text='Speicherverluste in Abhängigkeit der Speichergröße',
        xaxis_title='Batteriekapazität (kWh)',
        yaxis_title='Verluste (kWh/Jahr)',
        height=600,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_economic_optimization_extended(optimization_results: list):
    """
    Erstellt eine separate Grafik für wirtschaftliche Kennzahlen in Abhängigkeit der Speichergröße.
    
    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)
    
    # Erstelle einzelne Grafik für wirtschaftliche Kennzahlen
    fig = go.Figure()
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(3)
    
    # NPV
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["npv"], 
        mode='lines+markers', 
        name='NPV (€)', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ))
    
    # Amortisationszeit - Filtere unrealistische Werte
    valid_payback = df_results[df_results['payback_period_years'] < 100]
    if not valid_payback.empty:
        fig.add_trace(go.Scatter(
            x=valid_payback["battery_capacity_kwh"], 
            y=valid_payback["payback_period_years"], 
            mode='lines+markers', 
            name='Amortisationszeit (Jahre)', 
            line=dict(color=colors[1], dash='dot', width=2),
            marker=dict(size=6)
        ))
    
    # Modernes Layout
    fig.update_layout(
        title_text='Wirtschaftliche Kennzahlen in Abhängigkeit der Speichergröße',
        xaxis_title='Batteriekapazität (kWh)',
        yaxis_title='NPV (€) / Amortisation (Jahre)',
        height=600,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_extended_optimization_curves(optimization_results: list):
    """
    Erstellt erweiterte Optimierungsgrafiken mit allen gewünschten Parametern (kombiniert).
    
    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung.
    """
    df_results = pd.DataFrame(optimization_results)
    
    # Erstelle 3 Subplots: Energieflüsse, Verluste, und Wirtschaftlichkeit
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=('Energieflüsse (kWh/Jahr)', 'Speicherverluste (kWh/Jahr)', 'Wirtschaftliche Kennzahlen'),
        vertical_spacing=0.12,
        row_heights=[0.4, 0.3, 0.3]
    )
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(8)
    
    # === MITTLERER GRAPH: Speicherverluste ===
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["battery_losses_kwh"], 
        mode='lines+markers', 
        name='Speicherverluste', 
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ), row=2, col=1)
    
    # === UNTERER GRAPH: Wirtschaftliche Kennzahlen ===
    # NPV
    fig.add_trace(go.Scatter(
        x=df_results["battery_capacity_kwh"], 
        y=df_results["npv"], 
        mode='lines+markers', 
        name='NPV (€)', 
        line=dict(color='darkblue', width=3),
        marker=dict(size=6)
    ), row=3, col=1)
    
    # Amortisationszeit - Filtere unrealistische Werte
    valid_payback = df_results[df_results['payback_period_years'] < 100]
    if not valid_payback.empty:
        fig.add_trace(go.Scatter(
            x=valid_payback["battery_capacity_kwh"], 
            y=valid_payback["payback_period_years"], 
            mode='lines+markers', 
            name='Amortisationszeit (Jahre)', 
            line=dict(color='darkred', dash='dot', width=2),
            marker=dict(size=6),
            yaxis='y3'
        ), row=3, col=1)
    
    # === LAYOUT VERBESSERN ===
    fig.update_layout(
        title_text='Erweiterte Optimierungsanalyse - Energieflüsse, Verluste und Wirtschaftlichkeit',
        hovermode='x unified',
        height=1000,
        autosize=True,
        margin=dict(l=20, r=20, t=80, b=20),
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(color='black'),
        title_font=dict(color='black'),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color='black')
        )
    )
    
    # Achsenbeschriftungen
    fig.update_xaxes(title_text='Batteriekapazität (kWh)', row=3, col=1, title_font=dict(color='black'), tickfont=dict(color='black'))
    fig.update_yaxes(title_text='Energie (kWh/Jahr)', row=1, col=1, title_font=dict(color='black'), tickfont=dict(color='black'))
    fig.update_yaxes(title_text='Verluste (kWh/Jahr)', row=2, col=1, title_font=dict(color='black'), tickfont=dict(color='black'))
    fig.update_yaxes(title_text='NPV (€) / Amortisation (Jahre)', row=3, col=1, title_font=dict(color='black'), tickfont=dict(color='black'))
    
    # Feine Y-Achsen-Skalierung für bessere Lesbarkeit
    # Energieflüsse - dynamische Skalierung
    max_energy = max(
        df_results['pv_generation_kwh'].max(),
        df_results['total_consumption_kwh'].max(),
        df_results['self_consumption_kwh'].max(),
        df_results['grid_export_kwh'].max(),
        df_results['grid_import_kwh'].max()
    )
    fig.update_yaxes(
        dtick=max(100, max_energy/20),  # Feine Skalierung
        row=1, col=1
    )
    
    # Verluste - feine Skalierung
    max_losses = df_results['battery_losses_kwh'].max()
    fig.update_yaxes(
        dtick=max(10, max_losses/15),  # Feine Skalierung
        row=2, col=1
    )
    
    # Wirtschaftliche Kennzahlen - feine Skalierung
    max_npv = df_results['npv'].max()
    min_npv = df_results['npv'].min()
    npv_range = max_npv - min_npv
    fig.update_yaxes(
        dtick=max(100, npv_range/20),  # Feine Skalierung
        row=3, col=1
    )
    
    return fig

def plot_scenario_comparison(scenario_results: dict):
    """
    Visualisiert den Vergleich verschiedener Szenarien anhand von KPIs.

    Args:
        scenario_results (dict): Dictionary mit den Ergebnissen verschiedener Szenarien.
                                 Format: {"Szenario Name": {"kpis": {...}}}
    """
    kpis_to_compare = ["autarky_rate", "self_consumption_rate", "payback_period_years", "npv"]
    kpi_labels = {
        "autarky_rate": "Autarkiegrad (%)",
        "self_consumption_rate": "Eigenverbrauchsquote (%)",
        "payback_period_years": "Amortisationszeit (Jahre)",
        "npv": "NPV (€)"
    }

    scenario_names = list(scenario_results.keys())
    data = []

    for kpi in kpis_to_compare:
        values = [scenario_results[name]["kpis"][kpi] for name in scenario_names]
        data.append(go.Bar(name=kpi_labels[kpi], x=scenario_names, y=values))

    fig = go.Figure(data=data)
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(len(kpis_to_compare))
    for i, trace in enumerate(fig.data):
        trace.marker.color = colors[i]
    
    # Modernes Layout
    fig.update_layout(
        barmode='group', 
        title_text='Szenarien-Vergleich der KPIs',
        height=600,
        showlegend=True
    )
    
    # Achsenbeschriftungen
    fig.update_yaxes(title_text='Wert')
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    return fig

def plot_sankey_diagram(kpis: dict):
    """
    Erstellt ein verbessertes Sankey-Diagramm der jährlichen Energieflüsse.
    Zeigt klar: Gesamtverbrauch, Gesamterzeugung, Anteil aus Batterie, Netzbezug und Eigenverbrauch.
    Batterieverluste werden sowohl beim Laden als auch beim Entladen separat dargestellt.

    Args:
        kpis (dict): Dictionary mit den aggregierten KPIs der Simulation.
    """
    # Definition der Knoten (Quellen und Senken)
    labels = [
        "PV-Erzeugung", "Netzbezug", # Quellen
        "Verbrauch", # Senke
        "Direkter Eigenverbrauch", "Batterie Ladung", "Netzeinspeisung", # Zwischenstationen
        "Batterie Entladung", "Ladeverluste", "Entladeverluste" # Batterie-Komponenten
    ]

    # Indizes für die Knoten
    pv_gen_idx = 0
    grid_import_idx = 1
    consumption_idx = 2
    direct_self_cons_idx = 3
    battery_charge_idx = 4
    grid_export_idx = 5
    battery_discharge_idx = 6
    charge_losses_idx = 7
    discharge_losses_idx = 8

    # Berechnung der Flusswerte aus den tatsächlichen Simulationsdaten
    total_pv_generation = kpis["total_pv_generation_kwh"]
    total_consumption = kpis["total_consumption_kwh"]
    total_grid_import = kpis["total_grid_import_kwh"]
    total_grid_export = kpis["total_grid_export_kwh"]
    total_direct_self_consumption = kpis["total_direct_self_consumption_kwh"]
    total_battery_charge = kpis["total_battery_charge_kwh"]
    total_battery_discharge = kpis["total_battery_discharge_kwh"]
    
    # Batterieeffizienzen aus KPIs oder Standardwerte
    battery_efficiency_charge = kpis.get('battery_efficiency_charge', 0.95)
    battery_efficiency_discharge = kpis.get('battery_efficiency_discharge', 0.95)
    
    # Verluste beim Laden und Entladen separat berechnen
    charge_losses = total_battery_charge * (1 - battery_efficiency_charge)  # Verluste beim Laden
    discharge_losses = total_battery_discharge * (1 - battery_efficiency_discharge) / battery_efficiency_discharge  # Verluste beim Entladen
    
    # Nettoladung in Batterie (ohne Verluste)
    net_charge_to_battery = total_battery_charge * battery_efficiency_charge

    # Quellen und Ziele der Flüsse
    source = [
        pv_gen_idx, pv_gen_idx, pv_gen_idx,  # PV-Erzeugung
        grid_import_idx,  # Netzbezug
        direct_self_cons_idx,  # Direkter Eigenverbrauch -> Verbrauch
        battery_charge_idx, battery_charge_idx,  # Batterie Ladung (2 Flüsse)
        battery_discharge_idx, battery_discharge_idx  # Batterie Entladung (2 Flüsse)
    ]
    target = [
        direct_self_cons_idx, battery_charge_idx, grid_export_idx,  # PV-Erzeugung
        consumption_idx,  # Netzbezug
        consumption_idx,  # Direkter Eigenverbrauch -> Verbrauch
        charge_losses_idx, battery_discharge_idx,  # Batterie Ladung -> Ladeverluste + Batterie Entladung
        discharge_losses_idx, consumption_idx  # Batterie Entladung -> Entladeverluste + Verbrauch
    ]
    value = [
        total_direct_self_consumption,  # PV -> Direkter Eigenverbrauch
        total_battery_charge,  # PV -> Batterie Ladung
        total_grid_export,  # PV -> Netzeinspeisung
        total_grid_import,  # Netzbezug -> Verbrauch
        total_direct_self_consumption,  # Direkter Eigenverbrauch -> Verbrauch
        charge_losses,  # Batterie Ladung -> Ladeverluste
        net_charge_to_battery,  # Batterie Ladung -> Batterie Entladung (Netto)
        discharge_losses,  # Batterie Entladung -> Entladeverluste
        total_battery_discharge  # Batterie Entladung -> Verbrauch
    ]

    # Filter out zero values to avoid drawing unnecessary links
    filtered_source = []
    filtered_target = []
    filtered_value = []
    for s, t, v in zip(source, target, value):
        if v > 0.01:  # Schwellenwert, um sehr kleine Flüsse zu ignorieren
            filtered_source.append(s)
            filtered_target.append(t)
            filtered_value.append(v)

    # Farben für die Knoten - verbesserte Kontraste
    node_colors = [
        "#2E8B57",  # PV-Erzeugung (Dunkelgrün)
        "#4169E1",  # Netzbezug (Königsblau)
        "#DC143C",  # Verbrauch (Karmesinrot)
        "#32CD32",  # Direkter Eigenverbrauch (Limettengrün)
        "#9370DB",  # Batterie Ladung (Mittelviolett)
        "#FF8C00",  # Netzeinspeisung (Dunkelorange)
        "#8B4513",  # Batterie Entladung (Sattelbraun)
        "#FF6B6B",  # Ladeverluste (Hellrot)
        "#FFB347"   # Entladeverluste (Pfirsich)
    ]

    # Farben für die Energieflüsse - verschiedene Farben für bessere Unterscheidung
    link_colors = []
    for i, (s, t, v) in enumerate(zip(filtered_source, filtered_target, filtered_value)):
        if s == pv_gen_idx:  # Von PV-Erzeugung
            if t == direct_self_cons_idx:
                link_colors.append("rgba(50, 205, 50, 0.6)")  # Grün für direkten Eigenverbrauch
            elif t == battery_charge_idx:
                link_colors.append("rgba(147, 112, 219, 0.6)")  # Lila für Batterieladung
            elif t == grid_export_idx:
                link_colors.append("rgba(255, 140, 0, 0.6)")  # Orange für Netzeinspeisung
        elif s == grid_import_idx:  # Von Netzbezug
            if t == consumption_idx:
                link_colors.append("rgba(65, 105, 225, 0.6)")  # Blau für direkten Verbrauch
            elif t == battery_charge_idx:
                link_colors.append("rgba(147, 112, 219, 0.4)")  # Helllila für Batterieladung
        elif s == direct_self_cons_idx:  # Von direktem Eigenverbrauch
            link_colors.append("rgba(50, 205, 50, 0.8)")  # Grün für Verbrauch
        elif s == battery_charge_idx:  # Von Batterieladung
            if t == battery_discharge_idx:
                link_colors.append("rgba(139, 69, 19, 0.6)")  # Braun für Entladung
            elif t == charge_losses_idx:
                link_colors.append("rgba(255, 107, 107, 0.6)")  # Rot für Ladeverluste
        elif s == battery_discharge_idx:  # Von Batterieentladung
            if t == consumption_idx:
                link_colors.append("rgba(139, 69, 19, 0.8)")  # Braun für Verbrauch
            elif t == discharge_losses_idx:
                link_colors.append("rgba(255, 179, 71, 0.6)")  # Orange für Entladeverluste
        else:
            link_colors.append("rgba(128, 128, 128, 0.4)")  # Standard grau

    fig = go.Figure(data=[go.Sankey(
        node=dict(
            label=labels, 
            pad=25,  # Mehr Abstand zwischen Knoten
            thickness=30,  # Dickere Knoten für bessere Sichtbarkeit
            line=dict(color="black", width=1.5),  # Dickere Umrandung
            color=node_colors
        ),
        link=dict(
            source=filtered_source, 
            target=filtered_target, 
            value=filtered_value,
            color=link_colors,  # Individuelle Farben für jeden Fluss
            line=dict(width=0.8)  # Etwas dickere Linien für bessere Übersicht
        ),
        arrangement="snap"  # Bessere Anordnung der Knoten
    )])
    
    # Verbesserte Layout-Einstellungen - optimiert für bessere Lesbarkeit
    # Modernes Layout
    fig.update_layout(
        title_text="Jährliche Energieflüsse - Übersicht (mit separaten Lade- und Entladeverlusten)",
        height=800,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    # Berechnung der Kennzahlen für die detaillierte Übersicht
    total_energy_consumed = total_consumption
    total_energy_generated = total_pv_generation
    energy_from_battery = total_battery_discharge
    energy_from_grid = total_grid_import
    self_consumption = total_direct_self_consumption + total_battery_discharge
    total_battery_losses = charge_losses + discharge_losses
    
    # Download-Button hinzufügen
    fig.update_layout(
        modebar_add=[
            "downloadImage"
        ]
    )
    
    # Detaillierte Berechnungsübersicht zurückgeben
    calculation_details = {
        'total_consumption_kwh': total_consumption,
        'total_pv_generation_kwh': total_pv_generation,
        'total_direct_self_consumption_kwh': total_direct_self_consumption,
        'total_battery_charge_kwh': total_battery_charge,
        'total_battery_discharge_kwh': total_battery_discharge,
        'total_grid_import_kwh': total_grid_import,
        'total_grid_export_kwh': total_grid_export,
        'charge_losses_kwh': charge_losses,
        'discharge_losses_kwh': discharge_losses,
        'total_battery_losses_kwh': total_battery_losses,
        'self_consumption_kwh': self_consumption,
        'autarky_rate_percent': (total_consumption - total_grid_import) / total_consumption * 100 if total_consumption > 0 else 0,
        'self_consumption_rate_percent': self_consumption / total_pv_generation * 100 if total_pv_generation > 0 else 0,
        'battery_efficiency_charge': battery_efficiency_charge,
        'battery_efficiency_discharge': battery_efficiency_discharge
    }
    
    # Erstelle ein Dictionary mit Figure und calculation_details
    result = {
        'figure': fig,
        'calculation_details': calculation_details
    }
    
    return result

def plot_sankey_diagram_no_battery(kpis_no_battery: dict, total_consumption: float, total_pv_generation: float):
    """
    Erstellt ein Sankey-Diagramm der jährlichen Energieflüsse OHNE Batteriespeicher.
    Vereinfachte Version ohne Batterie-Komponenten.

    Args:
        kpis_no_battery (dict): Dictionary mit den KPIs ohne Batterie (grid_import, grid_export, self_consumption).
        total_consumption (float): Gesamtverbrauch in kWh
        total_pv_generation (float): Gesamt-PV-Erzeugung in kWh
    """
    # Definition der Knoten (Quellen und Senken) - ohne Batterie
    labels = [
        "PV-Erzeugung", "Netzbezug",  # Quellen
        "Verbrauch",  # Senke
        "Direkter Eigenverbrauch", "Netzeinspeisung"  # Zwischenstationen
    ]
    
    # Indizes
    pv_gen_idx = 0
    grid_import_idx = 1
    consumption_idx = 2
    direct_self_cons_idx = 3
    grid_export_idx = 4
    
    # Werte aus KPIs oder berechnen
    if isinstance(kpis_no_battery, dict):
        grid_import = kpis_no_battery.get("total_grid_import_kwh", 0)
        grid_export = kpis_no_battery.get("total_grid_export_kwh", 0)
        direct_self_consumption = kpis_no_battery.get("total_direct_self_consumption_kwh", 0)
    else:
        # Fallback: Berechne aus Gesamtwerten
        grid_import = max(0, total_consumption - min(total_pv_generation, total_consumption))
        grid_export = max(0, total_pv_generation - min(total_pv_generation, total_consumption))
        direct_self_consumption = min(total_pv_generation, total_consumption)
    
    # Flüsse definieren
    source = [
        pv_gen_idx, pv_gen_idx,  # PV-Erzeugung -> 2 Ziele
        grid_import_idx,  # Netzbezug -> Verbrauch
        direct_self_cons_idx  # Direkter Eigenverbrauch -> Verbrauch
    ]
    target = [
        direct_self_cons_idx, grid_export_idx,  # Von PV
        consumption_idx,  # Von Netzbezug
        consumption_idx  # Von direktem Eigenverbrauch
    ]
    value = [
        direct_self_consumption,  # PV -> Direkter Eigenverbrauch
        grid_export,  # PV -> Netzeinspeisung
        grid_import,  # Netzbezug -> Verbrauch
        direct_self_consumption  # Direkter Eigenverbrauch -> Verbrauch
    ]
    
    # Filtere Null-Werte
    filtered_source = []
    filtered_target = []
    filtered_value = []
    for s, t, v in zip(source, target, value):
        if v > 0.01:
            filtered_source.append(s)
            filtered_target.append(t)
            filtered_value.append(v)
    
    # Farben für Knoten
    node_colors = [
        "#2E8B57",  # PV-Erzeugung (Dunkelgrün)
        "#4169E1",  # Netzbezug (Königsblau)
        "#DC143C",  # Verbrauch (Karmesinrot)
        "#32CD32",  # Direkter Eigenverbrauch (Limettengrün)
        "#FF8C00"   # Netzeinspeisung (Dunkelorange)
    ]
    
    # Farben für Flüsse
    link_colors = []
    for s, t in zip(filtered_source, filtered_target):
        if s == pv_gen_idx:
            if t == direct_self_cons_idx:
                link_colors.append("rgba(50, 205, 50, 0.6)")
            elif t == grid_export_idx:
                link_colors.append("rgba(255, 140, 0, 0.6)")
            else:
                link_colors.append("rgba(46, 139, 87, 0.4)")
        elif s == grid_import_idx:
            link_colors.append("rgba(65, 105, 225, 0.6)")
        elif s == direct_self_cons_idx:
            link_colors.append("rgba(50, 205, 50, 0.8)")
        else:
            link_colors.append("rgba(128, 128, 128, 0.4)")
    
    # Erstelle Diagramm
    fig = go.Figure(data=[go.Sankey(
        node=dict(
            label=labels,
            pad=25,
            thickness=30,
            line=dict(color="black", width=1.5),
            color=node_colors
        ),
        link=dict(
            source=filtered_source,
            target=filtered_target,
            value=filtered_value,
            color=link_colors,
            line=dict(width=0.8)
        ),
        arrangement="snap"
    )])
    
    # Layout-Einstellungen
    fig.update_layout(
        title_text="Jährliche Energieflüsse ohne Batteriespeicher",
        height=800,
        showlegend=True
    )
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    # Download-Button hinzufügen
    fig.update_layout(
        modebar_add=[
            "downloadImage"
        ]
    )
    
    return fig

def plot_cost_comparison_with_without_battery(
    annual_energy_cost_with_battery: float,
    total_consumption: float,
    total_pv_generation: float,
    battery_capacity_kwh: float,
    battery_cost_curve: dict,  # Dictionary mit Kostenkurven-Daten
    project_lifetime_years: int,
    discount_rate: float,
    price_grid_per_kwh,
    price_feed_in_per_kwh,
    annual_capacity_loss_percent: float = 1.0,  # Jährlicher Kapazitätsverlust in Prozent
    consumption_series: pd.Series = None,  # Für tatsächliche Simulation ohne Batterie
    pv_generation_series: pd.Series = None,  # Für tatsächliche Simulation ohne Batterie
    battery_efficiency_charge: float = 0.95,  # Für tatsächliche Simulation
    battery_efficiency_discharge: float = 0.95,  # Für tatsächliche Simulation
    battery_max_charge_kw: float = 10.0,  # Für tatsächliche Simulation
    battery_max_discharge_kw: float = 10.0,  # Für tatsächliche Simulation
    initial_soc_percent: float = 50.0,  # Für tatsächliche Simulation
    min_soc_percent: float = 10.0,  # Für tatsächliche Simulation
    max_soc_percent: float = 90.0,  # Für tatsächliche Simulation
    grid_export_with_battery: float = 0,  # Netzeinspeisung mit Batterie (aus Simulation)
    grid_import_with_battery: float = 0   # Netzbezug mit Batterie (aus Simulation)
):
    """
    Erstellt einen detaillierten Kostenvergleich mit und ohne Batteriespeicher.
    Zeigt die Amortisationszeit und den Break-even-Punkt.
    """
    # Spezialfall: 0 kWh Batteriekapazität (Referenzfall ohne Batterie)
    if battery_capacity_kwh == 0:
        investment_cost = 0
        annual_savings = 0
        payback_period = np.nan
        break_even_year = None
        irr_percentage = np.nan
    else:
        # Investitionskosten für Batteriespeicher - aus Kostenkurve
        from data_import import get_battery_cost
        investment_cost = get_battery_cost(battery_capacity_kwh, battery_cost_curve)
    
    # Jährliche Stromkosten ohne Speicher (Referenz) - VERWENDE TATSÄCHLICHE SIMULATION
    if consumption_series is not None and pv_generation_series is not None:
        # Führe eine Simulation OHNE Batterie durch für realistische Referenz
        from model import simulate_one_year
        no_battery_sim = simulate_one_year(
            consumption_series=consumption_series,
            pv_generation_series=pv_generation_series,
            battery_capacity_kwh=0.0,  # Keine Batterie
            battery_efficiency_charge=battery_efficiency_charge,
            battery_efficiency_discharge=battery_efficiency_discharge,
            battery_max_charge_kw=0.0,
            battery_max_discharge_kw=0.0,
            price_grid_per_kwh=price_grid_per_kwh,
            price_feed_in_per_kwh=price_feed_in_per_kwh,
            initial_soc_percent=0.0,
            min_soc_percent=0.0,
            max_soc_percent=0.0,
            annual_capacity_loss_percent=0.0,
            simulation_year=1
        )
        annual_cost_no_battery = no_battery_sim["kpis"]["annual_energy_cost"]
        grid_export_no_battery = no_battery_sim["kpis"]["total_grid_export_kwh"]
        grid_import_no_battery = no_battery_sim["kpis"]["total_grid_import_kwh"]
        
        # Berechne Eigenverbrauch ohne Batterie aus der Simulation
        self_consumption_no_battery = total_consumption - grid_import_no_battery
    else:
        # ERFORDERT: Zeitreihen müssen verfügbar sein für korrekte Berechnung
        raise ValueError("consumption_series und pv_generation_series sind erforderlich für korrekte Kostenberechnung. "
                       "Vereinfachte Annahmen führen zu falschen Ergebnissen.")
    
    # Einspeisevergütung ohne Speicher
    if isinstance(price_feed_in_per_kwh, pd.Series):
        feed_in_revenue_no_battery = grid_export_no_battery * price_feed_in_per_kwh.mean()
    else:
        feed_in_revenue_no_battery = grid_export_no_battery * price_feed_in_per_kwh
    
    # Stromkosten ohne Speicher (Netzbezug - Einspeisevergütung)
    if isinstance(price_grid_per_kwh, pd.Series):
        annual_cost_no_battery = (grid_import_no_battery * price_grid_per_kwh.mean()) - feed_in_revenue_no_battery
    else:
        annual_cost_no_battery = (grid_import_no_battery * price_grid_per_kwh) - feed_in_revenue_no_battery
    
    # Jährliche Ersparnis durch den Speicher (nur für Kapazitäten > 0)
    # Berechne Ersparnis nach der korrekten Formel:
    # Ersparnis = (verringerter Netzbezug × Strompreis) - (verringerte Einspeisung × Einspeisevergütung)
    if battery_capacity_kwh > 0:
        # Berechne verringerten Netzbezug und verringerte Einspeisung
        reduced_grid_import = grid_import_no_battery - grid_import_with_battery
        reduced_grid_export = grid_export_no_battery - grid_export_with_battery  # Ohne Batterie - mit Batterie (mehr Einspeisung ohne Batterie)
        
        # KORREKTE Ersparnis-Berechnung nach der Formel
        # Ersparnis = Weniger Netzbezug (Kosteneinsparung) - Weniger Einspeisung (Kostenverlust)
        # Da reduced_grid_export positiv ist (mehr Einspeisung ohne Batterie), ist das ein Kostenverlust
        if isinstance(price_grid_per_kwh, pd.Series):
            annual_savings = (reduced_grid_import * price_grid_per_kwh.mean()) - (reduced_grid_export * price_feed_in_per_kwh.mean())
            savings_from_reduced_import_comparison = reduced_grid_import * price_grid_per_kwh.mean()
            loss_from_reduced_export_comparison = reduced_grid_export * price_feed_in_per_kwh.mean()
        else:
            annual_savings = (reduced_grid_import * price_grid_per_kwh) - (reduced_grid_export * price_feed_in_per_kwh)
            savings_from_reduced_import_comparison = reduced_grid_import * price_grid_per_kwh
            loss_from_reduced_export_comparison = reduced_grid_export * price_feed_in_per_kwh
        
        # Amortisationszeit berechnen - KORRIGIERT: Degradation der Ersparnisse
        payback_period = np.nan
        if annual_savings > 0:
            # Berechne kumulierte Ersparnisse über die Zeit mit Degradation
            cumulative_savings = 0
            for year in range(1, project_lifetime_years + 1):
                # Degradation der Ersparnisse über die Jahre (nicht der Kapazität!)
                degradation_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (year - 1)
                year_savings = annual_savings * degradation_factor
                
                cumulative_savings += year_savings
                
                # Prüfe Break-even
                if cumulative_savings >= investment_cost:
                    # Berechne genaue Payback Period mit Bruchteilen
                    if year == 1:
                        # Im ersten Jahr: Einfache Division
                        payback_period = investment_cost / year_savings
                    else:
                        # In späteren Jahren: Interpolation zwischen den Jahren
                        previous_cumulative = cumulative_savings - year_savings
                        remaining_investment = investment_cost - previous_cumulative
                        fraction_of_year = remaining_investment / year_savings
                        payback_period = (year - 1) + fraction_of_year
                    break
            else:
                payback_period = 999  # Kein Break-even gefunden
        else:
            payback_period = 999  # Unrealistisch hoher Wert
    else:
        annual_savings = 0
        payback_period = np.nan
        savings_from_reduced_import_comparison = None
        loss_from_reduced_export_comparison = None
    
    # IRR und NPV-Berechnung für Batteriespeicher
    if battery_capacity_kwh > 0:
        # Cash Flows für IRR/NPV-Berechnung
        cash_flows = [-investment_cost]  # Initialinvestition (negativ)
        
        # Jährliche Cash Flows mit Degradation
        for year in range(1, project_lifetime_years + 1):
            # Degradation der Ersparnisse über die Jahre (nicht der Kapazität!)
            # Wenn die Batteriekapazität um 1% abnimmt, nehmen auch die Ersparnisse um ~1% ab
            degradation_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (year - 1)
            year_savings = annual_savings * degradation_factor
            
            cash_flows.append(year_savings)
        
        # IRR berechnen
        irr = calculate_irr(cash_flows)
        irr_percentage = irr * 100 if not np.isnan(irr) else np.nan
        
        # NPV mit Batterie berechnen
        npv_with_battery = 0
        for i, cash_flow in enumerate(cash_flows):
            npv_with_battery += cash_flow / ((1 + discount_rate) ** i)
    else:
        irr_percentage = np.nan
        npv_with_battery = 0
    
    # NPV ohne Batterie = 0 (Referenzszenario / Nullalternative)
    # Dies ist der Standard-Ansatz in Investitionsrechnungen:
    # Das Basisszenario "nichts tun" wird als Null gesetzt
    npv_without_battery = 0
    
    # Zeitreihe für den Vergleich erstellen
    years = np.arange(0, project_lifetime_years + 1)
    
    # Kosten ohne Batterie (linear - bleibt konstant)
    costs_without_battery = annual_cost_no_battery * years
    
    # Kosten mit Batterie (Investition + jährliche Kosten mit Degradation)
    costs_with_battery = np.zeros_like(years, dtype=float)
    costs_with_battery[0] = investment_cost  # Initialinvestition
    
    # Berechne kumulative Kosten mit Degradation
    cumulative_costs = investment_cost
    for year in range(1, project_lifetime_years + 1):
        # Degradation der Ersparnisse über die Jahre
        degradation_factor = (1.0 - annual_capacity_loss_percent / 100.0) ** (year - 1)
        degraded_annual_cost = annual_energy_cost_with_battery + (annual_savings * (1 - degradation_factor))
        cumulative_costs += degraded_annual_cost
        costs_with_battery[year] = cumulative_costs
    
    # Ersparnisse über die Zeit (mit Degradation)
    savings_over_time = costs_without_battery - costs_with_battery
    
    # Break-even-Punkt finden (nur für Kapazitäten > 0)
    # Finde den Schnittpunkt der Kostenlinien mit Degradation
    break_even_year = None
    exact_break_even = None
    if battery_capacity_kwh > 0:
        # Suche den ersten Punkt, wo Ersparnisse >= 0 werden
        for i in range(len(savings_over_time)):
            if savings_over_time[i] >= 0:
                if i == 0:
                    # Break-even bereits im ersten Jahr
                    exact_break_even = 0.0
                    break_even_year = 0
                else:
                    # Interpolation zwischen den Jahren
                    prev_savings = savings_over_time[i-1]
                    curr_savings = savings_over_time[i]
                    if curr_savings > prev_savings:  # Sicherstellen, dass es eine positive Steigung gibt
                        fraction = abs(prev_savings) / (curr_savings - prev_savings)
                        exact_break_even = (i - 1) + fraction
                        break_even_year = i
                break
    
    # Plot erstellen
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=('Kostenvergleich mit/ohne Batteriespeicher', 'Kumulierte Ersparnisse'),
        vertical_spacing=0.15,
        row_heights=[0.6, 0.4]
    )
    
    # Moderne Farbpalette verwenden
    colors = create_modern_color_palette(4)
    
    # Obere Grafik: Kostenvergleich
    fig.add_trace(go.Scatter(
        x=years, y=costs_without_battery,
        mode='lines+markers',
        name='Kosten ohne Batterie',
        line=dict(color=colors[0], width=3),
        marker=dict(size=6)
    ), row=1, col=1)
    
    fig.add_trace(go.Scatter(
        x=years, y=costs_with_battery,
        mode='lines+markers',
        name='Kosten mit Batterie',
        line=dict(color=colors[1], width=3),
        marker=dict(size=6)
    ), row=1, col=1)
    
    # Break-even-Punkt markieren - verwende den mathematischen Schnittpunkt der Kostenlinien
    if break_even_year is not None and exact_break_even is not None:
        # Verwende den berechneten Schnittpunkt der Kostenlinien (ohne Degradation)
        
        fig.add_vline(
            x=exact_break_even,
            line_dash="dash",
            line_color="green",
            line_width=2,
            row=1, col=1
        )
        
        # Interpoliere die Kosten für den genauen Break-even-Punkt
        if exact_break_even <= len(costs_with_battery) - 1:
            # Lineare Interpolation für die Kosten
            year_floor = int(exact_break_even)
            year_ceil = min(year_floor + 1, len(costs_with_battery) - 1)
            fraction = exact_break_even - year_floor
            
            if year_ceil < len(costs_with_battery):
                interpolated_cost = costs_with_battery[year_floor] + fraction * (costs_with_battery[year_ceil] - costs_with_battery[year_floor])
            else:
                interpolated_cost = costs_with_battery[year_floor]
        else:
            interpolated_cost = costs_with_battery[-1]
        
        fig.add_annotation(
            x=exact_break_even,
            y=interpolated_cost,
            text=f"Break-even:<br>{exact_break_even:.1f} Jahre<br><span style='font-size:10px'>(mit Degradation)</span>",
            showarrow=True,
            arrowhead=2,
            arrowcolor="green",
            arrowwidth=2,
            bgcolor="rgba(0,255,0,0.8)",
            bordercolor="green",
            borderwidth=1,
            row=1, col=1
        )
    
    # Untere Grafik: Kumulierte Ersparnisse
    fig.add_trace(go.Scatter(
        x=years, y=savings_over_time,
        mode='lines+markers',
        name='Kumulierte Ersparnisse',
        line=dict(color=colors[2], width=3),
        marker=dict(size=6),
        fill='tonexty',
        fillcolor=f'rgba({int(colors[2][1:3], 16)}, {int(colors[2][3:5], 16)}, {int(colors[2][5:7], 16)}, 0.3)'
    ), row=2, col=1)
    
    # Null-Linie für Ersparnisse
    fig.add_hline(y=0, line_dash="dash", line_color="gray", row=2, col=1)
    
    # Break-even-Punkt in Ersparnissen markieren - verwende den mathematischen Schnittpunkt
    if break_even_year is not None and exact_break_even is not None:
        
        fig.add_vline(
            x=exact_break_even,
            line_dash="dash",
            line_color="green",
            line_width=2,
            row=2, col=1
        )
        
        fig.add_annotation(
            x=exact_break_even,
            y=0,
            text=f"Break-even:<br>{exact_break_even:.1f} Jahre<br><span style='font-size:10px'>(mit Degradation)</span>",
            showarrow=True,
            arrowhead=2,
            arrowcolor="green",
            arrowwidth=2,
            bgcolor="rgba(0,255,0,0.8)",
            bordercolor="green",
            borderwidth=1,
            row=2, col=1
        )
    
    # Layout verbessern
    # Modernes Layout
    fig.update_layout(
        title_text='Kostenvergleich: Mit vs. Ohne Batteriespeicher',
        height=800,
        showlegend=True
    )
    
    # Achsenbeschriftungen
    fig.update_xaxes(title_text='Jahre', row=2, col=1)
    fig.update_yaxes(title_text='Kosten (€)', row=1, col=1)
    fig.update_yaxes(title_text='Ersparnisse (€)', row=2, col=1)
    
    # Modernes Theme anwenden
    fig = apply_modern_plotly_theme(fig)
    
    # Detaillierte Berechnungsübersicht zurückgeben
    comparison_details = {
        'investment_cost': investment_cost,
        'annual_cost_no_battery': annual_cost_no_battery,
        'annual_cost_with_battery': annual_energy_cost_with_battery,
        'annual_savings': annual_savings,
        'savings_from_reduced_import': savings_from_reduced_import_comparison if battery_capacity_kwh > 0 else None,
        'loss_from_reduced_export': loss_from_reduced_export_comparison if battery_capacity_kwh > 0 else None,
        'payback_period_years': payback_period,
        'break_even_year': break_even_year,
        'total_savings_over_lifetime': savings_over_time[-1],
        'roi_percentage': (savings_over_time[-1] / investment_cost - 1) * 100 if investment_cost > 0 else 0,
        'irr_percentage': irr_percentage,
        'npv_without_battery': npv_without_battery,  # NPV ohne Batterie
        'npv_with_battery': npv_with_battery,  # NPV mit Batterie
        # Einspeisevergütung hinzugefügt
        'feed_in_revenue_without_battery': feed_in_revenue_no_battery,
        'feed_in_revenue_with_battery': grid_export_with_battery * (price_feed_in_per_kwh.mean() if isinstance(price_feed_in_per_kwh, pd.Series) else price_feed_in_per_kwh),
        'grid_import_no_battery': grid_import_no_battery,
        'grid_export_no_battery': grid_export_no_battery,
        'self_consumption_no_battery': self_consumption_no_battery
    }
    
    # Erstelle ein Dictionary mit Figure und comparison_details
    result = {
        'figure': fig,
        'comparison_details': comparison_details
    }
    
    return result 

def plot_data_control(consumption_series, pv_generation_series, start_date=None, end_date=None):
    """
    Erstellt eine einfache Datenkontrolle-Grafik mit PV-Ertrag und Verbrauch.
    
    Args:
        consumption_series (pd.Series): Verbrauchsdaten mit 15-Minuten-Intervallen
        pv_generation_series (pd.Series): PV-Erzeugungsdaten mit 15-Minuten-Intervallen
        start_date (str, optional): Startdatum im Format 'YYYY-MM-DD'
        end_date (str, optional): Enddatum im Format 'YYYY-MM-DD'
    
    Returns:
        plotly.graph_objects.Figure: Interaktive Grafik
    """
    try:
        # Erstelle DataFrame für die Grafik
        df_control = pd.DataFrame({
            'Verbrauch_kWh': consumption_series,
            'PV_Erzeugung_kWh': pv_generation_series
        })
        
        # Filtere nach Datum falls angegeben
        if start_date and end_date:
            df_control = df_control.loc[start_date:end_date]
        elif start_date:
            df_control = df_control.loc[start_date:]
        elif end_date:
            df_control = df_control.loc[:end_date]
        
        # Erstelle die Grafik
        fig = go.Figure()
        
        # PV-Erzeugung (grün)
        fig.add_trace(go.Scatter(
            x=df_control.index,
            y=df_control['PV_Erzeugung_kWh'],
            mode='lines',
            name='PV-Erzeugung',
            line=dict(color='green', width=2),
            fill='tozeroy',
            fillcolor='rgba(0,255,0,0.2)'
        ))
        
        # Verbrauch (rot)
        fig.add_trace(go.Scatter(
            x=df_control.index,
            y=df_control['Verbrauch_kWh'],
            mode='lines',
            name='Verbrauch',
            line=dict(color='red', width=2),
            fill='tozeroy',
            fillcolor='rgba(255,0,0,0.2)'
        ))
        
        # Modernes Layout
        fig.update_layout(
            title_text='Datenkontrolle - PV-Erzeugung vs. Verbrauch',
            xaxis_title='Zeit',
            yaxis_title='Leistung (kWh)',
            height=600,
            showlegend=True,
            hovermode='x unified'
        )
        
        # Modernes Theme anwenden
        fig = apply_modern_plotly_theme(fig)
        
        # Y-Achse anpassen - feine Skalierung
        max_value = max(df_control['PV_Erzeugung_kWh'].max(), df_control['Verbrauch_kWh'].max())
        fig.update_yaxes(
            gridcolor='lightgray',
            zeroline=True,
            zerolinecolor='black',
            zerolinewidth=1,
            dtick=max(0.5, max_value/20)  # Feine Skalierung
        )
        
        # X-Achse anpassen
        fig.update_xaxes(
            gridcolor='lightgray',
            showgrid=True
        )
        
        return fig
        
    except Exception as e:
        print(f"Fehler bei der Erstellung der Datenkontrolle-Grafik: {e}")
        # Fallback: Leere Grafik
        fig = go.Figure()
        fig.add_annotation(
            text="Fehler beim Laden der Datenkontrolle-Grafik",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="red")
        )
        return fig 