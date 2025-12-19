# Anleitung zur Entwicklung eines Python-Programms zur optimalen Batteriespeichergröße (Version 2.0)

## 1. Ziel des Programms

Das Ziel dieses Projekts ist die Entwicklung eines robusten und benutzerfreundlichen Python-Tools, das die optimale Größe und die wirtschaftliche Rentabilität eines Batteriespeichers für verschiedene Anwendungsfälle analysiert. Der Fokus liegt dabei auf der Simulation und Visualisierung der Auswirkungen unterschiedlicher Parameter auf wichtige Kennzahlen. Das Tool soll es Nutzern ermöglichen, interaktiv verschiedene Szenarien zu simulieren, indem sie Parameter wie die Größe der Photovoltaik (PV)-Anlage, Strompreise, Speicherkosten und Verbrauchscharakteristiken anpassen. Die Ergebnisse dieser Simulationen sollen in verständlichen und aussagekräftigen Grafiken dargestellt werden, um Einblicke in den Autarkiegrad, die Eigenverbrauchsquote, die Amortisationszeit und die Netto-Ersparnis zu gewinnen. Dies ermöglicht eine fundierte Entscheidungsfindung bei der Planung und Dimensionierung von Batteriespeichersystemen für Einfamilienhäuser, Mehrfamilienhäuser oder ganze Quartiere.

## 2. Projektstruktur (Dateien und Module)

Eine klare und modulare Projektstruktur ist entscheidend für die Wartbarkeit, Erweiterbarkeit und Verständlichkeit des Codes. Die vorgeschlagene Struktur gliedert das Projekt in logische Einheiten, die jeweils spezifische Aufgaben erfüllen. Dies fördert die Wiederverwendbarkeit von Code und erleichtert die Zusammenarbeit in Teams.

*   `main.py`: Dies ist das Hauptskript, das den Einstiegspunkt für die Anwendung darstellt. Es orchestriert den Aufruf der verschiedenen Module, initialisiert die Benutzeroberfläche (falls vorhanden) und startet den Simulations- und Analyseprozess. Von hier aus werden die Daten geladen, die Simulationen ausgeführt und die Ergebnisse visualisiert.

*   `data_import.py`: Dieses Modul ist verantwortlich für das Laden und die initiale Vorverarbeitung aller externen Daten, die für die Simulation benötigt werden. Dazu gehören historische Verbrauchsdaten, PV-Erzeugungsprofile und möglicherweise auch Strompreiszeitreihen. Eine robuste Implementierung dieses Moduls stellt sicher, dass die Daten korrekt formatiert und für die weitere Verarbeitung bereitgestellt werden.

*   `model.py`: Das Kernstück der Simulationslogik. Dieses Modul enthält die Algorithmen, die die Energieflüsse innerhalb des Systems (PV-Erzeugung, Verbrauch, Batteriespeicher, Netzbezug/-einspeisung) über einen bestimmten Zeitraum (z.B. ein Jahr) simulieren. Es berechnet stündlich den Ladezustand der Batterie und die resultierenden Energieflüsse zum und vom Netz.

*   `scenarios.py`: Dieses Modul definiert und verwaltet verschiedene Simulationsszenarien. Es ermöglicht die systematische Variation von Parametern, um unterschiedliche 


Fragestellungen zu untersuchen, wie z.B. die optimale Speichergröße, der Einfluss variabler Stromtarife oder der Vergleich verschiedener Haushaltstypen. Dieses Modul wird die `model.py` Funktionen mit unterschiedlichen Eingabeparametern aufrufen und die Ergebnisse sammeln.

*   `analysis.py`: Dieses Modul ist für die Berechnung von Key Performance Indicators (KPIs) und die Erstellung von Visualisierungen zuständig. Nach der Simulation durch `model.py` werden hier die Rohdaten verarbeitet, um Kennzahlen wie Autarkiegrad, Eigenverbrauchsquote, Amortisationszeit und Netto-Ersparnis zu ermitteln. Darüber hinaus werden hier die Funktionen zur Generierung der verschiedenen Diagramme und Grafiken implementiert, die die Simulationsergebnisse anschaulich darstellen.

*   `ui.py`: (Optional, aber dringend empfohlen für Interaktivität) Dieses Modul wird die grafische Benutzeroberfläche (GUI) des Tools implementieren. Für eine schnelle Entwicklung und gute Interaktivität bietet sich die Verwendung von Bibliotheken wie Streamlit an. Die GUI ermöglicht es dem Nutzer, Parameter einfach anzupassen, Simulationen zu starten und die Ergebnisse direkt im Browser zu visualisieren.

*   `config.py`: **(Neu)** Eine zentrale Konfigurationsdatei, die Standardwerte und Konstanten für das gesamte Projekt enthält. Dies umfasst beispielsweise Standard-Strompreise, Kostenannahmen für Batteriespeicher und PV-Anlagen, Wirkungsgrade oder andere globale Einstellungen. Die Auslagerung dieser Werte in eine separate Datei erleichtert die Verwaltung und Anpassung von Parametern, ohne den Kerncode ändern zu müssen. Dies ist besonders nützlich, um verschiedene Basisszenarien schnell zu definieren und zu wechseln.

*   `README.md`: **(Neu)** Eine umfassende Dokumentationsdatei im Markdown-Format, die eine Einführung in das Projekt gibt. Sie sollte Anweisungen zur Installation der benötigten Abhängigkeiten, zur Ausführung des Programms, eine kurze Beschreibung der Funktionalitäten und Beispiele für die Nutzung enthalten. Eine gute `README.md` ist essenziell für die Nutzbarkeit und die Zusammenarbeit an dem Projekt.

## 3. Datenimport und Vorverarbeitung (`data_import.py`)

Das `data_import.py`-Modul ist der erste Schritt in der Datenpipeline. Es ist entscheidend, dass die hier geladenen und vorverarbeiteten Daten in einem konsistenten und nutzbaren Format für die nachfolgenden Simulationsschritte bereitgestellt werden. Die Qualität und das Format der Eingangsdaten haben einen direkten Einfluss auf die Genauigkeit und Zuverlässigkeit der Simulationsergebnisse.

### 3.1. Funktion `load_data(file_path)`

Diese Funktion ist für das Einlesen der Rohdaten zuständig. Typischerweise werden hier Excel-Dateien verwendet, die Zeitreihen für Stromverbrauch und PV-Erzeugung sowie möglicherweise weitere Parameter enthalten. Die Funktion sollte robust genug sein, um verschiedene Spaltennamen oder Dateistrukturen zu berücksichtigen oder zumindest eine klare Erwartung an das Eingangsformat zu kommunizieren.

```python
import pandas as pd

def load_data(file_path):
    """
    Lädt die Daten aus einer Excel-Datei.
    Erwartet zwei Sheets: 'Stromspeicher_Parameter' und 'Verbrauchsdaten'.
    """
    try:
        # Laden der Parameter für den Stromspeicher (z.B. aus einem separaten Sheet)
        df_stromspeicher = pd.read_excel(file_path, sheet_name='Stromspeicher_Parameter')
        # Laden der Verbrauchs- und PV-Erzeugungsdaten (Zeitreihen)
        df_verbrauch = pd.read_excel(file_path, sheet_name='Verbrauchsdaten')
        return df_stromspeicher, df_verbrauch
    except FileNotFoundError:
        print(f"Fehler: Datei nicht gefunden unter {file_path}")
        return None, None
    except Exception as e:
        print(f"Fehler beim Laden der Daten: {e}")
        return None, None

```

### 3.2. Funktion `preprocess_data(df_stromspeicher, df_verbrauch)`

Nach dem Laden der Rohdaten müssen diese für die Simulation aufbereitet werden. Dies beinhaltet das Extrahieren relevanter Parameter und die Sicherstellung der Konsistenz der Zeitreihen. Eine kritische Anforderung ist, dass die Zeitreihen für Verbrauch und PV-Erzeugung einen lückenlosen, stündlichen Zeitindex für ein ganzes Jahr (8760 Datenpunkte) aufweisen. Fehlende Werte oder Inkonsistenzen müssen hier intelligent behandelt werden.

```python
import pandas as pd

def preprocess_data(df_stromspeicher, df_verbrauch):
    """
    Verarbeitet die geladenen Rohdaten und bereitet sie für die Simulation vor.
    Extrahiert Parameter und stellt die Konsistenz der Zeitreihen sicher.
    """
    # Beispiel: Extrahieren von Parametern aus df_stromspeicher
    # Annahme: df_stromspeicher hat Spalten wie 'Parameter' und 'Wert'
    parameters = df_stromspeicher.set_index('Parameter')['Wert'].to_dict()

    # Zeitreihen-Verarbeitung
    # Annahme: df_verbrauch hat Spalten wie 'Zeitstempel', 'Verbrauch_kWh', 'PV_Erzeugung_kWh'
    df_verbrauch['Zeitstempel'] = pd.to_datetime(df_verbrauch['Zeitstempel'])
    df_verbrauch = df_verbrauch.set_index('Zeitstempel')

    # Sicherstellen eines lückenlosen stündlichen Zeitindex für ein ganzes Jahr
    # Erstellen eines vollständigen Zeitindex für 8760 Stunden (ein Jahr)
    start_date = df_verbrauch.index.min().replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date + pd.Timedelta(days=365) - pd.Timedelta(hours=1) # Für 8760 Stunden
    full_time_index = pd.date_range(start=start_date, end=end_date, freq='H')

    df_processed = pd.DataFrame(index=full_time_index)
    df_processed = df_processed.merge(df_verbrauch[['Verbrauch_kWh', 'PV_Erzeugung_kWh']], 
                                     left_index=True, right_index=True, how='left')

    # Fehlende Werte füllen (z.B. mit 0 oder durch Interpolation)
    # Hier: Annahme, dass fehlende Werte 0 sind, falls keine Daten vorhanden sind
    df_processed['Verbrauch_kWh'] = df_processed['Verbrauch_kWh'].fillna(0) 
    df_processed['PV_Erzeugung_kWh'] = df_processed['PV_Erzeugung_kWh'].fillna(0)

    # Flexibilität: PV-Erzeugungsprofil kann auch aus Verbrauchs- und Netzeinspeisungsdaten rekonstruiert werden
    # Beispiel: Wenn nur Verbrauch und Netzbezug/-einspeisung gegeben sind
    # if 'Netzbezug_kWh' in df_verbrauch.columns and 'Netzeinspeisung_kWh' in df_verbrauch.columns:
    #    df_processed['PV_Erzeugung_kWh'] = df_processed['Verbrauch_kWh'] + df_verbrauch['Netzeinspeisung_kWh'] - df_verbrauch['Netzbezug_kWh']
    #    df_processed['PV_Erzeugung_kWh'] = df_processed['PV_Erzeugung_kWh'].clip(lower=0) # PV-Erzeugung kann nicht negativ sein

    return parameters, df_processed['Verbrauch_kWh'], df_processed['PV_Erzeugung_kWh']

```

## 4. Kernlogik der Simulation (`model.py`)

Das `model.py`-Modul ist das Herzstück des Analyse-Tools. Es implementiert die stündliche Simulationslogik, die die Energieflüsse zwischen PV-Anlage, Verbrauch, Batteriespeicher und dem öffentlichen Netz berechnet. Die Simulation muss flexibel genug sein, um verschiedene Parameter zu berücksichtigen und detaillierte Zeitreihen der Energieflüsse als Ergebnis zu liefern, die für die spätere Analyse und Visualisierung unerlässlich sind.

### 4.1. Funktion `simulate_one_year(...)`

Diese Funktion führt die eigentliche Simulation für ein ganzes Jahr (8760 Stunden) durch. Sie nimmt alle relevanten Parameter als Argumente entgegen, um eine hohe Flexibilität bei der Szenarienanalyse zu gewährleisten. Die Funktion sollte nicht nur die finalen KPIs zurückgeben, sondern auch die vollständigen Zeitreihen der Simulation, wie den Ladezustand der Batterie (SoC), den Stromfluss von/zur Batterie, den Netzbezug und die Netzeinspeisung.

```python
import numpy as np
import pandas as pd

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
    initial_soc_percent: float = 50.0 # Anfangsladezustand der Batterie in %
) -> dict:
    """
    Simuliert die Energieflüsse für ein Jahr auf Stundenbasis.

    Args:
        consumption_series (pd.Series): Stündlicher Stromverbrauch in kWh.
        pv_generation_series (pd.Series): Stündliche PV-Erzeugung in kWh.
        battery_capacity_kwh (float): Speicherkapazität der Batterie in kWh.
        battery_efficiency_charge (float): Wirkungsgrad der Batterie beim Laden (0-1).
        battery_efficiency_discharge (float): Wirkungsgrad der Batterie beim Entladen (0-1).
        battery_max_charge_kw (float): Maximale Ladeleistung der Batterie in kW.
        battery_max_discharge_kw (float): Maximale Entladeleistung der Batterie in kW.
        price_grid_per_kwh (float | pd.Series): Preis für Netzbezug in Euro/kWh. Kann eine Konstante oder eine Zeitreihe sein.
        price_feed_in_per_kwh (float | pd.Series): Preis für Netzeinspeisung in Euro/kWh. Kann eine Konstante oder eine Zeitreihe sein.
        initial_soc_percent (float): Anfangsladezustand der Batterie in % (0-100).

    Returns:
        dict: Ein Dictionary mit Zeitreihen der Energieflüsse und initialen KPIs.
    """

    num_hours = len(consumption_series)

    # Initialisierung der Zeitreihen für die Ergebnisse
    soc_series = np.zeros(num_hours) # State of Charge
    grid_import_series = np.zeros(num_hours) # Netzbezug
    grid_export_series = np.zeros(num_hours) # Netzeinspeisung
    battery_charge_series = np.zeros(num_hours) # Ladung der Batterie
    battery_discharge_series = np.zeros(num_hours) # Entladung der Batterie
    direct_self_consumption_series = np.zeros(num_hours) # Direkter Eigenverbrauch

    # Aktueller Ladezustand der Batterie (in kWh)
    current_soc_kwh = (initial_soc_percent / 100.0) * battery_capacity_kwh

    for i in range(num_hours):
        pv_gen = pv_generation_series.iloc[i]
        consumption = consumption_series.iloc[i]

        # 1. Direkter Eigenverbrauch (PV direkt an Last)
        direct_self_consumption = min(pv_gen, consumption)
        remaining_pv = pv_gen - direct_self_consumption
        remaining_consumption = consumption - direct_self_consumption

        # 2. Überschüssige PV-Energie in Batterie laden
        if remaining_pv > 0:
            # Maximal mögliche Ladung basierend auf Kapazität und max. Ladeleistung
            max_charge_possible_kwh = min(battery_capacity_kwh - current_soc_kwh, battery_max_charge_kw)
            charge_from_pv = min(remaining_pv, max_charge_possible_kwh / battery_efficiency_charge) # Berücksichtige Wirkungsgrad
            
            current_soc_kwh += charge_from_pv * battery_efficiency_charge
            battery_charge_series[i] = charge_from_pv
            remaining_pv -= charge_from_pv

        # 3. Fehlende Energie aus Batterie entladen oder aus Netz beziehen
        if remaining_consumption > 0:
            # Maximal mögliche Entladung basierend auf Kapazität und max. Entladeleistung
            max_discharge_possible_kwh = min(current_soc_kwh, battery_max_discharge_kw)
            discharge_to_consumption = min(remaining_consumption, max_discharge_possible_kwh * battery_efficiency_discharge) # Berücksichtige Wirkungsgrad
            
            current_soc_kwh -= discharge_to_consumption / battery_efficiency_discharge
            battery_discharge_series[i] = discharge_to_consumption
            remaining_consumption -= discharge_to_consumption

        # 4. Restlicher PV-Überschuss ins Netz einspeisen
        if remaining_pv > 0:
            grid_export_series[i] = remaining_pv

        # 5. Restlicher Verbrauch aus dem Netz beziehen
        if remaining_consumption > 0:
            grid_import_series[i] = remaining_consumption

        # Sicherstellen, dass SoC innerhalb der Grenzen bleibt (kleine Rundungsfehler)
        current_soc_kwh = np.clip(current_soc_kwh, 0, battery_capacity_kwh)
        soc_series[i] = current_soc_kwh

    # Erstellung eines Pandas DataFrame für die Zeitreihen-Ergebnisse
    time_index = consumption_series.index
    results_df = pd.DataFrame({
        'PV_Generation_kWh': pv_generation_series.values,
        'Consumption_kWh': consumption_series.values,
        'SOC_kWh': soc_series,
        'Battery_Charge_kWh': battery_charge_series,
        'Battery_Discharge_kWh': battery_discharge_series,
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

    # Jährliche Stromkosten ohne Speicher (Referenz)
    # Annahme: Ohne PV und Speicher wird alles aus dem Netz bezogen
    # Oder: PV-Erzeugung wird komplett eingespeist und Verbrauch komplett bezogen
    # Hier: Einfachheitshalber nur die Kosten des Netzbezugs mit Speicher
    annual_energy_cost = total_import_cost - total_export_revenue

    return {
        'time_series_data': results_df,
        'kpis': {
            'autarky_rate': autarky_rate,
            'self_consumption_rate': self_consumption_rate,
            'total_grid_import_kwh': total_grid_import,
            'total_grid_export_kwh': total_grid_export,
            'annual_energy_cost': annual_energy_cost,
            'total_pv_generation_kwh': total_pv_generation,
            'total_consumption_kwh': total_consumption
        }
    }

```

### 4.2. Erweiterung für variable Stromtarife (Day-Ahead)

Um den Einfluss variabler Stromtarife zu simulieren, muss die Logik in `simulate_one_year` angepasst werden. Die Entscheidung, wann die Batterie geladen oder entladen wird, hängt dann nicht mehr nur vom PV-Überschuss oder Verbrauch ab, sondern auch von den aktuellen und prognostizierten Strompreisen. Dies erfordert eine vorausschauende Steuerung der Batterie.

**Konzept der vorausschauenden Steuerung (Heuristik):**

Eine vollständige Optimierung der Batteriesteuerung unter variablen Preisen ist komplex und erfordert Optimierungsalgorithmen (z.B. Lineare Programmierung). Für eine erste Implementierung kann eine heuristische Steuerung verwendet werden:

*   **Laden aus dem Netz**: Wenn der Strompreis sehr niedrig ist (z.B. unter einem definierten Schwellenwert oder unter dem Einspeisepreis), kann die Batterie aus dem Netz geladen werden, um später teuren Netzbezug zu vermeiden oder teuer einzuspeisen.
*   **Entladen ins Netz**: Wenn der Strompreis sehr hoch ist (z.B. über einem definierten Schwellenwert oder über dem Bezugspreis), kann die Batterie entladen werden, um den Strom teuer zu verkaufen.
*   **Priorität**: PV-Eigenverbrauch hat immer höchste Priorität. Danach kommt das Laden/Entladen der Batterie zur Deckung des Eigenverbrauchs. Erst danach wird das Laden/Entladen zur Arbitrage (Kauf bei niedrigem Preis, Verkauf bei hohem Preis) in Betracht gezogen.

**Anpassung der `simulate_one_year` Funktion (Pseudocode-Ergänzung):**

```python
# ... innerhalb der Schleife for i in range(num_hours):

# ... (Bestehende Logik für direkten Eigenverbrauch und PV-Laden/Entladen)

# Erweiterung für variable Tarife und Arbitrage-Möglichkeiten
current_grid_price = price_grid_per_kwh.iloc[i] if isinstance(price_grid_per_kwh, pd.Series) else price_grid_per_kwh
current_feed_in_price = price_feed_in_per_kwh.iloc[i] if isinstance(price_feed_in_per_kwh, pd.Series) else price_feed_in_per_kwh

# Szenario 1: Batterie aus Netz laden, wenn Preis niedrig ist UND Batterie nicht voll ist
if current_grid_price < config.THRESHOLD_BUY_PRICE and current_soc_kwh < battery_capacity_kwh:
    charge_from_grid = min(config.MAX_GRID_CHARGE_KW, battery_capacity_kwh - current_soc_kwh) # Max. Ladung aus Netz
    current_soc_kwh += charge_from_grid * battery_efficiency_charge
    battery_charge_series[i] += charge_from_grid # Addiere zu bestehender Ladung
    grid_import_series[i] += charge_from_grid # Zähle als Netzbezug

# Szenario 2: Batterie ins Netz entladen, wenn Preis hoch ist UND Batterie geladen ist
if current_feed_in_price > config.THRESHOLD_SELL_PRICE and current_soc_kwh > 0:
    discharge_to_grid = min(config.MAX_GRID_DISCHARGE_KW, current_soc_kwh) # Max. Entladung ins Netz
    current_soc_kwh -= discharge_to_grid / battery_efficiency_discharge
    battery_discharge_series[i] += discharge_to_grid # Addiere zu bestehender Entladung
    grid_export_series[i] += discharge_to_grid # Zähle als Netzeinspeisung

# ... (Bestehende Logik für restlichen PV-Überschuss und restlichen Verbrauch)

```

**Hinweis**: Die Schwellenwerte `config.THRESHOLD_BUY_PRICE` und `config.THRESHOLD_SELL_PRICE` sowie `config.MAX_GRID_CHARGE_KW` und `config.MAX_GRID_DISCHARGE_KW` sollten in der `config.py` definiert werden und als Parameter an die Funktion übergeben werden können, um sie flexibel zu gestalten.

## 5. Szenarien und Parametervariation (`scenarios.py`)

Das `scenarios.py`-Modul ist dafür verantwortlich, verschiedene Simulationsläufe zu orchestrieren, indem es die Parameter der `simulate_one_year`-Funktion systematisch variiert. Dies ermöglicht die Beantwortung komplexer Fragen und die Durchführung von Sensitivitätsanalysen. Die Ergebnisse dieser Szenarien werden gesammelt und an das `analysis.py`-Modul zur weiteren Verarbeitung übergeben.

### 5.1. Szenario 1: Wirtschaftlich optimale Speichergröße

Dieses Szenario zielt darauf ab, die Batteriespeichergröße zu finden, die unter gegebenen Bedingungen die beste Wirtschaftlichkeit aufweist. Dies kann die kürzeste Amortisationszeit, den höchsten Net Present Value (NPV) oder die höchste Netto-Ersparnis über einen definierten Zeitraum bedeuten.

```python
import pandas as pd
from model import simulate_one_year
from analysis import calculate_financial_kpis # Annahme: Funktion existiert in analysis.py

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
    battery_cost_per_kwh: float, # Kosten pro kWh Speicherkapazität
    installation_cost_fixed: float, # Feste Installationskosten
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05 # Diskontierungsrate für NPV
) -> list:
    """
    Findet die wirtschaftlich optimale Speichergröße durch Iteration über verschiedene Kapazitäten.
    """
    results = []
    for capacity in np.arange(min_capacity_kwh, max_capacity_kwh + step_kwh, step_kwh):
        if capacity == 0: # Szenario ohne Speicher
            # Simulieren ohne Batterie, oder mit sehr kleiner Batterie, die nicht aktiv ist
            sim_result = simulate_one_year(
                consumption_series=consumption_series,
                pv_generation_series=pv_generation_series,
                battery_capacity_kwh=0.001, # Sehr kleine Kapazität, um Logik nicht zu stören
                battery_efficiency_charge=1.0, # Irrelevant
                battery_efficiency_discharge=1.0, # Irrelevant
                battery_max_charge_kw=0.0, # Keine Ladung
                battery_max_discharge_kw=0.0, # Keine Entladung
                price_grid_per_kwh=price_grid_per_kwh,
                price_feed_in_per_kwh=price_feed_in_per_kwh,
                initial_soc_percent=0.0
            )
        else:
            sim_result = simulate_one_year(
                consumption_series=consumption_series,
                pv_generation_series=pv_generation_series,
                battery_capacity_kwh=capacity,
                battery_efficiency_charge=battery_efficiency_charge,
                battery_efficiency_discharge=battery_efficiency_discharge,
                battery_max_charge_kw=battery_max_charge_kw,
                battery_max_discharge_kw=battery_max_discharge_kw,
                price_grid_per_kwh=price_grid_per_kwh,
                price_feed_in_per_kwh=price_feed_in_per_kwh
            )
        
        # Finanzielle KPIs berechnen
        financial_kpis = calculate_financial_kpis(
            sim_result['kpis']['annual_energy_cost'], # Jährliche Kosten mit Speicher
            total_consumption=sim_result['kpis']['total_consumption_kwh'],
            total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
            battery_capacity_kwh=capacity,
            battery_cost_per_kwh=battery_cost_per_kwh,
            installation_cost_fixed=installation_cost_fixed,
            project_lifetime_years=project_lifetime_years,
            discount_rate=discount_rate,
            price_grid_per_kwh=price_grid_per_kwh, # Für Referenzkosten ohne Speicher
            price_feed_in_per_kwh=price_feed_in_per_kwh # Für Referenzkosten ohne Speicher
        )

        results.append({
            'battery_capacity_kwh': capacity,
            'autarky_rate': sim_result['kpis']['autarky_rate'],
            'self_consumption_rate': sim_result['kpis']['self_consumption_rate'],
            'annual_energy_cost': sim_result['kpis']['annual_energy_cost'],
            **financial_kpis # Füge finanzielle KPIs hinzu
        })
    return results

```

### 5.2. Szenario 2: Einfluss variabler Stromtarife (Day-Ahead)

Dieses Szenario untersucht, wie sich die optimale Speicherstrategie und die Wirtschaftlichkeit ändern, wenn Strom stündlich unterschiedliche Preise hat. Dies erfordert, dass die `simulate_one_year`-Funktion in `model.py` die variablen Preise korrekt verarbeitet und die Batteriesteuerung entsprechend anpasst (siehe Abschnitt 4.2).

```python
import pandas as pd
from model import simulate_one_year
from analysis import calculate_financial_kpis

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
    battery_cost_per_kwh: float,
    installation_cost_fixed: float,
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05
) -> dict:
    """
    Führt eine Simulation mit variablen Stromtarifen durch.
    """
    sim_result = simulate_one_year(
        consumption_series=consumption_series,
        pv_generation_series=pv_generation_series,
        battery_capacity_kwh=battery_capacity_kwh,
        battery_efficiency_charge=battery_efficiency_charge,
        battery_efficiency_discharge=battery_efficiency_discharge,
        battery_max_charge_kw=battery_max_charge_kw,
        battery_max_discharge_kw=battery_max_discharge_kw,
        price_grid_per_kwh=price_grid_series, # Übergabe der Zeitreihe
        price_feed_in_per_kwh=price_feed_in_series # Übergabe der Zeitreihe
    )

    financial_kpis = calculate_financial_kpis(
        sim_result['kpis']['annual_energy_cost'],
        total_consumption=sim_result['kpis']['total_consumption_kwh'],
        total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
        battery_capacity_kwh=battery_capacity_kwh,
        battery_cost_per_kwh=battery_cost_per_kwh,
        installation_cost_fixed=installation_cost_fixed,
        project_lifetime_years=project_lifetime_years,
        discount_rate=discount_rate,
        price_grid_per_kwh=price_grid_series, # Für Referenzkosten ohne Speicher
        price_feed_in_per_kwh=price_feed_in_series # Für Referenzkosten ohne Speicher
    )

    return {
        'scenario_name': 'Variable Tarife',
        'time_series_data': sim_result['time_series_data'],
        'kpis': {**sim_result['kpis'], **financial_kpis}
    }

```

### 5.3. Szenario 3: Vergleich von Haushalts-Typen (EFH, MFH, Quartier)

Dieses Szenario ermöglicht den direkten Vergleich der Performance des Batteriespeichersystems für verschiedene Last- und PV-Profile, die typisch für Einfamilienhäuser (EFH), Mehrfamilienhäuser (MFH) oder ganze Quartiere sind. Dies erfordert, dass entsprechende Last- und PV-Profile in `data_import.py` geladen oder generiert werden können.

```python
import pandas as pd
from model import simulate_one_year
from analysis import calculate_financial_kpis

def compare_household_scenarios(
    household_profiles: dict, # Dictionary mit {Name: {'consumption': Series, 'pv_gen': Series, ...}}
    battery_capacity_kwh: float,
    battery_efficiency_charge: float,
    battery_efficiency_discharge: float,
    battery_max_charge_kw: float,
    battery_max_discharge_kw: float,
    price_grid_per_kwh,
    price_feed_in_per_kwh,
    battery_cost_per_kwh: float,
    installation_cost_fixed: float,
    project_lifetime_years: int = 20,
    discount_rate: float = 0.05
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
            price_feed_in_per_kwh=price_feed_in_per_kwh
        )

        financial_kpis = calculate_financial_kpis(
            sim_result['kpis']['annual_energy_cost'],
            total_consumption=sim_result['kpis']['total_consumption_kwh'],
            total_pv_generation=sim_result['kpis']['total_pv_generation_kwh'],
            battery_capacity_kwh=battery_capacity_kwh,
            battery_cost_per_kwh=battery_cost_per_kwh,
            installation_cost_fixed=installation_cost_fixed,
            project_lifetime_years=project_lifetime_years,
            discount_rate=discount_rate,
            price_grid_per_kwh=price_grid_per_kwh, # Für Referenzkosten ohne Speicher
            price_feed_in_per_kwh=price_feed_in_per_kwh # Für Referenzkosten ohne Speicher
        )

        all_scenario_results[name] = {
            'time_series_data': sim_result['time_series_data'],
            'kpis': {**sim_result['kpis'], **financial_kpis}
        }
    return all_scenario_results

```

## 6. Auswertung und Visualisierung (`analysis.py`)

Das `analysis.py`-Modul ist von zentraler Bedeutung, um die komplexen Simulationsergebnisse in verständliche und aussagekräftige Grafiken und Kennzahlen zu übersetzen. Hier werden Bibliotheken wie `matplotlib` für statische Grafiken oder `plotly` für interaktive Visualisierungen eingesetzt. Interaktive Grafiken sind besonders wertvoll, da sie es dem Nutzer ermöglichen, Details zu erkunden, zu zoomen und spezifische Datenpunkte abzufragen.

### 6.1. Berechnung finanzieller KPIs (`calculate_financial_kpis`)

Bevor die Visualisierungen erstellt werden, müssen die finanziellen Kennzahlen berechnet werden. Diese Funktion wird von den Szenarien-Modulen aufgerufen.

```python
import numpy as np
import pandas as pd

def calculate_financial_kpis(
    annual_energy_cost_with_battery: float, # Jährliche Energiekosten mit Batteriespeicher
    total_consumption: float, # Gesamter jährlicher Verbrauch
    total_pv_generation: float, # Gesamte jährliche PV-Erzeugung
    battery_capacity_kwh: float,
    battery_cost_per_kwh: float,
    installation_cost_fixed: float,
    project_lifetime_years: int,
    discount_rate: float,
    price_grid_per_kwh, # Kann float oder pd.Series sein
    price_feed_in_per_kwh # Kann float oder pd.Series sein
) -> dict:
    """
    Berechnet finanzielle KPIs wie Amortisationszeit und Net Present Value (NPV).
    """
    # Investitionskosten
    investment_cost = (battery_capacity_kwh * battery_cost_per_kwh) + installation_cost_fixed

    # Jährliche Stromkosten ohne Speicher (Referenz)
    # Annahme: Ohne Speicher wird der gesamte Verbrauch aus dem Netz bezogen, und die gesamte PV-Erzeugung eingespeist.
    # Dies ist eine vereinfachte Annahme. Eine genauere Referenzsimulation wäre besser.
    if isinstance(price_grid_per_kwh, pd.Series):
        annual_cost_no_battery = (total_consumption * price_grid_per_kwh.mean()) - (total_pv_generation * price_feed_in_per_kwh.mean())
    else:
        annual_cost_no_battery = (total_consumption * price_grid_per_kwh) - (total_pv_generation * price_feed_in_per_kwh)
    
    # Jährliche Ersparnis durch den Speicher
    annual_savings = annual_cost_no_battery - annual_energy_cost_with_battery

    # Amortisationszeit (Payback Period)
    payback_period = np.nan # Initialisiere mit NaN
    if annual_savings > 0:
        payback_period = investment_cost / annual_savings

    # Net Present Value (NPV)
    cash_flows = [-investment_cost] # Initialinvestition
    for _ in range(project_lifetime_years):
        cash_flows.append(annual_savings) # Jährliche Ersparnisse als Cash-Inflow
    
    npv = np.npv(discount_rate, cash_flows)

    return {
        'investment_cost': investment_cost,
        'annual_savings': annual_savings,
        'payback_period_years': payback_period,
        'npv': npv
    }

```

### 6.2. Visualisierung 1: Zeitreihen-Analyse (Tages-/Wochenansicht)

Diese Visualisierung bietet einen detaillierten Einblick in die stündlichen Energieflüsse über einen ausgewählten Zeitraum (z.B. einen Tag oder eine Woche). Sie hilft dem Nutzer, die Dynamik des Systems zu verstehen.

```python
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd

def plot_energy_flows_for_period(time_series_data: pd.DataFrame, start_date: str, end_date: str):
    """
    Visualisiert die stündlichen Energieflüsse für einen ausgewählten Zeitraum.

    Args:
        time_series_data (pd.DataFrame): DataFrame mit den Zeitreihen-Ergebnissen der Simulation.
        start_date (str): Startdatum im Format 'YYYY-MM-DD'.
        end_date (str): Enddatum im Format 'YYYY-MM-DD'.
    """
    period_data = time_series_data.loc[start_date:end_date]

    fig = make_subplots(rows=2, cols=1, 
                        shared_xaxes=True, 
                        vertical_spacing=0.1,
                        row_heights=[0.7, 0.3],
                        subplot_titles=('Energieflüsse', 'Batterie Ladezustand'))

    # Energieflüsse (obere Grafik)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['Consumption_kWh'], 
                             mode='lines', name='Verbrauch', line=dict(color='red')),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['PV_Generation_kWh'], 
                             mode='lines', name='PV-Erzeugung', line=dict(color='green')),
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
                             mode='lines', name='Batterie Ladung', line=dict(color='purple', dash='dot')),
                  row=1, col=1)
    fig.add_trace(go.Scatter(x=period_data.index, y=-period_data['Battery_Discharge_kWh'], # Negativ für Entladung
                             mode='lines', name='Batterie Entladung', line=dict(color='brown', dash='dot')),
                  row=1, col=1)

    # Batterie Ladezustand (untere Grafik)
    fig.add_trace(go.Scatter(x=period_data.index, y=period_data['SOC_kWh'], 
                             mode='lines', name='Batterie SoC', line=dict(color='darkgreen')),
                  row=2, col=1)

    fig.update_layout(title_text=f'Energieflüsse von {start_date} bis {end_date}', 
                      hovermode='x unified', 
                      height=700)
    fig.update_yaxes(title_text='Energie (kWh)', row=1, col=1)
    fig.update_yaxes(title_text='SoC (kWh)', row=2, col=1)
    fig.update_xaxes(title_text='Zeit', row=2, col=1)
    fig.show()

```

### 6.3. Visualisierung 2: Optimierung der Speichergröße

Diese Grafik zeigt, wie sich wichtige Kennzahlen (Autarkiegrad, Eigenverbrauchsquote, Amortisationszeit) mit zunehmender Batteriespeichergröße entwickeln. Sie hilft, den 


„Sweet Spot“ zu identifizieren, bei dem die Wirtschaftlichkeit optimal ist.

```python
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd

def plot_optimization_curve(optimization_results: list):
    """
    Visualisiert die Optimierungskurve für die Speichergröße.

    Args:
        optimization_results (list): Liste von Dictionaries mit den Ergebnissen der Optimierung (von find_optimal_size).
    """
    df_results = pd.DataFrame(optimization_results)

    fig = make_subplots(rows=1, cols=1, 
                        specs=[[{"secondary_y": True}]])

    # Autarkiegrad und Eigenverbrauchsquote
    fig.add_trace(go.Scatter(x=df_results["battery_capacity_kwh"], y=df_results["autarky_rate"] * 100, 
                             mode=\'lines+markers\', name=\'Autarkiegrad (in %)\' , line=dict(color=\'green\')),
                  secondary_y=False)
    fig.add_trace(go.Scatter(x=df_results["battery_capacity_kwh"], y=df_results["self_consumption_rate"] * 100, 
                             mode=\'lines+markers\', name=\'Eigenverbrauchsquote (in %)\' , line=dict(color=\'darkgreen\', dash=\'dot\')),
                  secondary_y=False)

    # Amortisationszeit
    fig.add_trace(go.Scatter(x=df_results["battery_capacity_kwh"], y=df_results["payback_period_years"], 
                             mode=\'lines+markers\', name=\'Amortisationszeit (Jahre)\' , line=dict(color=\'red\')),
                  secondary_y=True)
    
    # NPV
    fig.add_trace(go.Scatter(x=df_results["battery_capacity_kwh"], y=df_results["npv"], 
                             mode=\'lines+markers\', name=\'NPV (€)\' , line=dict(color=\'blue\', dash=\'dash\')),
                  secondary_y=True)

    fig.update_layout(title_text=\'Optimierung der Batteriespeichergröße\', 
                      hovermode=\'x unified\')
    fig.update_xaxes(title_text=\'Batteriekapazität (kWh)\')
    fig.update_yaxes(title_text=\'Autarkiegrad / Eigenverbrauchsquote (in %)\' , secondary_y=False)
    fig.update_yaxes(title_text=\'Amortisationszeit (Jahre) / NPV (€)\' , secondary_y=True)
    fig.show()

```

### 6.4. Visualisierung 3: Szenarien-Vergleich (Balkendiagramm)

Dieses Balkendiagramm ermöglicht einen direkten und übersichtlichen Vergleich der wichtigsten Kennzahlen (KPIs) zwischen verschiedenen Simulationsszenarien (z.B. unterschiedliche Haushaltstypen, variable vs. fixe Tarife). Es hilft, die Auswirkungen von Parameteränderungen auf die Gesamtperformance schnell zu erfassen.

```python
import plotly.graph_objects as go
import pandas as pd

def plot_scenario_comparison(scenario_results: dict):
    """
    Visualisiert den Vergleich verschiedener Szenarien anhand von KPIs.

    Args:
        scenario_results (dict): Dictionary mit den Ergebnissen verschiedener Szenarien.
                                 Format: {\"Szenario Name\": {\"kpis\": {...}}}
    """
    kpis_to_compare = [\"autarky_rate\", \"self_consumption_rate\", \"payback_period_years\", \"npv\"]
    kpi_labels = {
        \"autarky_rate\": \"Autarkiegrad (%)\",
        \"self_consumption_rate\": \"Eigenverbrauchsquote (%)\",
        \"payback_period_years\": \"Amortisationszeit (Jahre)\",
        \"npv\": \"NPV (€)\"
    }

    scenario_names = list(scenario_results.keys())
    data = []

    for kpi in kpis_to_compare:
        values = [scenario_results[name][\"kpis\"][kpi] for name in scenario_names]
        data.append(go.Bar(name=kpi_labels[kpi], x=scenario_names, y=values))

    fig = go.Figure(data=data)
    fig.update_layout(barmode=\'group\', title_text=\'Szenarien-Vergleich der KPIs\')
    fig.update_yaxes(title_text=\'Wert\')
    fig.show()

```

### 6.5. Visualisierung 4: Sankey-Diagramm der Energieflüsse

Ein Sankey-Diagramm bietet eine hervorragende Möglichkeit, die gesamten jährlichen Energieflüsse im System visuell darzustellen. Es zeigt, woher die Energie kommt, wie sie genutzt wird und wohin sie fließt, einschließlich der Verluste. Dies ist eine sehr intuitive Darstellung für eine ganzheitliche Übersicht.

```python
import plotly.graph_objects as go

def plot_sankey_diagram(kpis: dict):
    """
    Erstellt ein Sankey-Diagramm der jährlichen Energieflüsse.

    Args:
        kpis (dict): Dictionary mit den aggregierten KPIs der Simulation.
    """
    # Definition der Knoten (Quellen, Senken, Speicher)
    labels = [
        \"PV-Erzeugung\", \"Netzbezug\", # Quellen
        \"Direkter Eigenverbrauch\", \"Batterie (Laden)\", \"Netzeinspeisung\", # Zwischenstationen
        \"Verbrauch\", \"Batterie (Entladen)\", \"Batterieverluste\" # Senken/Nutzung
    ]

    # Indizes für die Knoten
    pv_gen_idx = 0
    grid_import_idx = 1
    direct_self_cons_idx = 2
    battery_charge_idx = 3
    grid_export_idx = 4
    consumption_idx = 5
    battery_discharge_idx = 6
    battery_losses_idx = 7

    # Berechnung der Flusswerte
    total_pv_generation = kpis[\"total_pv_generation_kwh\"]
    total_consumption = kpis[\"total_consumption_kwh\"]
    total_grid_import = kpis[\"total_grid_import_kwh\"]
    total_grid_export = kpis[\"total_grid_export_kwh\"]

    # Annahmen für Sankey-Diagramm (müssen aus detaillierten Zeitreihen kommen oder geschätzt werden)
    # Diese Werte sollten idealerweise direkt aus den aggregierten Werten der Simulation kommen
    # Hier nur als Beispiel, wie die Flüsse berechnet werden könnten
    # Für genaue Werte müssen die time_series_data aus model.py aggregiert werden

    # Beispielhafte Berechnung der Flüsse (Anpassung an Ihre sim_result time_series_data notwendig)
    # Nehmen wir an, wir haben die aggregierten Werte aus der Simulation:
    # total_pv_generation
    # total_consumption
    # total_grid_import
    # total_grid_export
    # total_direct_self_consumption (aus sim_result time_series_data)
    # total_battery_charge (aus sim_result time_series_data)
    # total_battery_discharge (aus sim_result time_series_data)

    # Für dieses Beispiel nehmen wir an, dass die KPIs diese Werte enthalten:
    # kpis[\"total_direct_self_consumption_kwh\"]
    # kpis[\"total_battery_charge_kwh\"]
    # kpis[\"total_battery_discharge_kwh\"]
    # kpis[\"battery_losses_kwh\"] # Muss in model.py berechnet werden

    # Da diese Werte nicht direkt in den aktuellen KPIs sind, müssen wir sie schätzen oder die sim_result.time_series_data nutzen.
    # Für ein vollständiges Sankey-Diagramm müssten diese Werte aus der Simulation kommen.
    # Hier eine vereinfachte Annahme basierend auf den vorhandenen KPIs:

    # PV-Erzeugung -> Direkter Eigenverbrauch
    pv_to_direct_self_consumption = kpis.get(\'total_direct_self_consumption_kwh\', total_pv_generation * 0.3) # Beispielwert
    
    # PV-Erzeugung -> Batterie (Laden)
    pv_to_battery_charge = kpis.get(\'total_battery_charge_kwh\', total_pv_generation * 0.4) # Beispielwert
    
    # PV-Erzeugung -> Netzeinspeisung
    pv_to_grid_export = total_pv_generation - pv_to_direct_self_consumption - pv_to_battery_charge
    pv_to_grid_export = max(0, pv_to_grid_export) # Kann nicht negativ sein

    # Netzbezug -> Verbrauch
    grid_import_to_consumption = total_grid_import # Annahme: Netzbezug geht direkt an Verbrauch

    # Batterie (Entladen) -> Verbrauch
    battery_discharge_to_consumption = kpis.get(\'total_battery_discharge_kwh\', total_pv_generation * 0.2) # Beispielwert

    # Batterie (Laden) -> Batterieverluste (Annahme: 1 - Wirkungsgrad)
    battery_losses = pv_to_battery_charge * (1 - kpis.get(\'battery_efficiency_charge\', 0.9)) # Beispielwert

    # Quellen und Ziele der Flüsse
    source = [
        pv_gen_idx, pv_gen_idx, pv_gen_idx, # PV-Erzeugung
        grid_import_idx, # Netzbezug
        battery_charge_idx, # Batterie (Laden)
        battery_discharge_idx # Batterie (Entladen)
    ]
    target = [
        direct_self_cons_idx, battery_charge_idx, grid_export_idx, # PV-Erzeugung
        consumption_idx, # Netzbezug
        battery_discharge_idx, battery_losses_idx, # Batterie (Laden)
        consumption_idx # Batterie (Entladen)
    ]
    value = [
        pv_to_direct_self_consumption,
        pv_to_battery_charge,
        pv_to_grid_export,
        grid_import_to_consumption,
        pv_to_battery_charge, # Hier muss der Wert der Ladung rein
        battery_losses,
        battery_discharge_to_consumption
    ]

    # Filter out zero values to avoid drawing unnecessary links
    filtered_source = []
    filtered_target = []
    filtered_value = []
    for s, t, v in zip(source, target, value):
        if v > 0.01: # Schwellenwert, um sehr kleine Flüsse zu ignorieren
            filtered_source.append(s)
            filtered_target.append(t)
            filtered_value.append(v)

    node = dict(label=labels, pad=15, thickness=20, line=dict(color=\"black\", width=0.5))
    link = dict(source=filtered_source, target=filtered_target, value=filtered_value)

    fig = go.Figure(data=[go.Sankey(node=node, link=link)])
    fig.update_layout(title_text=\"Jährliche Energieflüsse\", font_size=10)
    fig.show()

```

**Wichtiger Hinweis zum Sankey-Diagramm:** Die genaue Berechnung der Flüsse für das Sankey-Diagramm erfordert eine detailliertere Aggregation der `time_series_data` aus der `simulate_one_year`-Funktion. Insbesondere müssen die genauen Mengen für `total_direct_self_consumption_kwh`, `total_battery_charge_kwh`, `total_battery_discharge_kwh` und `battery_losses_kwh` aus den Zeitreihen summiert werden. Die obigen Beispielwerte sind Platzhalter und müssen durch die tatsächlichen Simulationsergebnisse ersetzt werden.

## 7. Benutzeroberfläche (`ui.py` mit Streamlit)

Eine interaktive Benutzeroberfläche ist entscheidend, um das Analyse-Tool für Endnutzer zugänglich und bedienbar zu machen. Streamlit ist eine ausgezeichnete Wahl für die schnelle Entwicklung von Daten-Apps in Python, da es minimale Frontend-Kenntnisse erfordert und direkt aus Python-Skripten heraus interaktive Webanwendungen erstellt. Die `ui.py` wird die verschiedenen Module (`data_import`, `model`, `scenarios`, `analysis`) zusammenführen und die Interaktion mit dem Nutzer ermöglichen.

### 7.1. Grundstruktur der Streamlit-App

```python
import streamlit as st
import pandas as pd

# Importieren der Module
from data_import import load_data, preprocess_data
from model import simulate_one_year
from scenarios import find_optimal_size, run_variable_tariff_scenario, compare_household_scenarios
from analysis import calculate_financial_kpis, plot_energy_flows_for_period, plot_optimization_curve, plot_scenario_comparison, plot_sankey_diagram
# from config import * # Importiere alle Konstanten aus config.py

st.set_page_config(layout=\"wide\", page_title=\"Batteriespeicher Analyse Tool\")

st.title(\"Batteriespeicher Analyse Tool\")
st.markdown(\"Analysieren Sie die optimale Größe und Wirtschaftlichkeit von Batteriespeichern.\")

# --- Seitenleiste für Parameter-Eingabe ---
st.sidebar.header(\"Simulationseinstellungen\")

# Dateiupload
uploaded_file = st.sidebar.file_uploader(\"Laden Sie Ihre Verbrauchs- und PV-Daten (Excel)\", type=\"xlsx\")

if uploaded_file is not None:
    df_stromspeicher, df_verbrauch = load_data(uploaded_file)
    if df_stromspeicher is not None and df_verbrauch is not None:
        parameters, consumption_series, pv_generation_series = preprocess_data(df_stromspeicher, df_verbrauch)
        st.sidebar.success(\"Daten erfolgreich geladen und vorverarbeitet.\")

        # --- Allgemeine Parameter ---
        st.sidebar.subheader(\"Allgemeine Parameter\")
        pv_system_size_kwp = st.sidebar.slider(
            \"PV-Anlagengröße (kWp)\", min_value=1.0, max_value=50.0, value=10.0, step=0.5
        )
        annual_consumption_kwh = st.sidebar.slider(
            \"Jährlicher Stromverbrauch (kWh)\", min_value=1000, max_value=100000, value=5000, step=500
        )
        # Annahme: consumption_series und pv_generation_series werden skaliert basierend auf den Slider-Werten
        # Dies erfordert eine Anpassung in preprocess_data oder hier, um die Zeitreihen zu skalieren.
        # Beispiel: consumption_series = consumption_series * (annual_consumption_kwh / consumption_series.sum())

        # --- Batterieparameter ---
        st.sidebar.subheader(\"Batterieparameter\")
        battery_capacity_kwh = st.sidebar.slider(
            \"Batteriekapazität (kWh)\", min_value=0.0, max_value=30.0, value=10.0, step=0.5
        )
        battery_cost_per_kwh = st.sidebar.number_input(
            \"Batteriekosten pro kWh (€/kWh)\", min_value=100, max_value=1000, value=500, step=10
        )
        installation_cost_fixed = st.sidebar.number_input(
            \"Feste Installationskosten (€)\", min_value=0, max_value=10000, value=2000, step=100
        )
        battery_efficiency_charge = st.sidebar.slider(
            \"Ladeeffizienz Batterie (%)\", min_value=80, max_value=100, value=95, step=1
        ) / 100
        battery_efficiency_discharge = st.sidebar.slider(
            \"Entladeeffizienz Batterie (%)\", min_value=80, max_value=100, value=95, step=1
        ) / 100
        battery_max_charge_kw = st.sidebar.slider(
            \"Max. Ladeleistung Batterie (kW)\", min_value=1.0, max_value=20.0, value=5.0, step=0.5
        )
        battery_max_discharge_kw = st.sidebar.slider(
            \"Max. Entladeleistung Batterie (kW)\", min_value=1.0, max_value=20.0, value=5.0, step=0.5
        )

        # --- Strompreise ---
        st.sidebar.subheader(\"Strompreise\")
        price_grid_per_kwh = st.sidebar.number_input(
            \"Strombezugspreis (ct/kWh)\", min_value=10.0, max_value=50.0, value=30.0, step=0.5
        ) / 100
        price_feed_in_per_kwh = st.sidebar.number_input(
            \"Einspeisevergütung (ct/kWh)\", min_value=5.0, max_value=20.0, value=8.0, step=0.5
        ) / 100

        # --- Finanzielle Parameter ---
        st.sidebar.subheader(\"Finanzielle Parameter\")
        project_lifetime_years = st.sidebar.slider(
            \"Projektlaufzeit (Jahre)\", min_value=5, max_value=30, value=20, step=1
        )
        discount_rate = st.sidebar.slider(
            \"Diskontierungsrate (%)\", min_value=0.0, max_value=10.0, value=5.0, step=0.1
        ) / 100

        # --- Szenario-Auswahl ---
        st.sidebar.subheader(\"Szenario-Auswahl\")
        selected_scenario = st.sidebar.selectbox(
            \"Wählen Sie ein Szenario\",
            (\"Einzel-Simulation\", \"Optimale Speichergröße\", \"Haushaltstypen-Vergleich\", \"Variable Tarife\")
        )

        if st.sidebar.button(\"Simulation starten\"):
            st.subheader(\"Simulationsergebnisse\")

            if selected_scenario == \"Einzel-Simulation\":
                with st.spinner(\"Führe Einzel-Simulation durch...\"):
                    sim_result = simulate_one_year(
                        consumption_series=consumption_series,
                        pv_generation_series=pv_generation_series,
                        battery_capacity_kwh=battery_capacity_kwh,
                        battery_efficiency_charge=battery_efficiency_charge,
                        battery_efficiency_discharge=battery_efficiency_discharge,
                        battery_max_charge_kw=battery_max_charge_kw,
                        battery_max_discharge_kw=battery_max_discharge_kw,
                        price_grid_per_kwh=price_grid_per_kwh,
                        price_feed_in_per_kwh=price_feed_in_per_kwh
                    )
                    st.success(\"Simulation abgeschlossen!\")

                    # Finanzielle KPIs berechnen
                    financial_kpis = calculate_financial_kpis(
                        sim_result[\"kpis\"][\"annual_energy_cost\"],
                        total_consumption=sim_result[\"kpis\"][\"total_consumption_kwh\"],
                        total_pv_generation=sim_result[\"kpis\"][\"total_pv_generation_kwh\"],
                        battery_capacity_kwh=battery_capacity_kwh,
                        battery_cost_per_kwh=battery_cost_per_kwh,
                        installation_cost_fixed=installation_cost_fixed,
                        project_lifetime_years=project_lifetime_years,
                        discount_rate=discount_rate,
                        price_grid_per_kwh=price_grid_per_kwh,
                        price_feed_in_per_kwh=price_feed_in_per_kwh
                    )
                    sim_result[\"kpis\"] = {**sim_result[\"kpis\"], **financial_kpis}

                    st.write(\"### Wichtige Kennzahlen (KPIs)\")
                    st.json(sim_result[\"kpis\"])

                    st.write(\"### Energieflüsse über die Zeit\")
                    # Datumsauswahl für Zeitreihen-Visualisierung
                    col1, col2 = st.columns(2)
                    start_date_plot = col1.date_input(\"Startdatum für Zeitreihen-Plot\", value=sim_result[\"time_series_data\"].index.min())
                    end_date_plot = col2.date_input(\"Enddatum für Zeitreihen-Plot\", value=sim_result[\"time_series_data\"].index.min() + pd.Timedelta(days=7))
                    
                    plot_energy_flows_for_period(sim_result[\"time_series_data\"], str(start_date_plot), str(end_date_plot))

                    st.write(\"### Jährliche Energieflüsse (Sankey-Diagramm)\")
                    plot_sankey_diagram(sim_result[\"kpis\"])

            elif selected_scenario == \"Optimale Speichergröße\":
                st.write(\"### Optimale Speichergröße ermitteln\")
                min_cap = st.sidebar.number_input(\"Min. Kapazität (kWh)\", value=0.0, step=0.5)
                max_cap = st.sidebar.number_input(\"Max. Kapazität (kWh)\", value=20.0, step=0.5)
                step_cap = st.sidebar.number_input(\"Schrittweite (kWh)\", value=1.0, step=0.5)

                with st.spinner(\"Ermittle optimale Speichergröße...\"):
                    optimization_results = find_optimal_size(
                        consumption_series=consumption_series,
                        pv_generation_series=pv_generation_series,
                        min_capacity_kwh=min_cap,
                        max_capacity_kwh=max_cap,
                        step_kwh=step_cap,
                        battery_efficiency_charge=battery_efficiency_charge,
                        battery_efficiency_discharge=battery_efficiency_discharge,
                        battery_max_charge_kw=battery_max_charge_kw,
                        battery_max_discharge_kw=battery_max_discharge_kw,
                        price_grid_per_kwh=price_grid_per_kwh,
                        price_feed_in_per_kwh=price_feed_in_per_kwh,
                        battery_cost_per_kwh=battery_cost_per_kwh,
                        installation_cost_fixed=installation_cost_fixed,
                        project_lifetime_years=project_lifetime_years,
                        discount_rate=discount_rate
                    )
                    st.success(\"Optimierung abgeschlossen!\")
                    plot_optimization_curve(optimization_results)

            elif selected_scenario == \"Haushaltstypen-Vergleich\":
                st.write(\"### Vergleich verschiedener Haushaltstypen\")
                st.warning(\"Bitte stellen Sie sicher, dass entsprechende Last- und PV-Profile für die verschiedenen Haushaltstypen in `data_import.py` verfügbar sind oder hier geladen werden können.\")
                
                # Beispielhafte Dummy-Profile für den Vergleich
                # In einer realen Anwendung müssten diese Profile geladen oder generiert werden
                household_profiles = {
                    \"Einfamilienhaus\": {
                        \"consumption\": consumption_series, # Hier echtes EFH-Profil
                        \"pv_gen\": pv_generation_series # Hier echtes EFH-PV-Profil
                    },
                    \"Mehrfamilienhaus\": {
                        \"consumption\": consumption_series * 5, # Beispiel: 5x Verbrauch
                        \"pv_gen\": pv_generation_series * 3 # Beispiel: 3x PV
                    },
                    \"Quartier\": {
                        \"consumption\": consumption_series * 20, # Beispiel: 20x Verbrauch
                        \"pv_gen\": pv_generation_series * 15 # Beispiel: 15x PV
                    }
                }

                with st.spinner(\"Vergleiche Haushaltstypen...\"):
                    comparison_results = compare_household_scenarios(
                        household_profiles=household_profiles,
                        battery_capacity_kwh=battery_capacity_kwh,
                        battery_efficiency_charge=battery_efficiency_charge,
                        battery_efficiency_discharge=battery_efficiency_discharge,
                        battery_max_charge_kw=battery_max_charge_kw,
                        battery_max_discharge_kw=battery_max_discharge_kw,
                        price_grid_per_kwh=price_grid_per_kwh,
                        price_feed_in_per_kwh=price_feed_in_per_kwh,
                        battery_cost_per_kwh=battery_cost_per_kwh,
                        installation_cost_fixed=installation_cost_fixed,
                        project_lifetime_years=project_lifetime_years,
                        discount_rate=discount_rate
                    )
                    st.success(\"Vergleich abgeschlossen!\")
                    plot_scenario_comparison(comparison_results)

            elif selected_scenario == \"Variable Tarife\":
                st.write(\"### Simulation mit variablen Tarifen\")
                st.warning(\"Für dieses Szenario benötigen Sie stündliche Strompreisdaten (Bezug und Einspeisung). Bitte stellen Sie sicher, dass diese in `data_import.py` geladen werden können.\")
                
                # Beispielhafte Dummy-Preiszeitreihen
                # In einer realen Anwendung müssten diese aus einer Datei geladen werden
                # oder über eine API bezogen werden.
                # Hier: Einfache Variation des Preises über den Tag
                hourly_prices_grid = pd.Series(price_grid_per_kwh * (1 + 0.2 * np.sin(np.linspace(0, 2 * np.pi * 365, len(consumption_series)))), index=consumption_series.index)
                hourly_prices_feed_in = pd.Series(price_feed_in_per_kwh * (1 - 0.1 * np.sin(np.linspace(0, 2 * np.pi * 365, len(consumption_series)))), index=consumption_series.index)

                with st.spinner(\"Führe Simulation mit variablen Tarifen durch...\"):
                    variable_tariff_result = run_variable_tariff_scenario(
                        consumption_series=consumption_series,
                        pv_generation_series=pv_generation_series,
                        battery_capacity_kwh=battery_capacity_kwh,
                        battery_efficiency_charge=battery_efficiency_charge,
                        battery_efficiency_discharge=battery_efficiency_discharge,
                        battery_max_charge_kw=battery_max_charge_kw,
                        battery_max_discharge_kw=battery_max_discharge_kw,
                        price_grid_series=hourly_prices_grid,
                        price_feed_in_series=hourly_prices_feed_in,
                        battery_cost_per_kwh=battery_cost_per_kwh,
                        installation_cost_fixed=installation_cost_fixed,
                        project_lifetime_years=project_lifetime_years,
                        discount_rate=discount_rate
                    )
                    st.success(\"Simulation abgeschlossen!\")
                    st.write(\"### Wichtige Kennzahlen (KPIs) für variable Tarife\")
                    st.json(variable_tariff_result[\"kpis\"])

                    st.write(\"### Energieflüsse über die Zeit (Variable Tarife)\")
                    col1, col2 = st.columns(2)
                    start_date_plot = col1.date_input(\"Startdatum für Zeitreihen-Plot (Variable Tarife)\", value=variable_tariff_result[\"time_series_data\"].index.min())
                    end_date_plot = col2.date_input(\"Enddatum für Zeitreihen-Plot (Variable Tarife)\", value=variable_tariff_result[\"time_series_data\"].index.min() + pd.Timedelta(days=7))
                    
                    plot_energy_flows_for_period(variable_tariff_result[\"time_series_data\"], str(start_date_plot), str(end_date_plot))

                    st.write(\"### Jährliche Energieflüsse (Sankey-Diagramm für variable Tarife)\")
                    plot_sankey_diagram(variable_tariff_result[\"kpis\"])

    else:
        st.info(\"Bitte laden Sie eine Excel-Datei, um die Simulation zu starten.\")

```

### 7.2. Ausführung der Streamlit-App

Um die Streamlit-Anwendung zu starten, speichern Sie den obigen Code als `ui.py` (oder `main.py`, wenn Sie es als Hauptskript verwenden) und führen Sie den folgenden Befehl im Terminal aus:

```bash
streamlit run ui.py
```

Dies öffnet die Anwendung in Ihrem Webbrowser. Änderungen am Code werden automatisch erkannt und die App neu geladen, was eine schnelle Entwicklung ermöglicht.

## 8. `config.py` (Zentrale Konfiguration)

Die `config.py`-Datei dient dazu, alle festen Parameter, Schwellenwerte und Standardwerte zu zentralisieren. Dies macht das System flexibler und einfacher zu warten, da Änderungen an diesen Werten nur an einer Stelle vorgenommen werden müssen.

```python
# config.py

# Standardwerte für Strompreise (falls nicht aus Datei geladen)
DEFAULT_PRICE_GRID_PER_KWH = 0.30 # Euro/kWh
DEFAULT_PRICE_FEED_IN_PER_KWH = 0.08 # Euro/kWh

# Standardwerte für Batterieeffizienz und Leistung
DEFAULT_BATTERY_EFFICIENCY_CHARGE = 0.95
DEFAULT_BATTERY_EFFICIENCY_DISCHARGE = 0.95
DEFAULT_BATTERY_MAX_CHARGE_KW = 5.0
DEFAULT_BATTERY_MAX_DISCHARGE_KW = 5.0

# Standardwerte für Batteriekosten
DEFAULT_BATTERY_COST_PER_KWH = 500 # Euro/kWh
DEFAULT_INSTALLATION_COST_FIXED = 2000 # Euro

# Finanzielle Parameter
DEFAULT_PROJECT_LIFETIME_YEARS = 20
DEFAULT_DISCOUNT_RATE = 0.05

# Schwellenwerte für variable Tarife (Arbitrage-Steuerung)
# Diese Werte müssen an die spezifischen Preisstrukturen angepasst werden
THRESHOLD_BUY_PRICE = 0.15 # Euro/kWh: Wenn der Preis darunter liegt, aus Netz laden
THRESHOLD_SELL_PRICE = 0.40 # Euro/kWh: Wenn der Preis darüber liegt, ins Netz entladen
MAX_GRID_CHARGE_KW = 3.0 # Max. Leistung für Laden aus Netz
MAX_GRID_DISCHARGE_KW = 3.0 # Max. Leistung für Entladen ins Netz

# Weitere Konstanten können hier hinzugefügt werden
# Beispiel: Pfade zu Standard-Lastprofilen, etc.

```

## 9. `README.md` (Projektdokumentation)

Eine gute `README.md`-Datei ist entscheidend für die Nutzbarkeit und Verständlichkeit Ihres Projekts. Sie sollte alle notwendigen Informationen enthalten, damit andere (oder Sie selbst in der Zukunft) das Projekt verstehen, installieren und ausführen können.

```markdown
# Batteriespeicher Analyse Tool

## Übersicht

Dieses Python-basierte Tool ermöglicht die Analyse und Optimierung der Größe von Batteriespeichern für verschiedene Anwendungsfälle (Einfamilienhaus, Mehrfamilienhaus, Quartier). Es simuliert Energieflüsse auf Stundenbasis und visualisiert die Auswirkungen von Parametern wie PV-Anlagengröße, Strompreisen und Speicherkosten auf wichtige Kennzahlen wie Autarkiegrad, Eigenverbrauchsquote, Amortisationszeit und Net Present Value (NPV).

## Funktionen

*   **Datenimport**: Laden von Verbrauchs- und PV-Erzeugungsdaten aus Excel-Dateien.
*   **Stündliche Simulation**: Detaillierte Modellierung der Energieflüsse.
*   **Szenarienanalyse**: Untersuchung der optimalen Speichergröße, des Einflusses variabler Stromtarife und des Vergleichs verschiedener Haushaltstypen.
*   **Interaktive Visualisierungen**: Darstellung der Ergebnisse mittels Plotly-Diagrammen (Zeitreihen, Optimierungskurven, Balkendiagramme, Sankey-Diagramme).
*   **Benutzeroberfläche**: Eine intuitive Web-Oberfläche basierend auf Streamlit zur einfachen Parameteranpassung und Ergebnisdarstellung.

## Installation

1.  **Python installieren**: Stellen Sie sicher, dass Python 3.8 oder höher auf Ihrem System installiert ist.

2.  **Virtuelle Umgebung erstellen (empfohlen)**:
    ```bash
    python3 -m venv venv
    source venv/bin/activate  # Unter Windows: .\venv\Scripts\activate
    ```

3.  **Abhängigkeiten installieren**:
    ```bash
    pip install pandas openpyxl numpy plotly streamlit
    ```

## Nutzung

1.  **Projektstruktur einrichten**:
    Erstellen Sie die folgenden Dateien in Ihrem Projektverzeichnis:
    *   `main.py` (oder `ui.py`)
    *   `data_import.py`
    *   `model.py`
    *   `scenarios.py`
    *   `analysis.py`
    *   `config.py`
    *   `README.md` (diese Datei)

2.  **Code einfügen**:
    Kopieren Sie den entsprechenden Code aus der Anleitung in die jeweiligen Dateien.

3.  **Daten vorbereiten**:
    Erstellen Sie eine Excel-Datei (z.B. `data.xlsx`) mit zwei Sheets:
    *   **`Stromspeicher_Parameter`**: Enthält Parameter wie in `data_import.py` erwartet (z.B. eine Spalte 'Parameter' und eine Spalte 'Wert').
    *   **`Verbrauchsdaten`**: Enthält stündliche Zeitreihen mit Spalten wie 'Zeitstempel', 'Verbrauch_kWh', 'PV_Erzeugung_kWh'. Stellen Sie sicher, dass die Zeitstempel lückenlos sind und ein ganzes Jahr abdecken (8760 Stunden).

4.  **Anwendung starten**:
    Öffnen Sie ein Terminal im Projektverzeichnis und führen Sie aus:
    ```bash
    streamlit run ui.py
    ```
    (oder `streamlit run main.py`, je nachdem, wie Sie Ihre Hauptdatei genannt haben).

5.  **Interaktion im Browser**:
    Die Anwendung öffnet sich automatisch in Ihrem Standard-Webbrowser. Sie können nun Ihre Excel-Datei hochladen, Parameter in der Seitenleiste anpassen und verschiedene Simulationsszenarien auswählen und starten.

## Weiterentwicklung

*   **Erweiterte Batteriesteuerung**: Implementierung von Optimierungsalgorithmen (z.B. lineare Programmierung) für die Batteriesteuerung unter variablen Stromtarifen.
*   **Kostenmodell**: Detaillierteres Kostenmodell für PV-Anlage, Speicher und Betrieb (z.B. Wartung, Degradation).
*   **Mehr Lastprofile**: Integration einer Bibliothek von Standardlastprofilen (SLP) oder die Möglichkeit, eigene Profile zu importieren.
*   **Wetterdaten**: Einbindung von Wetterdaten für genauere PV-Erzeugungsprognosen.
*   **Berichtserstellung**: Export der Ergebnisse und Grafiken in PDF-Berichte.

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz. Weitere Details finden Sie in der `LICENSE`-Datei.

```

## 10. Nächste Schritte

Nachdem Sie diese detaillierte Anleitung durchgearbeitet haben, können Sie mit der Implementierung beginnen. Es wird empfohlen, die Module schrittweise zu entwickeln und zu testen:

1.  **`config.py`**: Definieren Sie Ihre Konstanten.
2.  **`data_import.py`**: Stellen Sie sicher, dass Ihre Daten korrekt geladen und vorverarbeitet werden.
3.  **`model.py`**: Implementieren und testen Sie die Kernsimulationslogik. Überprüfen Sie die Energiebilanz.
4.  **`analysis.py`**: Implementieren Sie die KPI-Berechnungen und die Visualisierungsfunktionen. Testen Sie jede Grafik einzeln.
5.  **`scenarios.py`**: Orchestrieren Sie die verschiedenen Simulationsläufe.
6.  **`ui.py`**: Bauen Sie die Streamlit-Oberfläche schrittweise auf und integrieren Sie die anderen Module.
7.  **`README.md`**: Pflegen Sie die Dokumentation während der Entwicklung.

Beginnen Sie mit einem einfachen Datensatz und erweitern Sie die Komplexität schrittweise. Viel Erfolg bei der Entwicklung Ihres Batteriespeicher Analyse Tools!

