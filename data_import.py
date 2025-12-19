import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import holidays
import io

# Streamlit Caching für teure Excel-Ladeoperationen
try:
    import streamlit as st
except ImportError:
    # Dummy streamlit falls nicht verfügbar
    class DummyStreamlit:
        @staticmethod
        def cache_data(func):
            return func
    st = DummyStreamlit()

def create_weekday_sheet(year=2024):
    """
    Erstellt ein Sheet mit Wochentag-Typen für ein ganzes Jahr.
    WT = Werktag (Mo-Fr, außer Feiertage)
    SA = Samstag
    FT = Sonn-/Feiertag (Sonntage + gesetzliche Feiertage)
    
    Args:
        year (int): Jahr für das Sheet
        
    Returns:
        pd.DataFrame: DataFrame mit Datum und Wochentag-Typen für alle Bundesländer
    """
    # Bundesländer und ihre Codes
    bundeslaender = {
        'BW': 'Baden-Württemberg',
        'BY': 'Bayern', 
        'BE': 'Berlin',
        'BB': 'Brandenburg',
        'HB': 'Bremen',
        'HH': 'Hamburg',
        'HE': 'Hessen',
        'MV': 'Mecklenburg-Vorpommern',
        'NI': 'Niedersachsen',
        'NW': 'Nordrhein-Westfalen',
        'RP': 'Rheinland-Pfalz',
        'SL': 'Saarland',
        'SN': 'Sachsen',
        'ST': 'Sachsen-Anhalt',
        'SH': 'Schleswig-Holstein',
        'TH': 'Thüringen'
    }
    
    # Erstelle Datum-Serie für das ganze Jahr
    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # Erstelle DataFrame
    df = pd.DataFrame({'Datum': date_range})
    df['Wochentag'] = df['Datum'].dt.day_name()
    df['Wochentag_Deutsch'] = df['Datum'].dt.day_name().map({
        'Monday': 'Montag',
        'Tuesday': 'Dienstag', 
        'Wednesday': 'Mittwoch',
        'Thursday': 'Donnerstag',
        'Friday': 'Freitag',
        'Saturday': 'Samstag',
        'Sunday': 'Sonntag'
    })
    
    # Initialisiere alle Bundesländer mit Standard-Wochentag-Typ
    for code in bundeslaender.keys():
        df[f'{code}_Tagestyp'] = 'WT'  # Standard: Werktag
    
    # Markiere Samstage
    df.loc[df['Wochentag'] == 'Saturday', [f'{code}_Tagestyp' for code in bundeslaender.keys()]] = 'SA'
    
    # Markiere Sonntage
    df.loc[df['Wochentag'] == 'Sunday', [f'{code}_Tagestyp' for code in bundeslaender.keys()]] = 'FT'
    
    # Markiere Feiertage für jedes Bundesland
    for code, name in bundeslaender.items():
        # Erstelle Feiertags-Objekt für das Bundesland
        try:
            if code == 'BY':
                # Bayern hat zusätzliche Feiertage
                bundesland_holidays = holidays.Germany(state=code, years=year)
            else:
                bundesland_holidays = holidays.Germany(state=code, years=year)
            
            # Markiere Feiertage als FT
            for date, holiday_name in bundesland_holidays.items():
                df.loc[df['Datum'].dt.date == date, f'{code}_Tagestyp'] = 'FT'
                
        except Exception as e:
            print(f"Fehler beim Laden der Feiertage für {name} ({code}): {e}")
    
    # Füge Beschreibung hinzu
    df.insert(0, 'Beschreibung', 'Wochentag-Typen für Lastprofile: WT=Werktag, SA=Samstag, FT=Sonn-/Feiertag')
    
    return df

def get_day_type_for_date(date, bundesland_code):
    """
    Bestimmt den Wochentag-Typ für ein bestimmtes Datum und Bundesland.
    
    Args:
        date (datetime): Datum
        bundesland_code (str): Bundesland-Code (z.B. 'BW', 'BY', etc.)
        
    Returns:
        str: 'WT', 'SA', oder 'FT'
    """
    # Wochentag bestimmen
    weekday = date.weekday()  # 0=Montag, 6=Sonntag
    
    if weekday == 5:  # Samstag
        return 'SA'
    elif weekday == 6:  # Sonntag
        return 'FT'
    else:
        # Werktag - prüfe ob Feiertag
        try:
            bundesland_holidays = holidays.Germany(state=bundesland_code, years=date.year)
            if date.date() in bundesland_holidays:
                return 'FT'
            else:
                return 'WT'
        except:
            # Fallback: Werktag
            return 'WT'

def load_standard_load_profile(file_path, profile_type, annual_consumption_kwh, bundesland_code='BW', year=2024):
    """
    Lädt ein Standard-Lastprofil aus der Excel-Datei und skaliert es auf den gewünschten Jahresverbrauch.
    Robuste Implementierung: extrahiert alle numerischen 15-Minuten-Werte unabhängig von Spaltennamen
    und bildet daraus 8760 Stunden für das Jahr.
    
    Args:
        file_path: Pfad zur Excel-Datei
        profile_type: Typ des Lastprofils ('H25', 'G25', 'L25', 'P25', 'S25')
        annual_consumption_kwh: Gewünschter Jahresverbrauch in kWh
        bundesland_code: Bundesland-Code für Feiertagsberechnung
        year: Jahr für die Wochentag-Berechnung
        
    Returns:
        pd.Series: Skaliertes Lastprofil mit stündlichen Werten für ein Jahr
    """
    try:
        # Sheet laden (keine Annahmen über Header-Struktur)
        df_raw = pd.read_excel(file_path, sheet_name=profile_type, header=None)

        # In numerisch umwandeln; Nicht-Zahlen werden NaN
        df_num = df_raw.apply(pd.to_numeric, errors='coerce')

        # Vollständig leere Zeilen/Spalten entfernen
        df_num = df_num.dropna(how='all')
        df_num = df_num.dropna(axis=1, how='all')

        # Alle Zahlen zu einem flachen Vektor zusammenfassen (Zeile-für-Zeile)
        values = df_num.to_numpy().astype(float).ravel()
        values = values[np.isfinite(values)]

        # Ziel-Länge für 15-Minuten-Werte: 365/366 Tage * 24 h * 4 = 35040/35136
        # Prüfe ob Schaltjahr
        is_leap_year = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        days_in_year = 366 if is_leap_year else 365
        target_len_15min = days_in_year * 24 * 4

        fifteen_min_vector = None

        if len(values) >= target_len_15min:
            # Nimm die ersten Werte
            fifteen_min_vector = values[:target_len_15min]
        else:
            # Versuche sinnvolle Rekonstruktion
            if len(values) == 96:
                # Ein Tagesprofil -> Tage im Jahr mal wiederholen
                fifteen_min_vector = np.tile(values, days_in_year)
            elif len(values) > 0 and len(values) % 96 == 0:
                # Mehrere Tage vorhanden -> auf Tage im Jahr strecken (abschneiden/auffüllen)
                num_days = len(values) // 96
                if num_days >= days_in_year:
                    fifteen_min_vector = values[:target_len_15min]
                else:
                    repeats = (days_in_year + num_days - 1) // num_days
                    tiled = np.tile(values, repeats)
                    fifteen_min_vector = tiled[:target_len_15min]
            else:
                # Letzter Fallback: Gleichverteilung
                per_15min = annual_consumption_kwh / target_len_15min
                fifteen_min_vector = np.full(target_len_15min, per_15min)

        # Auf gewünschten Jahresverbrauch skalieren (15-Minuten-Werte)
        # Die Profile sind normiert auf 1.000.000 kWh
        # Berechne den Skalierungsfaktor: gewünschter Verbrauch / 1.000.000
        scaling_factor = annual_consumption_kwh / 1_000_000.0
        fifteen_min_vector = fifteen_min_vector * scaling_factor

        # Zeitindex erzeugen (15-Minuten-Intervalle)
        start_date = pd.Timestamp(f'{year}-01-01 00:00:00')
        time_index = pd.date_range(start=start_date, periods=target_len_15min, freq='15min')
        result_series = pd.Series(fifteen_min_vector, index=time_index, name='Verbrauch_kWh')

        # Debug-Ausgabe zum validieren
        print(f"Profil '{profile_type}' für {year}: Eingelesene 15-min Werte: {len(fifteen_min_vector)}; Jahresverbrauch: {result_series.sum():.2f} kWh")

        # Validiere Energiebilanz nach Skalierung (Qualitätssicherung)
        actual_annual = result_series.sum()
        expected_annual = annual_consumption_kwh
        deviation_kwh = actual_annual - expected_annual
        deviation_percent = abs(deviation_kwh) / expected_annual * 100 if expected_annual > 0 else 0
        
        # Toleranz: 0.1% (Rundungsfehler sind OK)
        TOLERANCE_PERCENT = 0.1
        
        if deviation_percent > TOLERANCE_PERCENT:
            print(f"⚠️ WARNUNG: Lastprofil-Energiebilanz-Abweichung!")
            print(f"  Profil: {profile_type}")
            print(f"  Jahr: {year}")
            print(f"  Erwartet: {expected_annual:.2f} kWh")
            print(f"  Tatsächlich: {actual_annual:.2f} kWh")
            print(f"  Differenz: {deviation_kwh:+.2f} kWh ({deviation_percent:.3f}%)")
            print(f"  Toleranz: {TOLERANCE_PERCENT}%")
        else:
            # Debug-Ausgabe für erfolgreiche Validierung
            print(f"✅ Lastprofil-Energiebilanz OK: {profile_type}, {deviation_percent:.4f}% Abweichung")

        return result_series

    except Exception as e:
        print(f"Fehler beim Laden des Lastprofils {profile_type}: {e}")
        # Fallback: Gleichmäßige Verteilung über das Jahr
        is_leap_year = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        days_in_year = 366 if is_leap_year else 365
        hours_in_year = days_in_year * 24
        
        start_date = pd.Timestamp(f'{year}-01-01 00:00:00')
        time_index = pd.date_range(start=start_date, periods=hours_in_year, freq='h')
        hourly_value = annual_consumption_kwh / hours_in_year
        return pd.Series([hourly_value] * hours_in_year, index=time_index, name='Verbrauch_kWh')

def load_standard_load_profile_with_weekdays(file_path, profile_type, annual_consumption_kwh, bundesland_code='BW', year=2024):
    """
    Lädt ein Standard-Lastprofil aus der Excel-Datei und berücksichtigt dabei die Wochentag-Sheets.
    Für jeden Tag wird der korrekte Wochentag-Typ (WT/SA/FT) bestimmt und die entsprechenden
    15-Minuten-Werte aus dem Lastprofil-Sheet entnommen.
    
    Args:
        file_path: Pfad zur Excel-Datei
        profile_type: Typ des Lastprofils ('H25', 'G25', 'L25', 'P25', 'S25')
        annual_consumption_kwh: Gewünschter Jahresverbrauch in kWh
        bundesland_code: Bundesland-Code für Feiertagsberechnung
        year: Jahr für die Wochentag-Berechnung
        
    Returns:
        pd.Series: Skaliertes Lastprofil mit 15-Minuten-Werten für ein Jahr
    """
    try:
        # 1. Lade das Wochentag-Sheet
        weekday_sheet_name = f'Wochentage_{year}'
        try:
            df_weekdays = pd.read_excel(file_path, sheet_name=weekday_sheet_name)

        except Exception as e:

            # Fallback: Verwende die alte Methode
            return load_standard_load_profile(file_path, profile_type, annual_consumption_kwh, bundesland_code, year)
        
        # 2. Lade das Lastprofil-Sheet ohne Header-Annahmen
        df_profile_raw = pd.read_excel(file_path, sheet_name=profile_type, header=None)
        
        # 3. Analysiere die Struktur - Die WT/SA/FT-Bezeichnungen stehen in Zeile 3 (Index 3)
        # und die numerischen Werte in den Zeilen darunter
        header_row = df_profile_raw.iloc[3]  # Zeile mit [kWh] SA FT WT SA FT WT...
        
        # Finde die Spalten für jeden Tagestyp
        wt_columns = []
        sa_columns = []
        ft_columns = []
        
        for col_idx, cell_value in enumerate(header_row):
            cell_str = str(cell_value).upper().strip()
            if cell_str == 'WT':
                wt_columns.append(col_idx)
            elif cell_str == 'SA':
                sa_columns.append(col_idx)
            elif cell_str == 'FT':
                ft_columns.append(col_idx)
        
        # Falls keine Spalten gefunden, verwende Fallback
        if not wt_columns and not sa_columns and not ft_columns:
            return load_standard_load_profile(file_path, profile_type, annual_consumption_kwh, bundesland_code, year)
        
        # 5. Erstelle Zeitindex für das ganze Jahr mit 15-Minuten-Intervallen
        is_leap_year = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        days_in_year = 366 if is_leap_year else 365
        intervals_per_day = 96  # 24h * 4 15-min-Intervalle
        total_intervals = days_in_year * intervals_per_day
        
        start_date = pd.Timestamp(f'{year}-01-01 00:00:00')
        time_index_15min = pd.date_range(start=start_date, periods=total_intervals, freq='15min')
        
        # 6. Neue korrekte Zuordnung basierend auf Datum, Monat und Tagestyp für 15-Min-Intervalle
        fifteen_min_values = []
        
        for interval_idx, timestamp in enumerate(time_index_15min):
            # Bestimme den Tag des Jahres (1-365/366)
            day_of_year = timestamp.timetuple().tm_yday
            
            # Hole den Wochentag-Typ aus dem Wochentag-Sheet
            if day_of_year <= len(df_weekdays):
                day_row = df_weekdays.iloc[day_of_year - 1]  # 0-basiert
                day_type = day_row[f'{bundesland_code}_Tagestyp']
            else:
                # Fallback für ungültige Tage
                day_type = 'WT'
            
            # Bestimme Monat (1-12), Stunde (0-23) und 15-Min-Intervall (0-3)
            month = timestamp.month  # 1-12
            hour_of_day = timestamp.hour  # 0-23
            minute_interval = timestamp.minute // 15  # 0, 15, 30, 45 -> 0, 1, 2, 3
            
            # Finde die richtige Spalte basierend auf Monat und Tagestyp
            month_base_col = 2 + (month - 1) * 3
            
            if day_type == 'SA':
                col_idx = month_base_col + 0  # SA ist erste Spalte des Monats
            elif day_type == 'FT':
                col_idx = month_base_col + 1  # FT ist zweite Spalte des Monats
            elif day_type == 'WT':
                col_idx = month_base_col + 2  # WT ist dritte Spalte des Monats
            else:
                col_idx = month_base_col + 2  # Fallback: WT
            
            # Finde die richtige Zeile basierend auf der Stunde und dem 15-Min-Intervall
            base_row = 4 + hour_of_day * 4  # Basis-Zeile für die Stunde
            row_idx = base_row + minute_interval  # Spezifische Zeile für das 15-Min-Intervall
            
            # Extrahiere den 15-Minuten-Wert
            fifteen_min_value = 0.0
            if row_idx < len(df_profile_raw) and col_idx < len(df_profile_raw.columns):
                try:
                    cell_value = df_profile_raw.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        numeric_value = float(cell_value)
                        if not pd.isna(numeric_value):
                            fifteen_min_value = numeric_value
                except (ValueError, TypeError, IndexError):
                    fifteen_min_value = 0.0
            
            # Korrekte Normierung: Verwende IMMER 1.000.000 kWh als Normierungsbasis (wie in der Excel)
            # Das ist die Standard-Methode für Lastprofile, auch wenn die tatsächliche Summe abweicht
            normalized_value = (fifteen_min_value / 1_000_000.0) * annual_consumption_kwh
            
            fifteen_min_values.append(normalized_value)
            

        
        # 7. Nachskalierung für exakte Energiebilanz
        # Problem: Excel-Profile sind nicht exakt auf 1.000.000 kWh normiert
        # Lösung: Skaliere auf die tatsächliche Summe
        current_sum = sum(fifteen_min_values)
        if current_sum > 0:
            scaling_factor = annual_consumption_kwh / current_sum
            fifteen_min_values = [v * scaling_factor for v in fifteen_min_values]
        else:
            scaling_factor = 1.0
        
        # 8. Erstelle die resultierende Series mit 15-Minuten-Intervallen
        result_series = pd.Series(fifteen_min_values, index=time_index_15min, name='Verbrauch_kWh')
        
        # 9. Validiere Energiebilanz nach Skalierung (Qualitätssicherung)
        actual_annual = result_series.sum()
        expected_annual = annual_consumption_kwh
        deviation_kwh = actual_annual - expected_annual
        deviation_percent = abs(deviation_kwh) / expected_annual * 100 if expected_annual > 0 else 0
        
        # Toleranz: 0.1% (Rundungsfehler sind OK)
        TOLERANCE_PERCENT = 0.1
        
        if deviation_percent > TOLERANCE_PERCENT:
            print(f"⚠️ WARNUNG: Lastprofil-Energiebilanz-Abweichung!")
            print(f"  Profil: {profile_type}")
            print(f"  Jahr: {year}")
            print(f"  Erwartet: {expected_annual:.2f} kWh")
            print(f"  Tatsächlich: {actual_annual:.2f} kWh")
            print(f"  Differenz: {deviation_kwh:+.2f} kWh ({deviation_percent:.3f}%)")
            print(f"  Toleranz: {TOLERANCE_PERCENT}%")
        else:
            # Debug-Ausgabe für erfolgreiche Validierung (kann auskommentiert werden)
            print(f"✅ Lastprofil-Energiebilanz OK: {profile_type}, {deviation_percent:.4f}% Abweichung")
        
        return result_series
        
    except Exception as e:
        print(f"Fehler beim Laden des Lastprofils {profile_type} mit Wochentagen: {e}")
        import traceback
        traceback.print_exc()
        # Fallback: Verwende die alte Methode
        return load_standard_load_profile(file_path, profile_type, annual_consumption_kwh, bundesland_code, year)

def scale_consumption_by_persons(consumption_series: pd.Series, number_of_persons: int, reference_persons: int = 1) -> pd.Series:
    """
    Skaliert die Verbrauchsdaten basierend auf der Anzahl der Haushalte.
    
    Args:
        consumption_series (pd.Series): Ursprüngliche Verbrauchsdaten (pro Haushalt)
        number_of_persons (int): Anzahl der Haushalte im Gebäude
        reference_persons (int): Referenz-Anzahl der Haushalte (Standard: 1)
    
    Returns:
        pd.Series: Skalierte Verbrauchsdaten
    """
    scaling_factor = number_of_persons / reference_persons
    scaled_consumption = consumption_series * scaling_factor
    return scaled_consumption

def scale_pv_generation(pv_generation_series: pd.Series, planned_pv_capacity_kwp: float, reference_pv_capacity_kwp: float = None) -> pd.Series:
    """
    Skaliert die PV-Erzeugungsdaten basierend auf der geplanten Anlagengröße in kWp.
    
    Args:
        pv_generation_series (pd.Series): Ursprüngliche PV-Erzeugungsdaten
        planned_pv_capacity_kwp (float): Geplante PV-Anlagengröße in kWp
        reference_pv_capacity_kwp (float): Referenz-Anlagengröße der ursprünglichen Daten (optional)
    
    Returns:
        pd.Series: Skalierte PV-Erzeugungsdaten
    """
    if reference_pv_capacity_kwp is None:
        # Wenn keine Referenz angegeben, nehmen wir die maximale Stundenerzeugung als Referenz
        max_hourly_generation = pv_generation_series.max()
        if max_hourly_generation > 0:
            # Schätzen der Referenz-Anlagengröße basierend auf maximaler Stundenerzeugung
            # Typische PV-Anlagen haben eine Spitzenleistung von etwa 80% der kWp-Leistung
            reference_pv_capacity_kwp = max_hourly_generation / 0.8
        else:
            # Fallback: Standard-Referenz von 10 kWp
            reference_pv_capacity_kwp = 10.0
    
    # Skalierungsfaktor berechnen
    scaling_factor = planned_pv_capacity_kwp / reference_pv_capacity_kwp
    
    # PV-Erzeugung skalieren
    scaled_pv_generation = pv_generation_series * scaling_factor
    
    return scaled_pv_generation

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
    full_time_index = pd.date_range(start=start_date, end=end_date, freq='h')

    df_processed = pd.DataFrame(index=full_time_index)
    df_processed = df_processed.merge(df_verbrauch[['Verbrauch_kWh', 'PV_Erzeugung_kWh']], 
                                     left_index=True, right_index=True, how='left')

    # Fehlende Werte validieren und behandeln
    missing_consumption = df_processed['Verbrauch_kWh'].isna().sum()
    missing_pv = df_processed['PV_Erzeugung_kWh'].isna().sum()
    
    if missing_consumption > 0 or missing_pv > 0:
        print(f"Warnung: {missing_consumption} fehlende Verbrauchswerte und {missing_pv} fehlende PV-Werte gefunden.")
        print("Fehlende Werte werden mit 0 gefüllt. Bitte überprüfen Sie Ihre Daten.")
    
    # Fehlende Werte mit 0 füllen (konservativer Ansatz)
    df_processed['Verbrauch_kWh'] = df_processed['Verbrauch_kWh'].fillna(0) 
    df_processed['PV_Erzeugung_kWh'] = df_processed['PV_Erzeugung_kWh'].fillna(0)
    
    # Validierung: Negative Werte sollten nicht auftreten
    negative_consumption = (df_processed['Verbrauch_kWh'] < 0).sum()
    negative_pv = (df_processed['PV_Erzeugung_kWh'] < 0).sum()
    
    if negative_consumption > 0 or negative_pv > 0:
        print(f"Warnung: {negative_consumption} negative Verbrauchswerte und {negative_pv} negative PV-Werte gefunden.")
        print("Negative Werte werden auf 0 gesetzt.")
        df_processed['Verbrauch_kWh'] = df_processed['Verbrauch_kWh'].clip(lower=0)
        df_processed['PV_Erzeugung_kWh'] = df_processed['PV_Erzeugung_kWh'].clip(lower=0)

    # Flexibilität: PV-Erzeugungsprofil kann auch aus Verbrauchs- und Netzeinspeisungsdaten rekonstruiert werden
    # Beispiel: Wenn nur Verbrauch und Netzbezug/-einspeisung gegeben sind
    # if 'Netzbezug_kWh' in df_verbrauch.columns and 'Netzeinspeisung_kWh' in df_verbrauch.columns:
    #    df_processed['PV_Erzeugung_kWh'] = df_processed['Verbrauch_kWh'] + df_verbrauch['Netzeinspeisung_kWh'] - df_verbrauch['Netzbezug_kWh']
    #    df_processed['PV_Erzeugung_kWh'] = df_processed['PV_Erzeugung_kWh'].clip(lower=0) # PV-Erzeugung kann nicht negativ sein

    return parameters, df_processed['Verbrauch_kWh'], df_processed['PV_Erzeugung_kWh']

def load_pv_generation_profile(file_path, annual_pv_generation_kwh, selected_year=2024):
    """
    Lädt PV-Erzeugungsdaten aus der Excel-Datei oder erstellt ein einfaches Profil.
    
    Args:
        file_path (str): Pfad zur Excel-Datei
        annual_pv_generation_kwh (float): Jährliche PV-Erzeugung in kWh
        selected_year (int): Jahr für die Daten
    
    Returns:
        pd.Series: PV-Erzeugungsdaten mit 15-Minuten-Intervallen
    """
    try:
        # Versuche PV-Erzeugung Sheet zu laden
        df_pv = pd.read_excel(file_path, sheet_name='PV_Erzeugung')
        
        # TODO: Verarbeite PV-Daten aus der Excel-Struktur
        # Für jetzt: Fallback auf synthetisches Profil

        
    except Exception as e:
        # Wenn PV-Sheet nicht gefunden, verwende vorerst ein einfaches Profil
        pass
    
    # Erstelle einfaches PV-Profil für 15-Minuten-Intervalle
    is_leap_year = selected_year % 4 == 0 and (selected_year % 100 != 0 or selected_year % 400 == 0)
    days_in_year = 366 if is_leap_year else 365
    intervals_per_day = 96  # 24h * 4 15-min-Intervalle
    total_intervals = days_in_year * intervals_per_day
    
    start_date = pd.Timestamp(f'{selected_year}-01-01 00:00:00')
    time_index_15min = pd.date_range(start=start_date, periods=total_intervals, freq='15min')
    
    pv_values = []
    
    for timestamp in time_index_15min:
        # Bestimme Tag im Jahr (1-365/366)
        day_of_year = timestamp.timetuple().tm_yday
        
        # Bestimme Stunde des Tages (0-23)
        hour_of_day = timestamp.hour
        minute = timestamp.minute
        
        # Einfaches PV-Profil: Nur zwischen 6:00 und 18:00 Uhr
        if 6 <= hour_of_day <= 18:
            # Saisonale Variation (Sommer höher als Winter)
            seasonal_factor = 0.4 + 0.6 * (1 + np.cos(2 * np.pi * (day_of_year - 172) / 365)) / 2
            
            # Tagesverlauf mit Maximum um 12:00
            hour_angle = (hour_of_day + minute/60 - 12) * np.pi / 6
            daily_factor = max(0, np.cos(hour_angle))
            
            # Kombiniere Faktoren
            pv_factor = seasonal_factor * daily_factor
            
            # Skaliere auf gewünschten Jahresertrag (nur 12h pro Tag aktiv)
            pv_value = pv_factor * (annual_pv_generation_kwh / (total_intervals * 0.5))
        else:
            pv_value = 0.0
        
        pv_values.append(pv_value)
    
    # Erstelle Series
    pv_generation_series = pd.Series(pv_values, index=time_index_15min, name='PV_Erzeugung_kWh')
    
    # Skaliere auf exakten Jahresertrag
    actual_annual = pv_generation_series.sum()
    if actual_annual > 0:
        scaling_factor = annual_pv_generation_kwh / actual_annual
        pv_generation_series = pv_generation_series * scaling_factor
    

    
    return pv_generation_series

def _detect_custom_pv_csv_format(lines: list[str]) -> str:
    """
    Ermittelt anhand der ersten Zeilen eines CSVs das vermutete Format.
    """
    if not lines:
        return 'unknown'
    head = [line.strip().lstrip('\ufeff') for line in lines if line.strip()][:5]
    head_lower = [line.lower() for line in head]

    if any('timestamp' in line or 'zeit' in line for line in head_lower) and any(';' in original for original in head):
        return 'semicolon_timestamp'
    if any(line.startswith('test1_day') or 'daily report' in line for line in head_lower):
        return 'daily_report'
    if any('plant name' in line and 'time' in line for line in head_lower):
        return 'daily_report'
    return 'pvgis'


def _infer_interval_minutes(timestamps: pd.Series) -> int:
    """
    Bestimmt die typische Intervalllänge (in Minuten) anhand der Zeitstempel.
    """
    try:
        diffs = timestamps.sort_values().diff().dt.total_seconds().dropna()
        diffs = diffs[diffs > 0]
        if diffs.empty:
            return 15
        mode = diffs.mode()
        base_seconds = mode.iloc[0] if not mode.empty else diffs.median()
        base_minutes = max(1, int(round(base_seconds / 60.0)))
        return base_minutes
    except Exception:
        return 15


def _normalize_energy_series(series: pd.Series | None, base_minutes: int) -> pd.Series | None:
    """
    Entfernt Duplikate, sortiert und füllt Lücken im Zeitindex mit 0-Werten.
    """
    if series is None or len(series) == 0:
        return None
    series = series.dropna()
    if series.empty:
        return None
    series = series.groupby(series.index).sum().sort_index()
    if series.empty:
        return None

    if base_minutes and len(series) > 1:
        try:
            freq = f'{int(base_minutes)}min'
            full_index = pd.date_range(series.index.min(), series.index.max(), freq=freq)
            series = series.reindex(full_index, fill_value=0.0)
        except Exception as e:
            print(f"Warnung: Konnte Zeitreihe nicht auf konstante Frequenz bringen ({e})")

    series.name = 'PV_Erzeugung_kWh'
    return series


def _parse_semicolon_timestamp_csv(text_content: str) -> pd.Series | None:
    """
    Verarbeitet CSVs vom Typ 'timestamp;load' mit Dezimal-Komma.
    """
    try:
        df = pd.read_csv(io.StringIO(text_content), sep=';', decimal=',', engine='python')
    except Exception as e:
        print(f"CSV-Parser (Semikolon) fehlgeschlagen: {e}")
        return None

    if df is None or df.empty:
        return None

    timestamp_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if any(keyword in col_lower for keyword in ['timestamp', 'zeit', 'datum']):
            timestamp_col = col
            break
    if timestamp_col is None:
        timestamp_col = df.columns[0]

    value_cols = [col for col in df.columns if col != timestamp_col]
    if not value_cols:
        return None
    value_col = value_cols[0]

    df = df.dropna(subset=[timestamp_col])
    timestamps = pd.to_datetime(
        df[timestamp_col].astype(str).str.strip(),
        dayfirst=True,
        errors='coerce'
    )
    df = df.assign(timestamp=timestamps).dropna(subset=['timestamp'])
    if df.empty:
        return None

    values = pd.to_numeric(
        df[value_col].astype(str).str.replace(',', '.'),
        errors='coerce'
    ).fillna(0.0)

    interval_minutes = _infer_interval_minutes(df['timestamp'])
    interval_hours = max(interval_minutes, 1) / 60.0
    # Annahme: Werte sind Leistungen in kW -> auf kWh pro Intervall umrechnen
    energy = values * interval_hours
    raw_series = pd.Series(energy.values, index=df['timestamp'])
    normalized = _normalize_energy_series(raw_series, interval_minutes)
    if normalized is not None:
        print(f"CSV erkannt: Semikolon-Timestamps ({interval_minutes}-Minuten-Auflösung, {len(normalized)} Werte).")
    return normalized


def _parse_daily_report_csv(text_content: str) -> pd.Series | None:
    """
    Verarbeitet Daily-Report-CSV (z.B. 'Plant name,Time,...').
    """
    lines = [line for line in text_content.splitlines() if line.strip()]
    if not lines:
        return None

    # Erste Zeile kann ein Titel ohne Delimiter sein
    if ',' not in lines[0] and len(lines) > 1:
        csv_body = '\n'.join(lines[1:])
    else:
        csv_body = '\n'.join(lines)

    try:
        df = pd.read_csv(io.StringIO(csv_body))
    except Exception as e:
        print(f"CSV-Parser (Daily Report) fehlgeschlagen: {e}")
        return None

    if df is None or df.empty:
        return None

    df.columns = [str(col).strip() for col in df.columns]
    if 'Time' not in df.columns:
        print("Daily-Report-CSV ohne 'Time'-Spalte.")
        return None

    timestamps = pd.to_datetime(df['Time'], errors='coerce')
    df = df.assign(timestamp=timestamps).dropna(subset=['timestamp']).sort_values('timestamp')
    if df.empty:
        return None

    interval_minutes = _infer_interval_minutes(df['timestamp'])
    energy_series = None

    power_col = next((col for col in df.columns if 'power' in col.lower()), None)
    if power_col:
        power_values = pd.to_numeric(df[power_col], errors='coerce').fillna(0.0)
        energy_series = power_values * (max(interval_minutes, 1) / 60.0)

    if energy_series is None:
        cumulative_col = next(
            (col for col in df.columns if 'yield' in col.lower() or 'energy' in col.lower()),
            None
        )
        if cumulative_col:
            cumulative = pd.to_numeric(df[cumulative_col], errors='coerce').ffill().fillna(0.0)
            diffs = cumulative.diff().fillna(cumulative.iloc[0])
            resets = diffs < 0
            if resets.any():
                diffs.loc[resets] = cumulative.loc[resets]
            energy_series = diffs.clip(lower=0.0)

    if energy_series is None:
        print("Daily-Report-CSV enthält keine nutzbaren Leistungs- oder Energie-Spalten.")
        return None

    raw_series = pd.Series(energy_series.values, index=df['timestamp'])
    normalized = _normalize_energy_series(raw_series, interval_minutes)
    if normalized is not None:
        print(f"CSV erkannt: Daily Report ({interval_minutes}-Minuten-Auflösung, {len(normalized)} Werte).")
    return normalized


def _finalize_pv_series(series: pd.Series | None, selected_year: int, source_label: str):
    """
    Führt die vorhandene Interpolationslogik aus und hängt Meta-Informationen an.
    """
    if series is None or series.empty:
        return None
    interpolation_result = interpolate_time_series_to_15min(series, selected_year)
    if interpolation_result is None:
        return None

    pv_generation_series_15min = interpolation_result['interpolated_series']
    pv_generation_series_15min.attrs = {
        'original_peak_power': interpolation_result['original_peak_power'],
        'original_resolution': interpolation_result['original_resolution'],
        'interpolation_method': interpolation_result['interpolation_method'],
        'original_stats': interpolation_result['original_stats'],
        'interpolated_stats': interpolation_result['interpolated_stats'],
        'energy_difference_percent': interpolation_result['energy_difference_percent'],
        'source_format': source_label
    }
    print(f"PV-Daten ({source_label}) erfolgreich auf 15-Minuten-Basis normalisiert.")
    return pv_generation_series_15min


def load_pv_generation_from_csv(file_path, selected_year=2024):
    """
    Lädt PV-Erzeugungsdaten aus einer CSV-Datei (z.B. PVGIS-Format).
    
    Args:
        file_path (str): Pfad zur CSV-Datei mit PV-Erzeugungsdaten
        selected_year (int): Jahr für die Daten
    
    Returns:
        pd.Series: PV-Erzeugungsdaten mit 15-Minuten-Intervallen oder None bei Fehler
    """
    try:
        # Lade CSV-Datei (robust für PVGIS: Header finden, Kommentare/Metadaten überspringen)
        text_content = None
        if hasattr(file_path, 'read'):
            # Streamlit UploadedFile -> Bytes einlesen
            raw = file_path.read()
            # Nach Einlesen den Pointer zurücksetzen, damit spätere Zugriffe möglich sind
            try:
                file_path.seek(0)
            except Exception:
                pass
            for enc in ['utf-8', 'latin1']:
                try:
                    text_content = raw.decode(enc)
                    break
                except Exception:
                    continue
        else:
            # Dateipfad
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text_content = f.read()

        if text_content is None:
            raise ValueError('CSV konnte nicht gelesen werden')

        lines = text_content.splitlines()
        format_hint = _detect_custom_pv_csv_format(lines)

        if format_hint == 'semicolon_timestamp':
            normalized_series = _parse_semicolon_timestamp_csv(text_content)
            return _finalize_pv_series(normalized_series, selected_year, "timestamp;load CSV")

        if format_hint == 'daily_report':
            normalized_series = _parse_daily_report_csv(text_content)
            return _finalize_pv_series(normalized_series, selected_year, "daily report CSV")

        header_idx = None
        delimiter = ','
        for idx, line in enumerate(lines):
            lower_line = line.lower()
            if 'time' in lower_line and 'p' in lower_line and ',' in line:
                header_idx = idx
                delimiter = ','
                break
            if 'time' in lower_line and 'p' in lower_line and ';' in line:
                header_idx = idx
                delimiter = ';'
                break
        if header_idx is None:
            for idx in range(max(0, len(lines)-100), len(lines)):
                line = lines[idx]
                lower_line = line.lower()
                if 'time' in lower_line and 'p' in lower_line and (',' in line or ';' in line):
                    header_idx = idx
                    delimiter = ',' if ',' in line else ';'
                    break
        if header_idx is None:
            raise ValueError("Kein Header mit Spalten 'time' und 'P' gefunden")

        csv_body = '\n'.join(lines[header_idx:])
        df_pv = pd.read_csv(io.StringIO(csv_body), sep=delimiter, engine='python')
        print(f"CSV (PVGIS) geladen ab Zeile {header_idx+1}: {len(df_pv)} Zeilen, Spalten: {list(df_pv.columns)}")
        
        # Prüfe ob die erwarteten Spalten vorhanden sind
        expected_columns = ['time', 'P', 'G(i)', 'H_sun', 'T2m', 'WS10m', 'Int']
        missing_columns = [col for col in expected_columns if col not in df_pv.columns]
        
        if missing_columns:
            print(f"Warnung: Fehlende Spalten: {missing_columns}")
            print(f"Verfügbare Spalten: {list(df_pv.columns)}")
        
        # Verwende 'P' Spalte für PV-Erzeugung (in W)
        if 'P' not in df_pv.columns:
            raise ValueError("Spalte 'P' (PV-Erzeugung) nicht gefunden")
        
        # Konvertiere Zeitstempel
        if 'time' not in df_pv.columns:
            raise ValueError("Spalte 'time' (Zeitstempel) nicht gefunden")
        
        # Zeitstempel parsen (PVGIS Format: YYYYMMDD:HHMM)
        print(f"Erste Zeitstempel-Werte: {df_pv['time'].head().tolist()}")
        
        # Versuche zuerst das explizite PVGIS-Format
        df_pv['time'] = pd.to_datetime(df_pv['time'].astype(str), format='%Y%m%d:%H%M', errors='coerce')
        
        # Falls das nicht funktioniert, versuche automatische Erkennung
        if df_pv['time'].isna().mean() > 0.5:
            print("Explizites Format fehlgeschlagen, versuche automatische Erkennung...")
            df_pv['time'] = pd.to_datetime(df_pv['time'], errors='coerce')
        
        df_pv = df_pv.dropna(subset=['time'])
        
        if len(df_pv) == 0:
            raise ValueError("Keine gültigen Zeitstempel gefunden")
        
        print(f"Zeitstempel erfolgreich geparst: {len(df_pv)} gültige Einträge")
        print(f"Zeitraum: {df_pv['time'].min()} bis {df_pv['time'].max()}")
        
        # Erstelle Series der Leistung in Watt (unterstützt Dezimal-Komma)
        print(f"Erste P-Werte: {df_pv['P'].head().tolist()}")
        print(f"P-Spalte Datentyp: {df_pv['P'].dtype}")
        
        p_numeric = pd.to_numeric(df_pv['P'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        print(f"Erste numerische P-Werte: {p_numeric.head().tolist()}")
        print(f"P-Werte Statistiken: Min={p_numeric.min()}, Max={p_numeric.max()}, Summe={p_numeric.sum()}")
        
        p_series_w = pd.Series(p_numeric.values, index=df_pv['time'], name='P_W')

        # Sortieren und Duplikate mitteln
        p_series_w = p_series_w.sort_index()
        p_series_w = p_series_w.groupby(p_series_w.index).mean()

        # Konvertiere Zeitstempel auf exakte Stunden (entferne Minuten)
        # PVGIS hat Zeitstempel wie 20230101:0010, wir brauchen 20230101:0000
        p_series_w.index = p_series_w.index.floor('h')
        print(f"Zeitstempel nach floor('h'): {p_series_w.index[:5].tolist()}")
        
        # Gruppiere nach exakten Stunden und summiere
        hourly_pv_w = p_series_w.groupby(p_series_w.index).sum()
        print(f"Stündliche PV-Werte nach Gruppierung: {len(hourly_pv_w)} Stunden, Summe: {hourly_pv_w.sum():.2f} W")
        print(f"Erste stündliche Werte: {hourly_pv_w.head().tolist()}")
        print(f"Index der gruppierten Daten: {hourly_pv_w.index[:5].tolist()}")
        
        # Konvertiere zu kWh
        hourly_pv_kwh = (hourly_pv_w / 1000.0)
        print(f"Nach kWh-Konvertierung: {len(hourly_pv_kwh)} Stunden, Summe: {hourly_pv_kwh.sum():.2f} kWh")
        
        # Erstelle Excel-Datei im richtigen Format
        if hasattr(file_path, 'name'):
            # Streamlit UploadedFile
            excel_path = file_path.name.replace('.csv', '_converted.xlsx')
        else:
            # String path
            excel_path = str(file_path).replace('.csv', '_converted.xlsx')
        
        df_excel = pd.DataFrame({
            'Zeitstempel': hourly_pv_kwh.index,
            'PV_Erzeugung_kWh': hourly_pv_kwh.values
        })
        df_excel.to_excel(excel_path, sheet_name='PV_Erzeugung', index=False)
        print(f"Excel-Datei erstellt: {excel_path}")
        
        # Verwende die PV-Daten direkt, unabhängig vom Jahr
        print(f"PV-Daten Jahr: {hourly_pv_kwh.index[0].year}, Erwartetes Jahr: {selected_year}")
        print(f"PV-Daten werden direkt verwendet, unabhängig vom Jahr")
        
        # Stelle sicher, dass die Länge exakt 8760 Stunden entspricht
        if len(hourly_pv_kwh) > 8760:
            print(f"Schaltjahr erkannt ({len(hourly_pv_kwh)} Stunden), kürze auf 8760 Stunden")
            hourly_pv_kwh = hourly_pv_kwh.head(8760)
        elif len(hourly_pv_kwh) < 8760:
            print(f"Unvollständiges Jahr ({len(hourly_pv_kwh)} Stunden), fülle mit Nullen auf")
            # Erweitere um fehlende Stunden
            missing_hours = 8760 - len(hourly_pv_kwh)
            additional_index = pd.date_range(start=hourly_pv_kwh.index[-1] + pd.Timedelta(hours=1), 
                                           periods=missing_hours, freq='h')
            additional_series = pd.Series(0, index=additional_index)
            hourly_pv_kwh = pd.concat([hourly_pv_kwh, additional_series])
        
        print(f"PV-Daten final: {len(hourly_pv_kwh)} Stunden, Summe: {hourly_pv_kwh.sum():.2f} kWh")
        
        # Stelle sicher, dass die Länge exakt 8760 Stunden entspricht (normales Jahr)
        # Falls es ein Schaltjahr ist, entferne die zusätzlichen Stunden
        if len(hourly_pv_kwh) > 8760:
            print(f"Schaltjahr erkannt ({len(hourly_pv_kwh)} Stunden), kürze auf 8760 Stunden")
            hourly_pv_kwh = hourly_pv_kwh.head(8760)
        elif len(hourly_pv_kwh) < 8760:
            print(f"Unvollständiges Jahr ({len(hourly_pv_kwh)} Stunden), fülle mit Nullen auf")
            # Erweitere um fehlende Stunden
            missing_hours = 8760 - len(hourly_pv_kwh)
            additional_index = pd.date_range(start=hourly_pv_kwh.index[-1] + pd.Timedelta(hours=1), 
                                           periods=missing_hours, freq='h')
            additional_series = pd.Series(0, index=additional_index)
            hourly_pv_kwh = pd.concat([hourly_pv_kwh, additional_series])
        
        print(f"PV-Daten nach Stunden aggregiert: {len(hourly_pv_kwh)} Stunden, Jahresertrag: {hourly_pv_kwh.sum():.2f} kWh")
        
        # Verwende die neue allgemeine Interpolationsfunktion
        print("Interpoliere PV-Daten auf 15-Minuten-Intervalle...")
        
        interpolation_result = interpolate_time_series_to_15min(hourly_pv_kwh, selected_year)
        
        if interpolation_result is None:
            print("Fehler bei der Interpolation, verwende Fallback-Methode")
            # Fallback: Alte Methode
            pv_generation_series_15min = hourly_pv_kwh.resample('15min').ffill() / 4.0
        else:
            pv_generation_series_15min = interpolation_result['interpolated_series']
            # Speichere ALLE Interpolationsinformationen für spätere Verwendung
            pv_generation_series_15min.attrs = {
                'original_peak_power': interpolation_result['original_peak_power'],
                'original_resolution': interpolation_result['original_resolution'],
                'interpolation_method': interpolation_result['interpolation_method'],
                'energy_difference_percent': interpolation_result['energy_difference_percent'],
                'original_stats': interpolation_result['original_stats'],
                'interpolated_stats': interpolation_result['interpolated_stats']
            }
        
        print(f"15-Minuten-PV-Daten erstellt: {len(pv_generation_series_15min)} Datenpunkte, Jahresertrag: {pv_generation_series_15min.sum():.2f} kWh")

        return pv_generation_series_15min
        
    except Exception as e:
        print(f"Fehler beim Laden der PV-Erzeugungsdaten aus CSV: {e}")
        import traceback
        traceback.print_exc()
        return None

def load_pv_generation_from_excel(file_path, selected_year=2024):
    """
    Lädt PV-Erzeugungsdaten aus einer separaten Excel-Datei.
    
    Args:
        file_path (str): Pfad zur Excel-Datei mit PV-Erzeugungsdaten
        selected_year (int): Jahr für die Daten
    
    Returns:
        pd.Series: PV-Erzeugungsdaten mit 15-Minuten-Intervallen oder None bei Fehler
    """
    try:
        # Versuche verschiedene Sheet-Namen für PV-Erzeugung
        possible_sheet_names = ['PV_Erzeugung', 'PV-Erzeugung', 'PV_Generation', 'Erzeugung', 'Generation', 'PV']
        
        df_pv = None
        for sheet_name in possible_sheet_names:
            try:
                df_pv = pd.read_excel(file_path, sheet_name=sheet_name)
                print(f"PV-Daten erfolgreich aus Sheet '{sheet_name}' geladen")
                break
            except:
                continue
        
        if df_pv is None:
            # Versuche das erste Sheet zu laden
            df_pv = pd.read_excel(file_path, sheet_name=0)
            print("PV-Daten aus dem ersten Sheet geladen")
        
        # Verarbeite die PV-Daten
        # Erwarte Spalten wie 'Zeitstempel', 'Datum', 'Zeit' und 'PV_Erzeugung_kWh', 'Erzeugung', etc.
        
        # Finde die Zeitstempel-Spalte
        time_column = None
        for col in df_pv.columns:
            if any(keyword in str(col).lower() for keyword in ['zeit', 'time', 'datum', 'date', 'timestamp']):
                time_column = col
                break
        
        if time_column is None:
            print("Keine Zeitstempel-Spalte gefunden, verwende Index")
            # Erstelle einen Standard-Zeitindex
            is_leap_year = selected_year % 4 == 0 and (selected_year % 100 != 0 or selected_year % 400 == 0)
            days_in_year = 366 if is_leap_year else 365
            intervals_per_day = 96  # 24h * 4 15-min-Intervalle
            total_intervals = days_in_year * intervals_per_day
            
            start_date = pd.Timestamp(f'{selected_year}-01-01 00:00:00')
            time_index = pd.date_range(start=start_date, periods=min(len(df_pv), total_intervals), freq='15min')
        else:
            # Konvertiere Zeitstempel-Spalte
            df_pv[time_column] = pd.to_datetime(df_pv[time_column])
            time_index = df_pv[time_column]
        
        # Finde die PV-Erzeugungsspalte
        pv_column = None
        for col in df_pv.columns:
            if any(keyword in str(col).lower() for keyword in ['pv', 'erzeugung', 'generation', 'solar', 'kwh']):
                if col != time_column:  # Nicht die Zeitstempel-Spalte
                    pv_column = col
                    break
        
        if pv_column is None:
            # Nimm die erste numerische Spalte
            for col in df_pv.columns:
                if col != time_column and pd.api.types.is_numeric_dtype(df_pv[col]):
                    pv_column = col
                    break
        
        if pv_column is None:
            raise ValueError("Keine PV-Erzeugungsspalte gefunden")
        
        # Erstelle die Series
        pv_values = df_pv[pv_column].fillna(0)  # Fehlende Werte mit 0 füllen
        
        # Stelle sicher, dass die Werte numerisch sind
        pv_values = pd.to_numeric(pv_values, errors='coerce').fillna(0)
        
        # Erstelle die finale Series
        pv_generation_series = pd.Series(pv_values.values, index=time_index, name='PV_Erzeugung_kWh')
        
        # Stelle sicher, dass der Index ein DatetimeIndex ist
        if not isinstance(pv_generation_series.index, pd.DatetimeIndex):
            pv_generation_series.index = pd.to_datetime(pv_generation_series.index)
        
        # Prüfe, ob die Daten stündlich sind und interpoliere auf 15-Minuten-Intervalle
        if len(pv_generation_series) <= 8760:  # Stündliche Daten (max 8784 für Schaltjahr)
            print("Stündliche PV-Daten erkannt, interpoliere auf 15-Minuten-Intervalle...")
            
            # Erstelle vollständigen 15-Minuten-Zeitindex für das Jahr
            is_leap_year = selected_year % 4 == 0 and (selected_year % 100 != 0 or selected_year % 400 == 0)
            days_in_year = 366 if is_leap_year else 365
            intervals_per_day = 96  # 24h * 4 15-min-Intervalle
            total_intervals = days_in_year * intervals_per_day
            
            start_date = pd.Timestamp(f'{selected_year}-01-01 00:00:00')
            end_date = pd.Timestamp(f'{selected_year}-12-31 23:59:59')
            time_index_15min = pd.date_range(start=start_date, end=end_date, freq='15min')
            
            # Resample stündliche Daten auf 15-Minuten-Intervalle
            # Jeder stündliche Wert wird auf 4 15-Minuten-Intervalle verteilt
            pv_generation_series = pv_generation_series.resample('15min').ffill()
            
            # Stelle sicher, dass die Länge korrekt ist
            if len(pv_generation_series) != total_intervals:
                # Kürze oder erweitere auf die korrekte Länge
                if len(pv_generation_series) > total_intervals:
                    pv_generation_series = pv_generation_series.head(total_intervals)
                else:
                    # Erweitere mit Nullen
                    additional_periods = total_intervals - len(pv_generation_series)
                    additional_index = pd.date_range(
                        start=pv_generation_series.index[-1] + pd.Timedelta(minutes=15),
                        periods=additional_periods,
                        freq='15min'
                    )
                    additional_series = pd.Series(0, index=additional_index)
                    pv_generation_series = pd.concat([pv_generation_series, additional_series])
        
        print(f"PV-Erzeugungsdaten erfolgreich geladen: {len(pv_generation_series)} Datenpunkte, Jahresertrag: {pv_generation_series.sum():.2f} kWh")
        
        return pv_generation_series
        
    except Exception as e:
        print(f"Fehler beim Laden der PV-Erzeugungsdaten: {e}")
        return None

def load_consumption_from_excel(file_path, selected_year=2024, annual_consumption_kwh=3500, number_of_persons=1, bundesland_code='BW', profile_type='H25'):
    """
    Lädt Verbrauchsdaten aus einer separaten Excel-Datei und berechnet 15-Minuten-Werte basierend auf Lastprofil.
    
    Args:
        file_path (str): Pfad zur Excel-Datei mit Verbrauchsdaten
        selected_year (int): Jahr für die Daten
        annual_consumption_kwh (float): Jährlicher Verbrauch pro Haushalt
        number_of_persons (int): Anzahl der Haushalte
        bundesland_code (str): Bundesland-Code für Feiertagsberechnung
        profile_type (str): Typ des Lastprofils ('H25', 'G25', 'L25', 'P25', 'S25')
    
    Returns:
        pd.Series: Verbrauchsdaten mit 15-Minuten-Intervallen oder None bei Fehler
    """
    try:
        # Verwende die bestehende Lastprofil-Funktion für 15-Minuten-Werte
        from data_import import load_standard_load_profile_with_weekdays
        
        # Lade 15-Minuten-Werte basierend auf Lastprofil
        # WICHTIG: annual_consumption_kwh ist bereits der Gesamtverbrauch für alle Haushalte
        consumption_series = load_standard_load_profile_with_weekdays(
            file_path, 
            profile_type,  # Profiltyp aus UI-Auswahl
            annual_consumption_kwh,  # Bereits Gesamtverbrauch für alle Haushalte
            bundesland_code, 
            selected_year
        )
        
        print(f"15-Minuten-Verbrauchsdaten erfolgreich geladen: {len(consumption_series)} Datenpunkte, Jahresverbrauch: {consumption_series.sum():.2f} kWh")
        
        return consumption_series
        
    except Exception as e:
        print(f"Fehler beim Laden der Verbrauchsdaten: {e}")
        return None

def preprocess_data_with_standard_profile(file_path, profile_type, annual_consumption_kwh, bundesland_code='BW', selected_year=2024, pv_system_size_kwp=10.0):
    """
    Verarbeitet Daten mit einem Standard-Lastprofil und PV-Erzeugungsdaten.
    
    Args:
        file_path (str): Pfad zur Excel-Datei
        profile_type (str): Typ des Lastprofils ('H25', 'G25', 'L25', 'P25', 'S25')
        annual_consumption_kwh (float): Gewünschter Jahresverbrauch in kWh
        bundesland_code (str): Bundesland-Code für Feiertagsberechnung
        selected_year (int): Jahr für die Wochentag-Berechnung
        pv_system_size_kwp (float): PV-Anlagengröße in kWp für die Erzeugungsberechnung
    
    Returns:
        tuple: (parameters, consumption_series, pv_generation_series)
    """
    # Lade das Standard-Lastprofil mit Wochentag-Berücksichtigung (15-Minuten-Intervalle)
    consumption_series = load_standard_load_profile_with_weekdays(file_path, profile_type, annual_consumption_kwh, bundesland_code, selected_year)
    
    # Erstelle Standard-Parameter
    parameters = {
        'profile_type': profile_type,
        'annual_consumption_kwh': annual_consumption_kwh,
        'bundesland_code': bundesland_code,
        'selected_year': selected_year,
        'pv_system_size_kwp': pv_system_size_kwp
    }
    
    # Berechne erwartete jährliche PV-Erzeugung basierend auf Anlagengröße
    # Typische PV-Anlage in Deutschland: ca. 950-1100 kWh/kWp pro Jahr
    # Wir verwenden 1000 kWh/kWp als Durchschnitt
    annual_pv_generation_kwh = pv_system_size_kwp * 1000.0
    
    # Lade PV-Erzeugungsprofil
    pv_generation_series = load_pv_generation_profile(file_path, annual_pv_generation_kwh, selected_year)
    
    return parameters, consumption_series, pv_generation_series 

def interpolate_time_series_to_15min(time_series, target_year=2024, preserve_peak_power=True):
    """
    Allgemeine Interpolationsfunktion für Zeitreihen auf 15-Minuten-Intervalle.
    
    Args:
        time_series (pd.Series): Zeitreihe mit beliebiger Zeitauflösung
        target_year (int): Zieljahr für die Interpolation
        preserve_peak_power (bool): Ob die ursprüngliche Spitzenleistung für Auswertungen gespeichert werden soll
    
    Returns:
        dict: {
            'interpolated_series': pd.Series,  # Interpolierte 15-Min-Daten
            'original_peak_power': float,      # Ursprüngliche Spitzenleistung
            'original_resolution': str,        # Ursprüngliche Zeitauflösung
            'interpolation_method': str,       # Verwendete Interpolationsmethode
            'original_stats': dict,            # Alle ursprünglichen Statistiken
            'energy_difference_percent': float # Energiebilanz-Differenz
        }
    """
    try:
        inferred_minutes = None
        if isinstance(time_series.index, pd.DatetimeIndex) and len(time_series) > 1:
            diffs = time_series.index.to_series().sort_index().diff().dt.total_seconds().dropna()
            diffs = diffs[diffs > 0]
            if not diffs.empty:
                inferred_minutes = diffs.mode().iloc[0] / 60.0

        if inferred_minutes is not None:
            if inferred_minutes <= 15:
                original_resolution = '15min'
                intervals_per_hour = max(1, int(round(60 / max(inferred_minutes, 0.25))))
            elif inferred_minutes <= 30:
                original_resolution = '30min'
                intervals_per_hour = max(1, int(round(60 / inferred_minutes)))
            elif inferred_minutes <= 60:
                original_resolution = 'hourly'
                intervals_per_hour = max(1, int(round(60 / inferred_minutes)))
            else:
                original_resolution = 'custom'
                intervals_per_hour = max(1, 60 / inferred_minutes)
        else:
            # Fallback auf Längen-Heuristik
            if len(time_series) <= 8784:  # Maximal für Schaltjahr
                if len(time_series) <= 8760:  # Normales Jahr
                    original_resolution = 'hourly'
                    intervals_per_hour = 1
                else:
                    original_resolution = 'hourly_leap_year'
                    intervals_per_hour = 1
            else:
                avg_interval_hours = 8760 / len(time_series)
                if avg_interval_hours <= 0.25:
                    original_resolution = '15min'
                    intervals_per_hour = 4
                elif avg_interval_hours <= 0.5:
                    original_resolution = '30min'
                    intervals_per_hour = 2
                else:
                    original_resolution = 'custom'
                    intervals_per_hour = 1 / avg_interval_hours
        
        # Speichere ALLE ursprünglichen Statistiken vor der Interpolation
        original_stats = {
            'peak_power_kw': time_series.max(),
            'min_power_kw': time_series.min(),
            'avg_power_kw': time_series.mean(),
            'total_energy_kwh': time_series.sum(),
            'std_power_kw': time_series.std(),
            'median_power_kw': time_series.median(),
            'non_zero_count': (time_series > 0).sum(),
            'max_consecutive_zeros': (time_series == 0).astype(int).groupby((time_series != 0).astype(int).cumsum()).sum().max() if (time_series == 0).any() else 0
        }
        
        # Erstelle vollständigen 15-Minuten-Zeitindex für das Zieljahr
        is_leap_year = target_year % 4 == 0 and (target_year % 100 != 0 or target_year % 400 == 0)
        days_in_year = 366 if is_leap_year else 365
        total_intervals = days_in_year * 24 * 4  # 15-Minuten-Intervalle
        
        start_date = pd.Timestamp(f'{target_year}-01-01 00:00:00')
        end_date = pd.Timestamp(f'{target_year}-12-31 23:59:59')
        target_time_index = pd.date_range(start=start_date, end=end_date, freq='15min')
        
        # Wähle Interpolationsmethode basierend auf ursprünglicher Auflösung
        if original_resolution == 'hourly' or original_resolution == 'hourly_leap_year':
            # Stündliche Daten: Gleichmäßige Verteilung auf 15-Min-Intervalle
            # WICHTIG: Energie gleichmäßig verteilen, Spitzenleistung bleibt für Auswertungen erhalten
            interpolated_series = time_series.resample('15min').ffill() / 4.0  # 4 15-Min-Intervalle pro Stunde
            interpolation_method = f"Gleichmäßige Verteilung (÷4)"
            
            # Für Auswertungen: Spitzenleistung unverändert lassen
            # Die interpolierten Werte sind für Simulationen, aber die ursprüngliche Spitzenleistung
            # bleibt für technische Auswertungen relevant
            
        elif original_resolution == '30min':
            # 30-Min-Daten: Gleichmäßige Aufteilung auf zwei 15-Min-Intervalle
            interpolated_series = time_series.resample('15min').ffill() / 2.0
            interpolation_method = "Gleichmäßige Verteilung (30->15)"
            
        elif original_resolution == '15min':
            # Bereits 15-Min-Daten: Direkte Verwendung
            interpolated_series = time_series
            interpolation_method = "Direkte Verwendung (bereits 15-Min)"
            
        else:
            # Custom/andere Auflösungen: Lineare Interpolation
            interpolated_series = time_series.resample('15min').interpolate(method='linear')
            interpolation_method = f"Lineare Interpolation (Custom: {original_resolution})"
        
        # Stelle sicher, dass die Länge korrekt ist
        if len(interpolated_series) != total_intervals:
            if len(interpolated_series) > total_intervals:
                interpolated_series = interpolated_series.head(total_intervals)
            else:
                # Erweitere mit Nullen
                additional_periods = total_intervals - len(interpolated_series)
                additional_index = pd.date_range(
                    start=interpolated_series.index[-1] + pd.Timedelta(minutes=15),
                    periods=additional_periods,
                    freq='15min'
                )
                additional_series = pd.Series(0, index=additional_index)
                interpolated_series = pd.concat([interpolated_series, additional_series])
        
        # Validiere Energiebilanz
        original_energy = time_series.sum()
        interpolated_energy = interpolated_series.sum()
        energy_difference = abs(original_energy - interpolated_energy) / original_energy * 100
        
        # Berechne interpolierte Statistiken zum Vergleich
        interpolated_stats = {
            'peak_power_kw': interpolated_series.max(),
            'min_power_kw': interpolated_series.min(),
            'avg_power_kw': interpolated_series.mean(),
            'total_energy_kwh': interpolated_series.sum(),
            'std_power_kw': interpolated_series.std(),
            'median_power_kw': interpolated_series.median(),
            'non_zero_count': (interpolated_series > 0).sum(),
            'max_consecutive_zeros': (interpolated_series == 0).astype(int).groupby((interpolated_series != 0).astype(int).cumsum()).sum().max() if (interpolated_series == 0).any() else 0
        }
        
        print(f"Interpolation abgeschlossen:")
        print(f"  Ursprüngliche Auflösung: {original_resolution}")
        print(f"  Ursprüngliche Spitzenleistung: {original_stats['peak_power_kw']:.2f} kW")
        print(f"  Interpolierte Spitzenleistung: {interpolated_stats['peak_power_kw']:.2f} kW")
        print(f"  Energiebilanz-Differenz: {energy_difference:.2f}%")
        print(f"  Methode: {interpolation_method}")
        print(f"  Betroffene Werte durch Interpolation:")
        print(f"    - Spitzenleistung: {original_stats['peak_power_kw']:.2f} -> {interpolated_stats['peak_power_kw']:.2f} kW")
        print(f"    - Minimalleistung: {original_stats['min_power_kw']:.2f} -> {interpolated_stats['min_power_kw']:.2f} kW")
        print(f"    - Durchschnittsleistung: {original_stats['avg_power_kw']:.2f} -> {interpolated_stats['avg_power_kw']:.2f} kW")
        
        return {
            'interpolated_series': interpolated_series,
            'original_peak_power': original_stats['peak_power_kw'],
            'original_resolution': original_resolution,
            'interpolation_method': interpolation_method,
            'original_stats': original_stats,
            'interpolated_stats': interpolated_stats,
            'energy_difference_percent': energy_difference
        }
        
    except Exception as e:
        print(f"Fehler bei der Interpolation: {e}")
        return None

@st.cache_data(show_spinner=False)
def load_battery_cost_curve():
    """
    Lädt die Batteriespeicher-Kostenkurve aus der Excel-Datei.
    WICHTIG: Mit @st.cache_data gecacht, damit Excel-Datei nur einmal geladen wird
    (verhindert Freezing bei Parameter-Eingabe).
    
    Returns:
        dict: Dictionary mit {Kapazität_kWh: Kosten_EUR} Mapping oder None bei Fehler
    """
    from config import BATTERY_COST_EXCEL_PATH
    import os
    
    try:
        if not os.path.exists(BATTERY_COST_EXCEL_PATH):
            raise FileNotFoundError(BATTERY_COST_EXCEL_PATH)
        # Robust: Suche in allen Sheets nach Kapazitäts- und Kostenspalten
        xls = pd.ExcelFile(BATTERY_COST_EXCEL_PATH, engine="openpyxl")
        print(f"[Kostenkurve] Datei geöffnet: {BATTERY_COST_EXCEL_PATH}")
        print(f"[Kostenkurve] Gefundene Sheets: {xls.sheet_names}")
        cost_dict = {}
        for sheet in xls.sheet_names:
            try:
                temp = pd.read_excel(BATTERY_COST_EXCEL_PATH, sheet_name=sheet, header=0, engine="openpyxl")
                if temp is None or temp.empty:
                    continue
                print(f"[Kostenkurve] Prüfe Sheet '{sheet}' mit {len(temp.columns)} Spalten / {len(temp)} Zeilen")
                print(f"[Kostenkurve] Erste Spalten: {list(temp.columns)[:6]}")
                # Kandidaten finden
                def find_col(candidates: list) -> str | None:
                    for cand in candidates:
                        for col in temp.columns:
                            if cand in str(col).lower():
                                return col
                    return None
                cols_lower = {c: str(c).lower() for c in temp.columns}
                # deutsche Varianten ergänzen
                capacity_keys = ["capacity_kwh", "kapaz", "kapazität", "kapazitaet", "speicher", "kwh"]
                cost_keys = ["total_cost_eur", "gesamtkosten", "kosten", "price", "invest", "gesamt", "total"]
                has_capacity = any(k in v for c, v in cols_lower.items() for k in capacity_keys)
                has_cost = any(k in v for c, v in cols_lower.items() for k in cost_keys)
                if not (has_capacity and has_cost):
                    # versuche header tiefer – einige Dateien haben Kopfzeilenblöcke
                    for hdr in range(1, 80):
                        try:
                            temp2 = pd.read_excel(BATTERY_COST_EXCEL_PATH, sheet_name=sheet, header=hdr, engine="openpyxl")
                            cols_lower2 = {c: str(c).lower() for c in temp2.columns}
                            if any(k in v for c, v in cols_lower2.items() for k in capacity_keys) and \
                               any(k in v for c, v in cols_lower2.items() for k in cost_keys):
                                temp = temp2
                                cols_lower = cols_lower2
                                print(f"[Kostenkurve] Verwende Header-Zeile {hdr} in Sheet '{sheet}'")
                                break
                        except Exception:
                            continue
                # Spalten bestimmen
                capacity_col = None
                for k in capacity_keys:
                    capacity_col = next((col for col in temp.columns if k in str(col).lower()), capacity_col)
                cost_col = None
                # bevorzuge deutsche "gesamtkosten" falls vorhanden
                for k in ["gesamtkosten"] + cost_keys:
                    cost_col = next((col for col in temp.columns if k in str(col).lower()), cost_col)
                if capacity_col is None or cost_col is None:
                    print(f"[Kostenkurve] In Sheet '{sheet}' keine passenden Spalten über Keywords gefunden. Versuche Heuristik...")
                    # Heuristik: finde Kapazitätsspalte (viele Werte im Bereich 1..200) und Kostenspalte (hohe Euro-Werte)
                    # 1) Spalten vorbereiten
                    temp_numeric = temp.copy()
                    for c in temp_numeric.columns:
                        temp_numeric[c] = pd.to_numeric(temp_numeric[c], errors='coerce')
                    # 2) Kandidaten für Kapazität: numerische Spalten mit genügend Werten 1..200
                    cap_candidates = []
                    for c in temp_numeric.columns:
                        series = temp_numeric[c].dropna()
                        if series.empty:
                            continue
                        values_in_range = series[(series >= 1) & (series <= 200)]
                        # mind. 5 gültige Einträge im Bereich
                        if len(values_in_range) >= 5:
                            # optional: hoher Anteil ganzzahlig
                            int_like_ratio = (values_in_range.round().eq(values_in_range)).mean()
                            cap_candidates.append((c, len(values_in_range), int_like_ratio))
                    cap_candidates.sort(key=lambda x: (x[1], x[2]), reverse=True)
                    # 3) Kandidaten für Kosten: numerische Spalten mit größeren Werten
                    cost_candidates = []
                    for c in temp_numeric.columns:
                        series = temp_numeric[c].dropna()
                        if series.empty:
                            continue
                        if series.mean() > 100 and series.max() > 500:  # grobe Schwellen für Euro
                            cost_candidates.append((c, series.mean(), series.max()))
                    cost_candidates.sort(key=lambda x: (x[1], x[2]), reverse=True)
                    # 4) Auswahl treffen, wobei Kapazität ≠ Kosten
                    if cap_candidates and cost_candidates:
                        capacity_col = cap_candidates[0][0]
                        # wähle erste Kosten-Spalte, die nicht die Kapazität ist
                        cost_col = next((c for c, _, _ in cost_candidates if c != capacity_col), None)
                        if cost_col is None:
                            print(f"[Kostenkurve] Heuristik fand keine unterschiedliche Kostenspalte in '{sheet}'.")
                    if capacity_col is None or cost_col is None:
                        print(f"[Kostenkurve] In Sheet '{sheet}' keine passenden Spalten gefunden.")
                        continue
                print(f"[Kostenkurve] Erkannte Spalten in '{sheet}': Kapazität='{capacity_col}', Kosten='{cost_col}'")
                # Numerik robust erzwingen (unterstützt deutsche Kommas)
                def to_numeric_eu(s):
                    if s.dtype == object:
                        s = s.astype(str).str.replace("\xa0", "", regex=False).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
                    return pd.to_numeric(s, errors="coerce")
                temp[capacity_col] = to_numeric_eu(temp[capacity_col])
                temp[cost_col] = to_numeric_eu(temp[cost_col])
                df_clean = temp[[capacity_col, cost_col]].dropna()
                df_clean = df_clean[
                    (df_clean[capacity_col] >= 1) &
                    (df_clean[capacity_col] <= 200) &
                    (df_clean[cost_col] > 0)
                ]
                for _, row in df_clean.iterrows():
                    cap_int = int(round(row[capacity_col]))
                    if cap_int not in cost_dict:  # ersten Eintrag je Größe nehmen
                        cost_dict[cap_int] = float(row[cost_col])
            except Exception:
                continue
        if not cost_dict:
            raise ValueError("Keine Kostendaten gefunden (Capacity/Cost Spalten nicht erkannt). Prüfe Spaltennamen und Header-Zeilen.")
        print(f"Batteriekostenkurve erfolgreich geladen: {len(cost_dict)} Einträge (1-{max(cost_dict.keys())} kWh)")
        sample10 = cost_dict.get(10)
        sample37 = cost_dict.get(37)
        sample100 = cost_dict.get(100)
        print(f"Kostenbeispiele: 10 kWh = {sample10 if sample10 is not None else 'N/A'} €, 37 kWh = {sample37 if sample37 is not None else 'N/A'} €, 100 kWh = {sample100 if sample100 is not None else 'N/A'} €")
        return cost_dict
        
    except FileNotFoundError:
        print(f"FEHLER: Batteriekostenkurven-Datei nicht gefunden: {BATTERY_COST_EXCEL_PATH}")
        print("Die Anwendung kann ohne diese Datei nicht korrekt funktionieren.")
        return None
        
    except Exception as e:
        print(f"FEHLER beim Laden der Batteriekostenkurve: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_battery_cost(capacity_kwh: float, cost_curve_data: dict) -> float:
    """
    Ermittelt die Kosten für eine Batteriespeicherkapazität aus der Kostenkurve.
    
    Args:
        capacity_kwh (float): Gewünschte Batteriekapazität in kWh
        cost_curve_data (dict): Dictionary mit Kostenkurven-Daten
    
    Returns:
        float: Kosten in EUR
    
    Raises:
        ValueError: Wenn die Kapazität außerhalb des Bereichs 1-200 kWh liegt
        TypeError: Wenn cost_curve_data None oder ungültig ist
    """
    if cost_curve_data is None:
        raise TypeError("Batteriekostenkurve nicht verfügbar. Bitte laden Sie die Datei battery_cost_curve_offers.xlsx.")
    
    # Spezialfall: 0 kWh (kein Speicher)
    if capacity_kwh == 0:
        return 0.0
    
    # Runde auf ganze Zahl (da nur ganzzahlige Werte in der Tabelle)
    capacity_rounded = int(round(capacity_kwh))
    
    # Prüfe Bereich 1-200 kWh
    if capacity_rounded < 1:
        raise ValueError(f"Batteriekapazität {capacity_rounded} kWh ist zu klein. Minimum: 1 kWh")
    if capacity_rounded > 200:
        raise ValueError(f"Batteriekapazität {capacity_rounded} kWh ist zu groß. Maximum: 200 kWh")
    
    # Hole Kosten aus Dictionary
    if capacity_rounded not in cost_curve_data:
        raise ValueError(f"Keine Kostendaten für {capacity_rounded} kWh verfügbar")
    
    return cost_curve_data[capacity_rounded] 

@st.cache_data(show_spinner=False)
def load_battery_tech_params(xlsm_path: str = "Batteriespeicherkosten.xlsm") -> dict:
    """
    Lädt technische Parameter (max. Lade-/Entladeleistung) je Batteriespeichergröße aus einer Excel (xlsm).
    
    Erwartete Daten (robust gegen unterschiedliche Header):
      - Kapazität: Spaltenname enthält eines von ["Capacity_kWh", "Kapazität", "kWh", "Speichergröße"]
      - Max. Ladeleistung: Spaltenname enthält eines von ["Max_Charge_kW", "Ladeleistung", "charge", "kW_Charge"]
      - Max. Entladeleistung: Spaltenname enthält eines von ["Max_Discharge_kW", "Entladeleistung", "discharge", "kW_Discharge"]
      - Alternativ: C-Rate-Spalte für (Laden/Entladen): ["C_Rate", "C-Rate", "Crate"]
    
    Rückgabe:
      { int_kapazitaet_kwh: {"max_charge_kw": float, "max_discharge_kw": float} }
    """
    try:
        # Versuche, das relevante Blatt automatisch zu finden: suche das erste Blatt mit passenden Spalten
        xls = pd.ExcelFile(xlsm_path, engine="openpyxl")
        print(f"[TechParams] Datei geöffnet: {xlsm_path}")
        print(f"[TechParams] Gefundene Sheets: {xls.sheet_names}")
        df = None
        used_sheet = None
        used_header = None
        for sheet in xls.sheet_names:
            for hdr in range(0, 81):
                try:
                    temp = pd.read_excel(xlsm_path, sheet_name=sheet, header=hdr, engine="openpyxl")
                    if temp is None or temp.empty:
                        continue
                    cols_lower = {c: str(c).lower() for c in temp.columns}
                    has_capacity = any(k in v for c, v in cols_lower.items() for k in ["capacity_kwh", "kapaz", "kapazität", "kapazitaet", "speicher", "kwh"] if isinstance(v, str))
                    has_charge = any(k in v for c, v in cols_lower.items() for k in ["max_charge_kw", "max ladeleistung", "ladeleistung", "charge"] if isinstance(v, str))
                    has_discharge = any(k in v for c, v in cols_lower.items() for k in ["max_discharge_kw", "max entladeleistung", "entladeleistung", "discharge"] if isinstance(v, str))
                    has_c_rate = any(k in v for c, v in cols_lower.items() for k in ["c_rate", "c-rate", "crate"] if isinstance(v, str))
                    if has_capacity and (has_charge and has_discharge or has_c_rate):
                        df = temp
                        used_sheet = sheet
                        used_header = hdr
                        print(f"[TechParams] Verwende Sheet '{sheet}' mit Header-Zeile {hdr}")
                        break
                except Exception:
                    continue
            if df is not None:
                break
        if df is None:
            raise ValueError("Keine geeignete Tabelle in Batteriespeicherkosten.xlsm gefunden.")
        
        # Spalten identifizieren
        def find_col(candidates: list) -> str | None:
            for cand in candidates:
                for col in df.columns:
                    if cand in str(col).lower():
                        return col
            return None
        
        capacity_col = find_col(["capacity_kwh", "kapaz", "kapazität", "kapazitaet", "speicher", "kwh"])
        charge_col = find_col(["max_charge_kw", "max ladeleistung", "ladeleistung", "charge"])
        discharge_col = find_col(["max_discharge_kw", "max entladeleistung", "entladeleistung", "discharge"])
        crate_col = find_col(["c_rate", "c-rate", "crate"])
        print(f"[TechParams] Erkannte Spalten: Kapazität='{capacity_col}', Lade='{charge_col}', Entlade='{discharge_col}', C-Rate='{crate_col}'")
        
        if capacity_col is None:
            raise ValueError("Kapazitätsspalte nicht gefunden (Capacity_kWh/Kapazität/kWh).")
        
        # Numerik robust erzwingen (unterstützt deutsche Kommas)
        def to_numeric_eu(s):
            if s.dtype == object:
                s = s.astype(str).str.replace("\xa0", "", regex=False).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
            return pd.to_numeric(s, errors="coerce")
        df[capacity_col] = to_numeric_eu(df[capacity_col])
        if charge_col:
            df[charge_col] = to_numeric_eu(df[charge_col])
        if discharge_col:
            df[discharge_col] = to_numeric_eu(df[discharge_col])
        if crate_col:
            df[crate_col] = to_numeric_eu(df[crate_col])
        
        df_clean = df.dropna(subset=[capacity_col]).copy()
        # Begrenze auf sinnvolle Kapazitäten
        df_clean = df_clean[(df_clean[capacity_col] >= 1) & (df_clean[capacity_col] <= 200)]
        
        tech_map: dict[int, dict] = {}
        for _, row in df_clean.iterrows():
            cap_int = int(round(row[capacity_col]))
            max_charge = row[charge_col] if charge_col in df_clean.columns else None
            max_discharge = row[discharge_col] if discharge_col in df_clean.columns else None
            c_rate = row[crate_col] if crate_col in df_clean.columns else None
            
            # Falls nur C-Rate vorhanden oder einzelne Werte fehlen -> berechne aus C-Rate
            if (max_charge is None or pd.isna(max_charge)) and c_rate is not None and pd.notna(c_rate):
                max_charge = float(c_rate) * float(cap_int)
            if (max_discharge is None or pd.isna(max_discharge)) and c_rate is not None and pd.notna(c_rate):
                max_discharge = float(c_rate) * float(cap_int)
            
            # Nur aufnehmen, wenn wir mindestens einen Wert haben
            entry = {}
            if max_charge is not None and not pd.isna(max_charge):
                entry["max_charge_kw"] = float(max_charge)
            if max_discharge is not None and not pd.isna(max_discharge):
                entry["max_discharge_kw"] = float(max_discharge)
            
            if entry:
                # Falls Duplikat derselben Kapazität -> ersten Eintrag beibehalten
                if cap_int not in tech_map:
                    tech_map[cap_int] = entry
                else:
                    # Ergänze fehlende Felder
                    tech_map[cap_int].update({k: v for k, v in entry.items() if k not in tech_map[cap_int]})
        
        if not tech_map:
            raise ValueError("Keine gültigen technischen Parameter gefunden.")
        
        return tech_map
    except FileNotFoundError:
        print(f"FEHLER: Datei nicht gefunden: {xlsm_path}")
        return None
    except Exception as e:
        print(f"FEHLER beim Laden technischer Parameter: {e}")
        import traceback
        traceback.print_exc()
        return None
