# config.py
import os

# Bestimme das Verzeichnis, in dem sich diese config.py Datei befindet
_CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))

# Standardwerte für Strompreise (falls nicht aus Datei geladen)
DEFAULT_PRICE_GRID_PER_KWH = 0.30 # Euro/kWh
DEFAULT_PRICE_FEED_IN_PER_KWH = 0.08 # Euro/kWh

# Standardwerte für Batterieeffizienz und Leistung
DEFAULT_BATTERY_EFFICIENCY_CHARGE = 0.95
DEFAULT_BATTERY_EFFICIENCY_DISCHARGE = 0.95
DEFAULT_BATTERY_MAX_CHARGE_KW = 5.0
DEFAULT_BATTERY_MAX_DISCHARGE_KW = 5.0
DEFAULT_INITIAL_SOC_PERCENT = 50.0 # Standard-Anfangsladezustand in %
DEFAULT_MIN_SOC_PERCENT = 10.0 # Minimaler Ladezustand in %
DEFAULT_MAX_SOC_PERCENT = 90.0 # Maximaler Ladezustand in %

# VEREINHEITLICHTE Standardwerte für alle Berechnungen
DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT = 1.0 # Jährlicher Kapazitätsverlust in % (vereinheitlicht)
DEFAULT_PROJECT_LIFETIME_YEARS = 15  # Projektlaufzeit in Jahren (vereinheitlicht)
DEFAULT_DISCOUNT_RATE = 0.02  # Abzinsungssatz (vereinheitlicht)
DEFAULT_PROJECT_INTEREST_RATE_DB = 0.03  # Zinssatz für Deckungsbeitragsrechnung

# Batteriekosten - Excel-Datei mit Kostenkurve
# Verwende absoluten Pfad basierend auf dem Verzeichnis der config.py
BATTERY_COST_EXCEL_PATH = os.path.join(_CONFIG_DIR, "Batteriespeicherkosten.xlsm")

# Alte Standardwerte für Batteriekosten (nicht mehr verwendet, nur als Fallback)
# DEFAULT_BATTERY_COST_PER_KWH = 500 # Euro/kWh
# DEFAULT_INSTALLATION_COST_FIXED = 2000 # Euro

# Legacy-Konstanten für Rückwärtskompatibilität (verwendet die vereinheitlichten Werte)
DEFAULT_PROJECT_LIFETIME_YEARS_DB = DEFAULT_PROJECT_LIFETIME_YEARS
DEFAULT_DISCOUNT_RATE_DB = DEFAULT_DISCOUNT_RATE
DEFAULT_BATTERY_DEGRADATION_RATE = DEFAULT_ANNUAL_CAPACITY_LOSS_PERCENT / 100.0

# Standard-Optimierungskriterium
DEFAULT_OPTIMIZATION_CRITERION = "Deckungsbeitrag III gesamt (Barwert)"

# Schwellenwerte für variable Tarife (Arbitrage-Steuerung)
# Diese Werte müssen an die spezifischen Preisstrukturen angepasst werden
THRESHOLD_BUY_PRICE = 0.15 # Euro/kWh: Wenn der Preis darunter liegt, aus Netz laden
THRESHOLD_SELL_PRICE = 0.40 # Euro/kWh: Wenn der Preis darüber liegt, ins Netz entladen
MAX_GRID_CHARGE_KW = 3.0 # Max. Leistung für Laden aus Netz
MAX_GRID_DISCHARGE_KW = 3.0 # Max. Leistung für Entladen ins Netz

# Weitere Konstanten können hier hinzugefügt werden
# Beispiel: Pfade zu Standard-Lastprofilen, etc. 