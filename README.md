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

## Technische Batterie-Parameter aus Excel
Die maximalen Lade-/Entladeleistungen werden je Batteriespeichergröße automatisch aus der Datei `Batteriespeicherkosten.xlsm` gelesen. Es werden folgende Spalten (robust) erkannt:
- Kapazität: z. B. `Capacity_kWh`, `Kapazität`, `Speichergröße`, `kWh`
- Max. Ladeleistung: z. B. `Max_Charge_kW`, `Ladeleistung`
- Max. Entladeleistung: z. B. `Max_Discharge_kW`, `Entladeleistung`
- Alternativ: `C_Rate`/`C-Rate` – dann wird kW = C-Rate × Kapazität_kWh berechnet

Falls für eine exakte Kapazität kein Eintrag existiert, wird der nächstliegende Wert verwendet. Die UI enthält keine manuellen Eingaben mehr für diese Leistungen.

## Nutzung

### Lokale Installation

1.  **Projektstruktur einrichten**:
    Erstellen Sie die folgenden Dateien in Ihrem Projektverzeichnis:
    *   `ui.py` (Hauptanwendung)
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

5.  **Interaktion im Browser**:
    Die Anwendung öffnet sich automatisch in Ihrem Standard-Webbrowser. Sie können nun Ihre Excel-Datei hochladen, Parameter in der Seitenleiste anpassen und verschiedene Simulationsszenarien auswählen und starten.

### Online-Deployment (Streamlit Cloud)

Die Anwendung kann auch online auf Streamlit Cloud betrieben werden:

1. **GitHub Repository**: Der Code ist bereits auf GitHub verfügbar
2. **Streamlit Cloud**: Folgen Sie der Anleitung in `STREAMLIT_CLOUD_SETUP.md`
3. **App-URL**: Nach dem Deployment erhalten Sie eine öffentliche URL für Ihre App

Siehe `STREAMLIT_CLOUD_SETUP.md` für detaillierte Anweisungen.

## Projektstruktur

```
Batteriespeicher Kalkulator/
├── config.py              # Zentrale Konfiguration und Konstanten
├── data_import.py         # Datenimport und Vorverarbeitung
├── model.py              # Kernsimulationslogik
├── scenarios.py          # Szenarien und Parametervariation
├── analysis.py           # Auswertung und Visualisierung
├── ui.py                 # Streamlit-Benutzeroberfläche
├── README.md             # Projektdokumentation
└── Anleitung/            # Anleitungsdateien
```

## Szenarien

Das Tool bietet vier verschiedene Simulationsszenarien:

1. **Einzel-Simulation**: Grundlegende Simulation mit festen Parametern
2. **Optimale Speichergröße**: Automatische Ermittlung der wirtschaftlich optimalen Batteriekapazität
3. **Haushaltstypen-Vergleich**: Vergleich verschiedener Haushaltstypen (EFH, MFH, Quartier)
4. **Variable Tarife**: Simulation mit stündlich variablen Strompreisen

## Visualisierungen

* **Zeitreihen-Analyse**: Detaillierte Darstellung der Energieflüsse über ausgewählte Zeiträume
* **Optimierungskurven**: Visualisierung der Abhängigkeit von KPIs von der Speichergröße
* **Szenarien-Vergleich**: Balkendiagramme zum Vergleich verschiedener Szenarien
* **Sankey-Diagramme**: Übersichtliche Darstellung der jährlichen Energieflüsse

## Weiterentwicklung

*   **Erweiterte Batteriesteuerung**: Implementierung von Optimierungsalgorithmen (z.B. lineare Programmierung) für die Batteriesteuerung unter variablen Stromtarifen.
*   **Kostenmodell**: Detaillierteres Kostenmodell für PV-Anlage, Speicher und Betrieb (z.B. Wartung, Degradation).
*   **Mehr Lastprofile**: Integration einer Bibliothek von Standardlastprofilen (SLP) oder die Möglichkeit, eigene Profile zu importieren.
*   **Wetterdaten**: Einbindung von Wetterdaten für genauere PV-Erzeugungsprognosen.
*   **Berichtserstellung**: Export der Ergebnisse und Grafiken in PDF-Berichte.

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz. Weitere Details finden Sie in der `LICENSE`-Datei. 