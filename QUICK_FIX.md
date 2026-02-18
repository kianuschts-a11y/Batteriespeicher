# 🔧 Schnelle Fehlerbehebung für Streamlit Cloud

## 📋 Bitte teilen Sie die Fehlermeldungen aus den Logs:

1. Gehen Sie zu: [share.streamlit.io](https://share.streamlit.io)
2. Klicken Sie auf Ihre App → **"Manage app"** → **"Logs"**
3. Kopieren Sie die Fehlermeldungen (besonders Zeilen mit "ERROR" oder "Exception")

## 🔍 Häufigste Fehler und sofortige Lösungen:

### 1. ❌ ModuleNotFoundError
**Fehler:** `ModuleNotFoundError: No module named 'xyz'`

**Lösung:** 
- Prüfen Sie `requirements.txt` - ist das Modul enthalten?
- Falls nicht: Hinzufügen und zu GitHub pushen

### 2. ❌ FileNotFoundError  
**Fehler:** `FileNotFoundError: Batteriespeicherkosten.xlsm`

**Lösung:** 
- ✅ Bereits behoben - App sollte mit Warnung starten
- Falls nicht: Prüfen Sie die Fehlerbehandlung in `ui.py` Zeile 1228

### 3. ❌ SyntaxError
**Fehler:** `SyntaxError: invalid syntax at line X`

**Lösung:**
- Lokal testen: `streamlit run ui.py`
- Zeile X im Code prüfen

### 4. ❌ Python Version
**Fehler:** Python-Version Probleme

**Lösung:**
- In Streamlit Cloud Settings: Python auf **3.11** setzen (nicht 3.13!)

### 5. ❌ Import Error bei lokalen Modulen
**Fehler:** `cannot import name 'xyz' from 'config'`

**Lösung:**
- Prüfen Sie, ob alle Python-Dateien im Repository sind:
  - ✅ `config.py`
  - ✅ `settings_manager.py`
  - ✅ `excel_export.py`
  - ✅ `data_import.py`
  - ✅ `model.py`
  - ✅ `scenarios.py`
  - ✅ `analysis.py`

## 🚀 Schnelltest lokal:

```bash
# Testen Sie die App lokal
streamlit run ui.py
```

Wenn lokal alles funktioniert, sollte es auch auf Streamlit Cloud funktionieren.

## 📞 Nächste Schritte:

**Bitte teilen Sie:**
1. Die vollständige Fehlermeldung aus den Logs
2. Oder einen Screenshot der Logs
3. Dann kann ich gezielt helfen!
