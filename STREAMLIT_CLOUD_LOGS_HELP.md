# Streamlit Cloud Logs - Häufige Fehler und Lösungen

## Wie Sie die Logs anzeigen:

1. Gehen Sie zu [share.streamlit.io](https://share.streamlit.io)
2. Klicken Sie auf Ihre App
3. Klicken Sie auf **"Manage app"** (oder das ⚙️-Symbol)
4. Klicken Sie auf **"Logs"** im Menü

## Häufige Fehler und Lösungen:

### 1. Import-Fehler: "ModuleNotFoundError"

**Fehler:**
```
ModuleNotFoundError: No module named 'xyz'
```

**Lösung:**
- Prüfen Sie `requirements.txt` - fehlt das Modul?
- Fügen Sie es hinzu und pushen Sie zu GitHub

### 2. Datei nicht gefunden: "FileNotFoundError"

**Fehler:**
```
FileNotFoundError: [Errno 2] No such file or directory: 'Batteriespeicherkosten.xlsm'
```

**Lösung:**
- Diese Datei muss von Benutzern hochgeladen werden
- Die App sollte auch ohne diese Datei starten (mit Warnung)
- Prüfen Sie, ob die Fehlerbehandlung korrekt ist

### 3. Syntax-Fehler

**Fehler:**
```
SyntaxError: invalid syntax
```

**Lösung:**
- Prüfen Sie die angegebene Zeile im Code
- Testen Sie lokal: `streamlit run ui.py`

### 4. Python-Version

**Fehler:**
```
Python version X.Y.Z is not supported
```

**Lösung:**
- In Streamlit Cloud Settings: Python-Version auf 3.11 oder 3.12 setzen
- Nicht 3.13 verwenden (kann Probleme haben)

### 5. Port bereits belegt

**Fehler:**
```
Address already in use
```

**Lösung:**
- Normalerweise kein Problem auf Streamlit Cloud
- Falls auftritt: App neu starten

## Debugging-Schritte:

1. **Lokale Tests:**
   ```bash
   streamlit run ui.py
   ```

2. **Requirements prüfen:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Python-Version prüfen:**
   - Streamlit Cloud unterstützt Python 3.8 - 3.12
   - Empfohlen: 3.11

4. **Logs kopieren:**
   - Kopieren Sie die Fehlermeldungen aus den Logs
   - Suchen Sie nach "ERROR" oder "Exception"

## Wenn die App nicht startet:

1. Prüfen Sie die ersten Zeilen der Logs (Startup)
2. Suchen Sie nach "ERROR" oder "Exception"
3. Prüfen Sie, ob alle Python-Dateien im Repository sind
4. Stellen Sie sicher, dass `ui.py` im Root-Verzeichnis liegt

## Nützliche Befehle für lokale Tests:

```bash
# App lokal starten
streamlit run ui.py

# Mit Debug-Ausgaben
streamlit run ui.py --logger.level=debug

# Port ändern (falls nötig)
streamlit run ui.py --server.port=8502
```
