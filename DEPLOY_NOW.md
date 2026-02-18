# 🚀 App jetzt deployen - Schritt für Schritt

## ✅ Vorbereitung abgeschlossen!

Alle Dateien sind auf GitHub und bereit für das Deployment:
- ✅ Repository: `kianuschts-a11y/Batteriespeicher`
- ✅ Branch: `main`
- ✅ Alle Python-Dateien vorhanden
- ✅ `requirements.txt` vollständig
- ✅ `.streamlit/config.toml` konfiguriert

## 📋 Deployment-Schritte (5 Minuten):

### Schritt 1: Gehen Sie zu Streamlit Cloud
👉 [share.streamlit.io](https://share.streamlit.io)

### Schritt 2: Login
- Klicken Sie auf **"Sign in"**
- Melden Sie sich mit Ihrem **GitHub-Account** an

### Schritt 3: Neue App erstellen
1. Klicken Sie auf **"New app"** (oben rechts)
2. Füllen Sie die Felder aus:

   **Repository:**
   ```
   kianuschts-a11y/Batteriespeicher
   ```
   ⚠️ WICHTIG: Vollständiger Name mit `kianuschts-a11y/` und großem `B`!

   **Branch:**
   ```
   main
   ```

   **Main file path:**
   ```
   ui.py
   ```

   **App URL:** (optional - wird automatisch generiert)
   ```
   batteriespeicher-kalkulator
   ```
   Oder lassen Sie es leer für automatische Generierung.

### Schritt 4: Deploy starten
- Klicken Sie auf **"Deploy"** (unten rechts)
- ⏱️ Warten Sie 2-5 Minuten

### Schritt 5: App öffnen
- Nach erfolgreichem Deployment wird die App automatisch geöffnet
- Sie erhalten eine URL wie: `https://batteriespeicher-kalkulator.streamlit.app`

## 🔧 Falls die App bereits existiert:

### Option A: App neu starten
1. Gehen Sie zu Ihrer App
2. Klicken Sie auf **"Manage app"** (⚙️)
3. Klicken Sie auf **"Reboot app"**

### Option B: App-Einstellungen prüfen
1. Gehen Sie zu Ihrer App
2. Klicken Sie auf **"Settings"** (⚙️)
3. Prüfen Sie:
   - **Repository**: `kianuschts-a11y/Batteriespeicher`
   - **Branch**: `main`
   - **Main file path**: `ui.py`
4. Klicken Sie auf **"Save"**

## ❌ Häufige Probleme:

### Problem: "Repository not found"
**Lösung:** 
- Prüfen Sie, ob das Repository öffentlich ist
- Oder autorisieren Sie Streamlit in GitHub Settings → Applications

### Problem: "Main module does not exist"
**Lösung:**
- Stellen Sie sicher, dass `Main file path` = `ui.py` (nicht `./ui.py`)

### Problem: "Failed to clone"
**Lösung:**
- Prüfen Sie den Repository-Namen: Muss `kianuschts-a11y/Batteriespeicher` sein
- Prüfen Sie den Branch: Muss `main` sein

## 📊 Deployment-Status prüfen:

Nach dem Klick auf "Deploy" sehen Sie:
1. ✅ "Cloning repository..." (grün = erfolgreich)
2. ✅ "Installing dependencies..." (grün = erfolgreich)
3. ✅ "Starting app..." (grün = erfolgreich)
4. ✅ "App is running!" → App öffnet sich automatisch

## 🆘 Bei Fehlern:

1. Klicken Sie auf **"Manage app"** → **"Logs"**
2. Kopieren Sie die Fehlermeldungen
3. Prüfen Sie die häufigsten Fehler in `STREAMLIT_CLOUD_LOGS_HELP.md`

## ✅ Erfolg!

Wenn alles funktioniert, sehen Sie:
- ✅ Die App läuft
- ✅ Sie können Dateien hochladen
- ✅ Die Simulation funktioniert

**Ihre App-URL:** `https://[ihr-app-name].streamlit.app`

---

**Viel Erfolg! 🎉**
