# Streamlit Cloud Deployment - Anleitung

## ✅ Schritt 1: Code ist bereits auf GitHub hochgeladen

Der Code wurde erfolgreich zu GitHub gepusht:
- Repository: `https://github.com/kianuschts-a11y/Batteriespeicher.git`
- Branch: `main`

## 📋 Schritt 2: Streamlit Cloud Account erstellen

1. Gehen Sie zu [share.streamlit.io](https://share.streamlit.io)
2. Klicken Sie auf **"Sign up"** oder **"Sign in"**
3. Melden Sie sich mit Ihrem GitHub-Account an (empfohlen)

## 🚀 Schritt 3: App auf Streamlit Cloud deployen

1. Nach dem Login klicken Sie auf **"New app"**
2. Wählen Sie:
   - **Repository**: `kianuschts-a11y/Batteriespeicher`
   - **Branch**: `main`
   - **Main file path**: `ui.py`
   - **App URL**: (wird automatisch generiert, z.B. `batteriespeicher-kalkulator.streamlit.app`)

3. Klicken Sie auf **"Deploy"**

## ⏱️ Schritt 4: Warten auf Deployment

- Das Deployment dauert ca. 2-5 Minuten
- Sie sehen den Fortschritt in Echtzeit
- Bei Erfolg wird die App automatisch geöffnet

## 🔧 Schritt 5: Konfiguration (optional)

Falls Sie Anpassungen benötigen:

### Secrets/Umgebungsvariablen hinzufügen:
1. In Streamlit Cloud: **"Settings"** → **"Secrets"**
2. Fügen Sie bei Bedarf Konfigurationen hinzu (z.B. API-Keys)

### App-Einstellungen:
- Die App verwendet bereits die Konfiguration aus `.streamlit/config.toml`
- Theme und Server-Einstellungen sind vorkonfiguriert

## 📝 Wichtige Hinweise

### Datei-Uploads:
- Streamlit Cloud unterstützt Datei-Uploads
- Hochgeladene Dateien werden in der Session gespeichert
- Nach dem Neuladen der Seite sind die Dateien weg (das ist normal)

### Ressourcen:
- **Kostenloser Plan**: 
  - 1 App gleichzeitig
  - Begrenzte CPU/RAM
  - Für normale Nutzung ausreichend

### Updates:
- Bei jedem Push zu GitHub wird die App automatisch neu deployed
- Sie können auch manuell neu deployen: **"Manage app"** → **"Reboot app"**

## 🔗 Ihre App-URL

Nach erfolgreichem Deployment erhalten Sie eine URL wie:
```
https://batteriespeicher-kalkulator.streamlit.app
```

Diese URL können Sie mit anderen teilen!

## ❓ Troubleshooting

### App startet nicht:
- Prüfen Sie die Logs in Streamlit Cloud: **"Manage app"** → **"Logs"**
- Stellen Sie sicher, dass `requirements.txt` alle Dependencies enthält

### Fehler beim Import:
- Prüfen Sie, ob alle Python-Dateien im Repository sind
- Stellen Sie sicher, dass `ui.py` im Root-Verzeichnis liegt

### Dateien fehlen:
- Große Dateien (Excel, PDF) sind in `.gitignore` ausgeschlossen
- Diese müssen Benutzer selbst hochladen

## 📞 Support

Bei Problemen:
1. Prüfen Sie die Streamlit Cloud Logs
2. Testen Sie die App lokal: `streamlit run ui.py`
3. Prüfen Sie die GitHub Repository auf Fehler

---

**Viel Erfolg mit Ihrer App! 🎉**

