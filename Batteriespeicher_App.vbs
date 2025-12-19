Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Arbeitsverzeichnis setzen (Verzeichnis in dem sich diese .vbs-Datei befindet)
Dim scriptFolder
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = scriptFolder

' Pfad zur Batch-Datei erstellen
Dim batchFile
batchFile = fso.BuildPath(scriptFolder, "start_batteriespeicher.bat")

' Prüfen ob Batch-Datei existiert
If Not fso.FileExists(batchFile) Then
    MsgBox "FEHLER: Die Datei start_batteriespeicher.bat wurde nicht gefunden!" & vbCrLf & vbCrLf & "Pfad: " & batchFile, vbCritical, "Fehler"
    WScript.Quit
End If

' Streamlit-URL
Dim streamlitUrl
streamlitUrl = "http://localhost:8501"

' Batch-Datei im Hintergrund starten (0 = verstecktes Fenster)
' False = wartet nicht auf Beendigung (läuft parallel)
WshShell.Run """" & batchFile & """", 0, False

' Warten bis Streamlit-Server läuft (5 Sekunden)
WScript.Sleep 5000

' Browser im App-Modus öffnen
' Versuche zuerst Microsoft Edge (meist auf Windows 10/11 verfügbar)
Dim edgePath, chromePath
Dim browserFound
browserFound = False

' Edge-Pfade versuchen (verschiedene mögliche Installationen)
edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
If Not fso.FileExists(edgePath) Then
    edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
End If

If fso.FileExists(edgePath) Then
    ' Edge im App-Modus öffnen
    WshShell.Run """" & edgePath & """ --app=" & streamlitUrl & " --disable-extensions", 0, False
    browserFound = True
Else
    ' Fallback: Chrome versuchen
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    If Not fso.FileExists(chromePath) Then
        chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    End If
    
    If fso.FileExists(chromePath) Then
        ' Chrome im App-Modus öffnen
        WshShell.Run """" & chromePath & """ --app=" & streamlitUrl & " --disable-extensions", 0, False
        browserFound = True
    Else
        ' Fallback: Standard-Browser verwenden (nicht im App-Modus)
        WshShell.Run streamlitUrl, 1, False
        browserFound = True
    End If
End If

' Warten bis Streamlit-Server bereit ist (zusätzliche 2 Sekunden)
WScript.Sleep 2000

' Das VBScript läuft im Hintergrund weiter
' Die Batch-Datei (Streamlit-Server) läuft parallel
' Wenn der Benutzer die Anwendung beendet (über den "Schließen"-Button in der App),
' wird Streamlit beendet und damit auch die Batch-Datei
' Das VBScript endet automatisch, wenn die Batch-Datei beendet wird

' Einfache Überwachung: Warte auf ein Event (in der Praxis läuft es endlos im Hintergrund)
' Das VBScript bleibt aktiv, solange Streamlit läuft
On Error Resume Next
Do
    WScript.Sleep 10000
    ' Das Skript läuft einfach weiter, bis Streamlit beendet wird
Loop

