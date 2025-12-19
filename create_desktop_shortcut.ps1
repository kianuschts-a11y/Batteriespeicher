# PowerShell-Skript zum Erstellen eines Desktop-Shortcuts für die Batteriespeicher-Anwendung
# UTF-8 Encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Desktop-Shortcut erstellen" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Projektverzeichnis bestimmen (Verzeichnis in dem sich dieses Skript befindet)
$PROJECT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $PROJECT_DIR

Write-Host "Projektverzeichnis: $PROJECT_DIR" -ForegroundColor Gray
Write-Host ""

# Prüfen ob die VBS-Datei existiert
$VBS_FILE = Join-Path $PROJECT_DIR "Batteriespeicher_App.vbs"
if (-not (Test-Path $VBS_FILE)) {
    Write-Host "[FEHLER] Batteriespeicher_App.vbs wurde nicht gefunden!" -ForegroundColor Red
    Write-Host "Pfad: $VBS_FILE" -ForegroundColor Gray
    Write-Host ""
    Read-Host "Drücken Sie Enter zum Beenden"
    exit 1
}

# Prüfen ob das Icon existiert
$ICON_FILE = Join-Path $PROJECT_DIR "lava_energy_icon.ico"
if (-not (Test-Path $ICON_FILE)) {
    Write-Host "[WARNUNG] lava_energy_icon.ico wurde nicht gefunden!" -ForegroundColor Yellow
    Write-Host "Der Shortcut wird ohne Icon erstellt." -ForegroundColor Gray
    Write-Host ""
    $ICON_FILE = $null
}

# Desktop-Pfad ermitteln
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $DesktopPath "Batteriespeicher Analyse Tool.lnk"

Write-Host "Shortcut-Ziel: $ShortcutPath" -ForegroundColor Gray
Write-Host ""

# Verknüpfung erstellen
try {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    
    # Ziel setzen (VBS-Datei)
    $Shortcut.TargetPath = $VBS_FILE
    
    # Arbeitsverzeichnis setzen
    $Shortcut.WorkingDirectory = $PROJECT_DIR
    
    # Beschreibung setzen
    $Shortcut.Description = "Batteriespeicher Analyse Tool - Startet die Anwendung ohne Terminal-Fenster"
    
    # Icon setzen (falls vorhanden)
    if ($ICON_FILE -and (Test-Path $ICON_FILE)) {
        $Shortcut.IconLocation = $ICON_FILE
        Write-Host "Icon gesetzt: $ICON_FILE" -ForegroundColor Green
    }
    
    # Verknüpfung speichern
    $Shortcut.Save()
    
    Write-Host ""
    Write-Host "[ERFOLG] Desktop-Shortcut wurde erfolgreich erstellt!" -ForegroundColor Green
    Write-Host "Pfad: $ShortcutPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Sie können die Anwendung jetzt über den Desktop-Shortcut starten." -ForegroundColor White
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "[FEHLER] Shortcut konnte nicht erstellt werden!" -ForegroundColor Red
    Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "Drücken Sie Enter zum Beenden"
    exit 1
}

# COM-Objekt freigeben
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Fertig!" -ForegroundColor Green
Write-Host ""
Read-Host "Drücken Sie Enter zum Beenden"

