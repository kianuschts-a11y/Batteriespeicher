@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

if not exist "ui.py" (
    echo [FEHLER] ui.py wurde nicht gefunden!
    echo Aktuelles Verzeichnis: %CD%
    timeout /t 10
    exit /b 1
)

if exist "venv\Scripts\activate.bat" (
    call "venv\Scripts\activate.bat"
    set "PYTHON_EXE=venv\Scripts\python.exe"
    if not exist "!PYTHON_EXE!" (
        set "PYTHON_EXE=python"
    )
) else (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [FEHLER] Python wurde nicht gefunden!
    echo.
    echo Bitte installieren Sie Python von: https://www.python.org/downloads/
    echo Achten Sie darauf, "Add Python to PATH" anzukreuzen!
    echo.
    timeout /t 15
    exit /b 1
)

for /f "tokens=*" %%i in ('"%PYTHON_EXE%" --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Python-Version: %PYTHON_VERSION%
echo.

"%PYTHON_EXE%" -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo [HINWEIS] Streamlit ist nicht installiert.
    echo Installiere erforderliche Pakete...
    echo.
    "%PYTHON_EXE%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [FEHLER] Installation fehlgeschlagen!
        echo Bitte fuehren Sie manuell aus:
        echo    python -m pip install streamlit
        echo.
        timeout /t 15
        exit /b 1
    )
    echo Installation erfolgreich!
    echo.
)

echo.
echo ========================================
echo   Starte Anwendung...
echo ========================================
echo.
echo Die Anwendung oeffnet sich automatisch im Browser.
echo.
echo [WICHTIG] Zum Beenden druecken Sie STRG+C
echo           oder schliessen Sie einfach dieses Fenster.
echo.
echo ========================================
echo.

REM Starte Streamlit im Headless-Modus (kein Browser-Autostart)
"%PYTHON_EXE%" -m streamlit run ui.py

echo.
echo.
echo ========================================
echo   Anwendung wurde beendet
echo ========================================
echo.
echo Das Fenster schliesst sich in 5 Sekunden...
timeout /t 5
exit
