@echo off
REM ===================================================================
REM  Convertitore to PDF - avvio
REM  Al primo avvio crea un ambiente virtuale e installa le dipendenze.
REM  Agli avvii successivi parte subito.
REM ===================================================================
setlocal
cd /d "%~dp0"

REM Trova Python
where python >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non e' stato trovato nel PATH.
    echo Installa Python 3 da https://www.python.org/downloads/
    echo Ricordati di selezionare "Add python.exe to PATH".
    pause
    exit /b 1
)

REM Crea l'ambiente virtuale la prima volta
if not exist ".venv\Scripts\python.exe" (
    echo Primo avvio: preparazione dell'ambiente in corso...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERRORE] Impossibile creare l'ambiente virtuale.
        pause
        exit /b 1
    )
    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERRORE] Installazione delle dipendenze fallita.
        pause
        exit /b 1
    )
) else (
    call ".venv\Scripts\activate.bat"
)

REM Avvia l'applicazione
python convertitore.py

endlocal
