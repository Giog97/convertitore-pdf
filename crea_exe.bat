@echo off
REM ===================================================================
REM  Crea il file eseguibile unico: dist\Convertitore PDF.exe
REM  Richiede che l'ambiente sia gia' stato preparato con avvia.bat.
REM ===================================================================
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERRORE] Ambiente non trovato. Avvia prima "avvia.bat" almeno una volta.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

REM Installa PyInstaller se manca
python -c "import PyInstaller" 2>nul
if errorlevel 1 pip install pyinstaller

echo.
echo Creazione dell'eseguibile in corso (puo' richiedere 1-2 minuti)...
echo.

python build_exe.py
if errorlevel 1 (
    echo.
    echo [ERRORE] Creazione dell'eseguibile fallita.
    pause
    exit /b 1
)

echo.
echo ===================================================================
echo  FATTO! Trovi il programma in:  dist\Convertitore PDF.exe
echo  Puoi copiare quel singolo file su una chiavetta o un altro PC.
echo ===================================================================
pause
endlocal
