@echo off
setlocal

cd /d "C:\Users\anzillotti\OneDrive - CGT Edilizia S.p.a\Documenti\GitHub\alayan_pv"

echo.
echo === Generazione sito Alayan PV ===
echo.

python "scripts\generate_site.py"
if errorlevel 1 (
    echo.
    echo Errore durante la generazione del sito.
    pause
    exit /b 1
)

echo.
echo Sito aggiornato in locale.
echo Ora esegui commit e push con GitHub Desktop.
pause
endlocal
