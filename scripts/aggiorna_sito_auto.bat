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

where git >nul 2>nul
if errorlevel 1 (
    echo.
    echo Git non trovato nel PATH.
    echo Il sito locale e' stato generato, ma commit e push automatici non sono disponibili.
    echo Installa Git for Windows oppure usa GitHub Desktop.
    pause
    exit /b 1
)

git add .
git diff --cached --quiet
if %errorlevel%==0 (
    echo.
    echo Nessuna modifica da pubblicare.
    exit /b 0
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set NOW=%%i

git commit -m "Aggiornamento dashboard PV %NOW%"
if errorlevel 1 (
    echo.
    echo Errore durante il commit.
    pause
    exit /b 1
)

git push origin main
if errorlevel 1 (
    echo.
    echo Errore durante il push su GitHub.
    pause
    exit /b 1
)

echo.
echo Aggiornamento completato con successo.
endlocal
