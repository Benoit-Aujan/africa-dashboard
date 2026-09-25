@echo off
:: publish.bat
:: 1. Push updated data.json to GitHub Pages.
:: 2. Take a fresh MTD screenshot → scripts/mtd_snapshot.png
::    (The Monday/Friday notification email reads this pre-saved file.)

cd /d "C:\Claude Projects\projects\africa-dashboard"

:: ── Step 1: Push data.json if it changed ─────────────────────────────────────
git add data/data.json
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo No data changes to publish.
    goto snapshot
)

:: Use Python for the date stamp — wmic parsing is unreliable in Task Scheduler
for /f %%d in ('"C:\Users\benoit.haas\AppData\Local\Programs\Python\Python314\python.exe" -c "import datetime; print(datetime.date.today().isoformat())"') do set stamp=%%d

git commit -m "Daily update %stamp%"

git push origin main
if %errorlevel% neq 0 (
    echo WARNING: git push failed, retrying in 15s...
    timeout /t 15 /nobreak >nul
    git push origin main
    if %errorlevel% neq 0 (
        echo ERROR: git push failed twice. Data committed but NOT published to GitHub Pages.
        exit /b 1
    )
)
echo Published to GitHub Pages: %stamp%

:: ── Step 2: Take dashboard snapshot (with deployment wait baked in) ───────────
:snapshot
echo Taking dashboard snapshot...
python scripts\take_snapshot.py --wait 90
