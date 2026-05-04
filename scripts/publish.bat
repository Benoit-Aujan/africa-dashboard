@echo off
:: publish.bat
:: 1. Push updated data.json to GitHub Pages.
:: 2. Wait for GitHub Pages deployment (~90 s).
:: 3. Take a fresh MTD screenshot → scripts/mtd_snapshot.png
::    (The Monday/Friday notification email reads this pre-saved file.)
::
:: Called automatically by the daily 12:00 PM scheduled task.

cd /d "C:\Claude Projects\projects\africa-dashboard"

:: ── Step 1: Push data.json if it changed ─────────────────────────────────────
git add data/data.json
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo No data changes to publish — skipping push.
    goto snapshot
)

for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set stamp=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%

git commit -m "Daily update %stamp%"
git push origin main
echo Published to GitHub Pages: %stamp%

:: ── Step 2: Wait for GitHub Pages deployment ─────────────────────────────────
echo Waiting 90s for GitHub Pages deployment...
timeout /t 90 /nobreak >nul

:: ── Step 3: Take dashboard snapshot ──────────────────────────────────────────
:snapshot
echo Taking dashboard snapshot...
python scripts\take_snapshot.py
