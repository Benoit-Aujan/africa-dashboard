@echo off
:: Pushes updated data.json to GitHub Pages so the live dashboard reflects today's data.
:: Called automatically by the Windows scheduled task after update_from_email.py runs.

cd /d "C:\Claude Projects\projects\africa-dashboard"

git add data/data.json
git diff --cached --quiet && (
    echo No data changes to publish.
    exit /b 0
)

for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set stamp=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%

git commit -m "Daily update %stamp%"
git push origin main

echo Published to GitHub Pages: %stamp%
