@echo off
cd /d "%~dp0"

python src\whiskey_case_analysis.py
if errorlevel 1 (
  echo.
  echo Analysis failed.
  pause
  exit /b 1
)

echo.
echo Analysis complete. Opening the main chart...
start "" "exports\figures\recommended_mix.png"
pause
