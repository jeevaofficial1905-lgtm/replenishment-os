@echo off
setlocal
cd /d "%~dp0"
echo ===================================================
echo   ReplenishmentOS Data Pipeline
echo ===================================================

if "%~1"=="" (
    echo Processing latest file from Downloads...
    python process_pipeline.py "C:\Users\Jeeva\Downloads\pipeline as of 09-10-2026.xlsx" "C:\Users\Jeeva\Downloads\PIPELINE WORKING.xlsx"
) else (
    echo Processing dropped file: "%~1"...
    python process_pipeline.py "%~1" "C:\Users\Jeeva\Downloads\PIPELINE WORKING.xlsx"
)

echo.
echo ===================================================
echo Done! Upload 'pipeline_data.json' or 'PIPELINE WORKING.xlsx' in ReplenishmentOS.
echo ===================================================
pause
