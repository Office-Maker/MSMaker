@echo off
setlocal enabledelayedexpansion

:: Define registry key for Office detection
set "clicktorun_key=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

:: Initialize detection flag
set "is_ltsc=0"

:: Check the Click-to-Run registry key for ProductReleaseIds
reg query "%clicktorun_key%" /v "ProductReleaseIds" >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=2*" %%A in ('reg query "%clicktorun_key%" /v "ProductReleaseIds"') do set "release_ids=%%B"

    :: Check for ProPlus2024Volume
    echo !release_ids! | find /i "ProPlus2024Volume" >nul
    if !errorlevel! equ 0 (
        set "is_ltsc=1"
        goto :end_check
    )
) else (
    :: ProductReleaseIds not found
    set "is_ltsc=0"
)

:end_check
:: Return the detection result
if "!is_ltsc!" equ "1" (
    exit /b 1
) else (
    exit /b 0
)
