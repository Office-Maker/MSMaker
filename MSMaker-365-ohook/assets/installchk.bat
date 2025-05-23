@echo off
setlocal enabledelayedexpansion

:: Define registry key for Office detection
set "clicktorun_key=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

:: Initialize detection flag
set "is_m365_enterprise=0"

:: Check the Click-to-Run registry key for ProductReleaseIds
reg query "%clicktorun_key%" /v "ProductReleaseIds" >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=2*" %%A in ('reg query "%clicktorun_key%" /v "ProductReleaseIds"') do set "release_ids=%%B"

    :: Check for O365ProPlusRetail or O365ProPlusEEANoTeamsRetail
    echo !release_ids! | find /i "O365ProPlusRetail" >nul
    if !errorlevel! equ 0 (
        set "is_m365_enterprise=1"
        goto :end_check
    )
    echo !release_ids! | find /i "O365ProPlusEEANoTeamsRetail" >nul
    if !errorlevel! equ 0 (
        set "is_m365_enterprise=1"
        goto :end_check
    )
) else (
    :: ProductReleaseIds not found
    set "is_m365_enterprise=0"
)

:end_check
:: Return the detection result
if "!is_m365_enterprise!" equ "1" (
    exit /b 1
) else (
    exit /b 0
)
