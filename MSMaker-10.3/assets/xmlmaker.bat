@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM DIRECT RUN FILTER
set RED=!ESC![31m
set GREEN=!ESC![32m
set YELLOW=!ESC![33m
set BLUE=!ESC![34m
set LIGHTBLUE=!ESC![94m
set DARKER=!ESC![90m
set RESET=!ESC![0m!ESC![97m



:: Set the folder names to check for
set "folder1=System32"
set "folder2=assets"

:: Check if the current path ends with System32 or assets
for %%F in ("%CD%") do set "lastFolder=%%~nxF"
:: Check for System32
if /I "%lastFolder%"=="%folder1%" (
    exit /b
)
:: Check for assets
if /I "%lastFolder%"=="%folder2%" (
    exit /b
)


REM File path for the configuration XML file
set xmlFile=assets\config-temp.xml

REM Language prompt
:RETRYLANGUAGE
set "languages=ar-sa eu-es bg-bg ca-es zh-cn zh-tw hr-hr cs-cz da-dk nl-nl en-us en-gb et-ee fi-fi fr-fr gl-es de-de el-gr he-il hi-in hu-hu id-id it-it ja-jp kk-kz ko-kr lv-lv lt-lt ms-my nb-no pl-pl pt-br pt-pt ro-ro ru-ru sr-latn-rs sk-sk sl-si es-es sv-se th-th tr-tr uk-ua vi-vn"
set /p lang="%LIGHTBLUE%Enter a valid Language Code%RESET% (e.g.: en-us, en-gb, de-de): "

:: Check if input is in the list
set found=0
for %%L in (%languages%) do (
    if "%%L"=="%lang%" set found=1
)

if %found%==1 (
    echo Valid language code: %lang%
) else (
    echo Invalid language code: %lang%
    goto RETRYLANGUAGE
)


REM App selection loop
set "apps=Access Excel Lync OneDrive OneNote Outlook PowerPoint Publisher Word"
set "includeApps="

echo Select apps to include (press %LIGHTBLUE%[Y]%RESET% to include or %LIGHTBLUE%[N]%RESET% to exclude):
for %%a in (%apps%) do (
    CHOICE /C YN /N /M "Include %%a (Y/N)?"
    if !errorlevel!==1 (
        set "includeApps=!includeApps! %%a"
    )
)

REM Generate XML based on selected options
(
    echo ^<Configuration ID="0e76bc3c-228e-4640-b779-98f5ca61ff2e"^>
    echo     ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^>
    echo         ^<Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"^>
    echo             ^<Language ID="%lang%" /^>

    REM Exclude apps not selected
    for %%a in (%apps%) do (
        echo !includeApps! | findstr /i "%%a" >nul
        if errorlevel 1 (
            echo             ^<ExcludeApp ID="%%a" /^>
        )
    )

    echo         ^</Product^>
    echo     ^</Add^>
    echo     ^<Property Name="SharedComputerLicensing" Value="0" /^>
    echo     ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^>
    echo     ^<Property Name="DeviceBasedLicensing" Value="0" /^>
    echo     ^<Property Name="SCLCacheOverride" Value="0" /^>
    echo     ^<Property Name="AUTOACTIVATE" Value="1" /^>
    echo     ^<Updates Enabled="TRUE" /^>
    echo     ^<RemoveMSI /^>
    echo     ^<AppSettings^>
    echo         ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>
    echo         ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>
    echo         ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>
    echo     ^</AppSettings^>
    echo ^</Configuration^>
) > "%xmlFile%"

echo config file successfully written: %xmlFile%
REM option to save configuration
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] Successfully written config file                                                                │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ (Y) Continue with installation                                                                      │
echo         │ (N) Cancel and delete configuration                                                                 │
echo         │ (S) Continue with installation and save configuration                                               │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Enter a option to continue (Y/N/S)%RESET%                                                                  │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
choice /c YNS /n
if /i %errorlevel%==2 (echo cancel > %xmlFile%)
if /i %errorlevel%==3 (
	copy assets\config-temp.xml assets\save.xml
	echo saved configuration as: assets\save.xml
)

