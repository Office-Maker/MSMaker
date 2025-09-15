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
    echo Selected Language Pack: %lang%
    goto NOLANGUAGEERROR
)

start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ %RED%[X] Error: Invalid language code%RESET%                                                                    │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Please enter a valid language code (e.g.: en-us, en-gb, de-de)                                      │
echo         │ If you are unsure, please open the readme file and check if your language code is supported.        │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
echo.
goto RETRYLANGUAGE
:NOLANGUAGEERROR


REM App selection
set "apps=Word,Excel,PowerPoint,OneNote,Access,Publisher,Outlook,OutlookForWindows,OneDrive,Groove,Teams,Lync"
set "names=Word,Excel,PowerPoint,OneNote,Access,Publisher,Outlook (classic),Outlook (new),OneDrive,OneDrive for Business,Microsoft Teams,Skype for Business"

set "includeApps="
set basicAppsSelected=0
call :getLength names len
for /L %%i in (0,1,%len%) do (
    call :getElement names %%i currentName
    call :getElement apps %%i currentApp

    REM Check if this is a basic office app

    CHOICE /C YN /N /M "Include !currentName! (Y/N)?"
    if !errorlevel!==1 (
        set "includeApps=!includeApps! !currentApp!"
        if /I "!currentApp!"=="Word" set basicAppsSelected=1
        if /I "!currentApp!"=="Excel" set basicAppsSelected=1
        if /I "!currentApp!"=="PowerPoint" set basicAppsSelected=1
        if /I "!currentApp!"=="OneNote" set basicAppsSelected=1
    )
)

REM Validate at least one basic app is selected
if !basicAppsSelected!==1 (goto NOSELECTIONERROR)
start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ %RED%[X] Error: Invalid program selection%RESET%                                                                │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ You must select at least one basic Office program (Word, Excel, PowerPoint, OneNote)^^!               │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
echo.
goto RETRYLANGUAGE
:NOSELECTIONERROR



REM Generate XML based on selected options
(
    echo ^<Configuration ID="bce5cf30-f4b8-4dba-91b3-4b4c9cae26f0"^>
    echo     ^<Add OfficeClientEdition="64" Channel="Current"^>
    echo         ^<Product ID="O365ProPlusRetail"^>
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
    echo     ^<Updates Enabled="TRUE" /^>
    echo     ^<RemoveMSI /^>
    echo     ^<AppSettings^>
    echo         ^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>
    echo         ^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>
    echo         ^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>
    echo     ^</AppSettings^>
    echo ^</Configuration^>
) > "%xmlFile%"

echo config file written: %xmlFile%
REM option to save configuration
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] Successfully written config file                                                                │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ (Y) Continue                                                                                        │
echo         │ (N) Cancel and delete configuration                                                                 │
echo         │ (S) Continue and save configuration                                                                 │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Enter a option to continue (Y/N/S)%RESET%                                                                  │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
choice /c YNS /n
if /i %errorlevel%==2 (echo cancel > %xmlFile%)
if /i %errorlevel%==3 (
    copy assets\config-temp.xml assets\save.xml >nul
    echo saved configuration as: assets\save.xml
)

exit /b

:getElement
setlocal enabledelayedexpansion
set "list=!%1!"
set /a idx=%2
set /a count=0

:loop
if "!list!"=="" (
    endlocal & set "%3=" & goto :eof
)

for /f "tokens=1* delims=," %%a in ("!list!") do (
    if !count! equ %idx% (
        endlocal & set "%3=%%a" & goto :eof
    )
    set "list=%%b"
    set /a count+=1
)
goto loop

:getLength
setlocal enabledelayedexpansion
set "list=!%1!"
set /a count=0

:countLoop
if "!list!"=="" (
    rem If count is 0 (empty list), return -1, else count-1
    if %count% equ 0 (
        set /a count=-1
    ) else (
        set /a count-=1
    )
    endlocal & set /a "%2=%count%-1" & goto :eof
)

for /f "tokens=1* delims=," %%a in ("!list!") do (
    set "list=%%b"
    set /a count+=1
)
goto countLoop