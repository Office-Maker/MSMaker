@echo off
chcp 65001 >nul  :: Force UTF-8
:ENTRY
cls
setlocal enabledelayedexpansion
cd %~dp0

REM elevator
for /f %%a in ('echo prompt $E ^| cmd') do set ESC=%%a
title MSMaker
color 0f
net session >nul 2>&1
if %errorLevel% neq 0 (
    goto ELEVATE
) else (goto PRESETTER)

:ELEVATE
echo Requesting administrative privileges...
powershell -Command "Start-Process cmd -ArgumentList '/c \"\"%~fnx0\"\"' -Verb RunAs"
if not %errorlevel%==0 (goto ELEVATEERROR)
exit

:ELEVATEERROR
cls
start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo  [X] ACCESS DENIED
echo ----------------------------------------------------------------------------------------------------------
echo  This script requires admin privileges to function correctly.
echo  Either you denied the admin priviliges or the script failed to start for some other reason,
echo  if the issue persists, please check if your powershell version is up to date.
echo ----------------------------------------------------------------------------------------------------------
echo  Press any key to try again or close this prompt
pause >nul
goto ENTRY



:PRESETTER
set RED=!ESC![31m
set GREEN=!ESC![32m
set YELLOW=!ESC![33m
set BLUE=!ESC![34m
set LIGHTBLUE=!ESC![94m
set DARKER=!ESC![90m
set RESET=!ESC![0m!ESC![97m


REM base integrity check
:baseintegritycheck
echo.
echo %DARKER%[PACKAGE VERIFICATION]%RESET%
set error=false
if not exist readme.txt (
	set error=true
	echo readme.txt %RED%[FAILED]%RESET%
) else (
	echo readme.txt %GREEN%[OK]%RESET%
)
if not exist assets\ (
    set error=true
	echo ^<DIR^> assets %RED%[FAILED]%RESET%
) else (
	echo assets %GREEN%[OK]%RESET%
)
if not exist assets\xmlmaker.bat (
    set error=true
	echo assets\xmlmaker.bat %RED%[FAILED]%RESET%
) else (
	echo assets\xmlmaker.bat %GREEN%[OK]%RESET%
)
if not exist assets\setup.exe (
    set error=true
	echo assets\setup.exe %RED%[FAILED]%RESET%
) else (
	echo assets\setup.exe %GREEN%[OK]%RESET%
)
if not exist assets\HWIDActivation.cmd (
    set error=true
	echo assets\HWIDActivation.cmd %RED%[FAILED]%RESET%
) else (
	echo assets\HWIDActivation.cmd %GREEN%[OK]%RESET%
)
if not exist assets\startmanager.bat (
    set error=true
	echo assets\startmanager.bat %RED%[FAILED]%RESET%
) else (
	echo assets\startmanager.bat %GREEN%[OK]%RESET%
)


if %error%==false (goto baseintegritycheck2)
start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [X] ERROR                                                                                           │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Woops, looks like other files required for this program are missing.                                │
echo         │ Note that you must extract the entire downloaded zip file into a folder to allow this script to     │
echo         │ access its essential files.                                                                         │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press any key to close this prompt%RESET%                                                                  │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto OUT


:baseintegritycheck2
set error=false
if not exist assets\config-DE.xml (
    set error=true
	echo assets\config-DE.xml %YELLOW%[FAILED]%RESET%
) else (
	echo assets\config-DE.xml %GREEN%[OK]%RESET%
)
if not exist assets\config-UK.xml (
    set error=true
	echo assets\config-UK.xml %YELLOW%[FAILED]%RESET%
) else (
	echo assets\config-UK.xml %GREEN%[OK]%RESET%
)
if not exist assets\config-US.xml (
    set error=true
	echo assets\config-US.xml %YELLOW%[FAILED]%RESET%
) else (
	echo assets\config-US.xml %GREEN%[OK]%RESET%
)

if %error%==true (
	echo.
	echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
	echo         │ [^^!] WARNING                                                                                         │
	echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
	echo         │ It looks like some configuration files are missing, however this will not affect the base           │
	echo         │ functionality of this program as the custom installation will still work.                           │
	echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
	echo         │ %LIGHTBLUE%Press any key to continue anyway%RESET%                                                                    │
	echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
	pause >nul
		goto DISCLAMER
) else (echo all files complete)




:DISCLAMER
set input=null
cls
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [^^!] DISCLAIMER                                                                                      │
echo         ├─────────────────────────────────────────────────────────────────────────────────┬───────────────────┤
echo         │ %RED%By running this script, you confirm that you have%RESET%                               │%YELLOW%         █         %RESET%│
echo         │ %RED%read the README file and agree to all terms and%RESET%                                 │%YELLOW%        ███        %RESET%│
echo         │ %RED%conditions stated within it^^! %RESET%                                                   │%YELLOW%       ██ ██       %RESET%│
echo         │ %RED%You accept full responsibility for any actions%RESET%                                  │%YELLOW%      ███ ███      %RESET%│
echo         │ %RED%performed by this script and any files included%RESET%                                 │%YELLOW%     ████ ████     %RESET%│
echo         │ %RED%in this package. The creator assumes no liability%RESET%                               │%YELLOW%    ███████████    %RESET%│
echo         │ %RED%for any damage, loss, or legal consequences^^! %RESET%                                   │%YELLOW%   ██████ ██████   %RESET%│
echo         │ %RED%Continue at your own risk^^! %RESET%                                                     │%YELLOW%  ███████████████  %RESET%│
echo         ├─────────────────────────────────────────────────────────────────────────────────┴───────────────────┤
echo         │ %LIGHTBLUE%(A) Agree and continue%RESET%                                                                              │
echo         │ %LIGHTBLUE%(R) Open readme file%RESET%                                                                                │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
choice /c AR /n
if %errorlevel%==1 (goto STARTMANAGER)
if %errorlevel%==2 (start readme.txt)
goto DISCLAMER



:STARTMANAGER
cls
call assets\startmanager.bat
if %errorlevel%==1 (goto INSTALLATIONRUNCHK)
if %errorlevel%==2 (goto ACTIVATEWIN)



:INSTALLATIONRUNCHK
cls
echo %DARKER%[CHECKING CONDITIONS]%RESET%
echo|set /p=searching for running office installer... 
tasklist /FI "IMAGENAME eq setup.exe" | find /I "setup.exe" >nul 2>&1
if errorlevel 1 (
	
    echo %GREEN%[NOT FOUND]%RESET%
    goto ALREADYINSTALLEDCHK
)
echo %RED%[FOUND]%RESET%
start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [X] ENTRY BLOCKED                                                                                   │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ An active office installer has been found running in the background, we cannot proceed as long the  │
echo         │ installation is not finished. Please do not kill the installer task, instead wait for the           │
echo         │ installation to finish.                                                                             │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press any key to try again%RESET%                                                                          │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto INSTALLATIONRUNCHK
:ALREADYINSTALLEDCHK
echo|set /p=searching for installed Office LTSC 2024... 
call assets\installchk.bat
set result=%errorlevel%
if not %result%==1 (
    echo %GREEN%[NOT FOUND]%RESET%
	goto PRESTART
)
echo %BLUE%[FOUND]%RESET%
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] Information                                                                                     │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ We already found an installed version of Office LTSC 2024, if you don't want to reinstall Office    │
echo         │ you can select 'More options' in the main menu and only run the activation process.                 │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press any key to continue to main menu%RESET%                                                              │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto PRESTART



:PRESTART
call :CLEANUP
:START
REM ////////////////////////////////////////////////////////////////////////////////////
set config_file=none.xml
REM ////////////////////////////////////////////////////////////////////////////////////
cd %~dp0
cls
title OfficeMaker
echo.
echo     %RED%████████  %GREEN%████████%RESET%     
echo     %RED%████████  %GREEN%████████%RESET%      .d88888b.   .d888  .d888 d8b
echo     %RED%████████  %GREEN%████████%RESET%     d88P" "Y88b d88P"  d88P"  Y8P
echo     %RED%████████  %GREEN%████████%RESET%     888     888 888    888
echo                            888     888 888888 888888 888  .d8888b .d88b.
echo     %BLUE%████████  %YELLOW%████████%RESET%     888     888 888    888    888 d88P"   d8P  Y8b
echo     %BLUE%████████  %YELLOW%████████%RESET%     888     888 888    888    888 888     88888888
echo     %BLUE%████████  %YELLOW%████████%RESET%     Y88b. .d88P 888    888    888 Y88b.   Y8b.
echo     %BLUE%████████  %YELLOW%████████%RESET%      "Y88888P"  888    888    888  "Y8888P "Y8888
echo.
echo     888b     d888        d8888 888    d8P  8888888888 8888888b.  
echo     8888b   d8888       d88888 888   d8P   888        888   Y88b 
echo     88888b.d88888      d88P888 888  d8P    888        888    888 
echo     888Y88888P888     d88P 888 888d88K     8888888    888   d88P 
echo     888 Y888P 888    d88P  888 8888888b    888        8888888P"  
echo     888  Y8P  888   d88P   888 888  Y88b   888        888 T88b   
echo     888   "   888  d8888888888 888   Y88b  888        888  T88b  
echo     888       888 d88P     888 888    Y88b 8888888888 888   T88b
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ Welcome to OfficeMaker, please select your installation of Office                                   │
echo         ├──────────────────────────────────────────┬──────────────────────────────────────────────────────────┤
echo         │ 0. Custom Office                         │ 3. Start Menu                                            │
echo         │ 1. Basic Office (German DE)              │ 4. More Options                                          │
echo         │ 2. Basic Office (English US)             │ 5. Open README                                           │
echo         ├──────────────────────────────────────────┴──────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press a number to select%RESET%                                                                            │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
choice /c 012345D /n
if %errorlevel%==1 (goto CUSTOM)
if %errorlevel%==2 (set config_file=config-DE.xml)
if %errorlevel%==3 (set config_file=config-US.xml)
if %errorlevel%==4 (goto STARTMANAGER)
if %errorlevel%==5 (goto MOREOPT)
if %errorlevel%==6 (
	start readme.txt
	goto START
)
if %errorlevel%==7 (goto DEV)


goto PROCESS
:MOREOPT
cls
echo.
echo         ╭─────────────────────────────┬───────────────────────────────────────────────────────────────────────╮
echo         │ More Options                │ Description                                                           │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ 1. Retry activation         │ Only executes activation process, use if activation failed before.    │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ 2. Install 'custom.xml'     │ Installs Office with 'custom.xml' configuration which you have to     │
echo         │                             │ import manually into the assets folder.                               │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ 3. Install 'save.xml'       │ Installs Office with the optionally saved 'Custom Office' version     │
echo         │                             │ You can save a configuration by selecting '0' in the main menu.       │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ 4. Retry temp-file removal  │ Cleans up unnecessary files left over in the assets folder.           │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ 0. Main menu                │ Closes this menu                                                      │
echo         ├─────────────────────────────┴───────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press a number to select%RESET%                                                                            │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
choice /c 12340 /n
if %errorlevel%==1 (
	goto REACTIVATE)
if %errorlevel%==2 (
	set config_file=custom.xml
	goto PROCESS)
if %errorlevel%==3 (
	set config_file=save.xml
	goto PROCESS)
if %errorlevel%==4 (
	goto RECLEANUP)
if %errorlevel%==5 (goto START)
goto MOREOPT

:REACTIVATE
echo.
echo %BLUE%[OFFICE ACTIVATION] Step 1/1%RESET%
call :ACTIVATION
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] DONE                                                                                            │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Tried to activate Office.                                                                           │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%You may now close this prompt. (Press any key to return to main menu)%RESET%                               │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto START
:RECLEANUP
echo.
echo %DARKER%[TEMP-FILE REMOVAL] Step 1/1%RESET%
call :CLEANUP
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] DONE                                                                                            │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Removed all unnecessary files from Office Maker.                                                    │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%You may now close this prompt. (Press any key to return to main menu)%RESET%                               │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto START



:PROCESS

REM ///FILE CHECKUP///
echo.
echo %RED%[FILE INTEGRITY CHECK] Step 1/5%RESET%
set error=false
if not exist assets\%config_file% (
	set error=true
	echo %config_file% %RED%[FAILED]%RESET%
) else (
	echo %config_file% %GREEN%[OK]%RESET%
)
if not exist assets\setup.exe (
    set error=true
	echo setup.exe %RED%[FAILED]%RESET%
) else (
	echo setup.exe %GREEN%[OK]%RESET%
)

if %error%==true (
	start /min cmd /c "assets\sounder.bat >nul 2>&1"
	echo.
	echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
	echo         │ [X] ERROR                                                                                           │
	echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
	echo         │ Whoops, looks like a necessary installation file is missing.                                        │
	echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
	echo         │ %LIGHTBLUE%Press any key to return to menu%RESET%                                                                     │
	echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
	pause >nul
		goto START
) else (echo all files complete)


REM ///REMOVAL///
echo.
echo %YELLOW%[OFFICE REMOVAL] Step 2/5%RESET%
echo Please remove any old installations of Microsoft Office, Microsoft 365 or
echo free standalone programs like OneNote which belong to Microsoft 365.
control appwiz.cpl
echo If you removed all Office installations, %LIGHTBLUE%press any key to continue.%RESET%
pause >nul

:CRITICALPROCESS
REM ///INSTALLATION///
cd assets
echo.
echo %GREEN%[OFFICE INSTALLATION] Step 3/5%RESET%
echo %YELLOW%WARNING: Do not close this script or power down your system! %RESET%
echo|set /p=running external installation wizard... 
call :StartTaskWithSpinner setup.exe /configure %config_file%

REM Checking if installation has been finished or has crashed
call installchk.bat
set result=%errorlevel%
if %result%==1 (
    echo %GREEN%[DONE]%RESET%
) else (
	echo %RED%[FAILED]%RESET%
    goto INSTALLFAILED
)


REM ///ACTIVATION///
echo.
echo %BLUE%[OFFICE ACTIVATION] Step 4/5%RESET%
call :ACTIVATION

REM ///FILE CLEANUP///
echo.
echo %DARKER%[OFFICE-MAKER RESET] Step 5/5%RESET%
call :CLEANUP

REM ///DONE///
:DONE
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] DONE                                                                                            │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Microsoft Office 2024 is successfully installed and should have been activated^^!                     │
echo         │ If your Office didn't activate the KMS server might be temporarily down, return to the main menu    │
echo         │ and select 'More options' to retry the activation process.                                          │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%You may now close this prompt. (Press any key to return to main menu)%RESET%                               │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
goto START

REM ///INSTALL FAILED///
:INSTALLFAILED
start /min cmd /c "assets\sounder.bat >nul 2>&1"
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [X] ERROR                                                                                           │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Whoops, looks like the installation got interrupted.                                                │
echo         │ Please check if office really was fully uninstalled before starting this installation.              │
echo         │ If you chose 'Custom Office', you may want to double-check your inputs.                             │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%Press any key to return to menu%RESET%                                                                     │
echo         ╰─────────────────────────────────────────────────────────────────────────────────────────────────────╯
pause >nul
call :CLEANUP
goto START












REM ///FILE CLEANUP function///
:CLEANUP
cd %~dp0\assets
echo deleting temp files...
del config-temp.xml
echo temp files deleted
exit /b

:ACTIVATION
REM ///ACTIVATION function///
cd C:\Program Files\Microsoft Office\Office16
echo|set /p=connecting to kms server... 
call :StartTaskWithSpinner cscript ospp.vbs /sethst:kms.03k.org >nul
echo %GREEN%[CONNECTED]%RESET%
echo|set /p=trying to activate key... 
call :StartTaskWithSpinner cscript ospp.vbs /act >nul
echo %GREEN%[DONE]%RESET%
exit /b


:ACTIVATEWIN
cls
title WinMaker
echo.
echo     %RED%████████  %GREEN%████████%RESET%
echo     %RED%████████  %GREEN%████████%RESET%     888       888 d8b  
echo     %RED%████████  %GREEN%████████%RESET%     888   o   888 Y8P          
echo     %RED%████████  %GREEN%████████%RESET%     888  d8b  888              
echo                            888 d888b 888 888 88888b. 
echo     %BLUE%████████  %YELLOW%████████%RESET%     888d88888b888 888 888 "88b 
echo     %BLUE%████████  %YELLOW%████████%RESET%     88888P Y88888 888 888  888 
echo     %BLUE%████████  %YELLOW%████████%RESET%     8888P   Y8888 888 888  888 
echo     %BLUE%████████  %YELLOW%████████%RESET%     888P     Y888 888 888  888
echo.
echo     888b     d888        d8888 888    d8P  8888888888 8888888b.  
echo     8888b   d8888       d88888 888   d8P   888        888   Y88b 
echo     88888b.d88888      d88P888 888  d8P    888        888    888 
echo     888Y88888P888     d88P 888 888d88K     8888888    888   d88P 
echo     888 Y888P 888    d88P  888 8888888b    888        8888888P"  
echo     888  Y8P  888   d88P   888 888  Y88b   888        888 T88b   
echo     888   "   888  d8888888888 888   Y88b  888        888  T88b  
echo     888       888 d88P     888 888    Y88b 8888888888 888   T88b
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [^^!] HWID 3.0 WINDOWS ACTIVATOR                                                                      │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ You are about to activate Windows using the HWID activation exploit.                                │
echo         │ Are you sure you want to proceed? If Windows is already activated, running this script may pose     │
echo         │ unnecessary risks.                                                                                  │
echo         ├───────────────────────────────────────────────┬─────────────────────────────────────────────────────┤
echo         │ %LIGHTBLUE%(Y) Run HWID activation%RESET%                       │ %LIGHTBLUE%(N) Cancel%RESET%                                          │
echo         ╰───────────────────────────────────────────────┴─────────────────────────────────────────────────────╯
choice /c YN /n
if %errorlevel%==1 (start assets\HWIDActivation.cmd)
if %errorlevel%==2 (goto STARTMANAGER)
exit



REM instance to xml editor
:CUSTOM
echo.
echo %DARKER%[OFFICE CONFIGURATOR]%RESET%
call assets\xmlmaker.bat

findstr /B /I "cancel" "assets\config-temp.xml" >nul
if %errorlevel%==0 (
    goto PRESTART
)

set config_file=config-temp.xml
goto PROCESS



:DEV
cls
echo.
echo         ╭─────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo         │ [i] DEV MODE                                                                                        │
echo         ├─────────────────────────────────────────────────────────────────────────────────────────────────────┤
echo         │ Here you can run commands with more detailed information about what is happening (for debugging)    │
echo         ├─────────────────────────────┬───────────────────────────────────────────────────────────────────────┤
echo         │ connect                     │ Connects to kms server.                                               │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ rearm                       │ Rearms kms, clears cache.                                             │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ act                         │ Runs kms activation.                                                  │
echo         ├─────────────────────────────┼───────────────────────────────────────────────────────────────────────┤
echo         │ menu                        │ Exits dev mode.                                                       │
echo         ╰─────────────────────────────┴───────────────────────────────────────────────────────────────────────╯
:DEVI
echo %LIGHTBLUE%Waiting for command input...%RESET%
set input=null
set /p input=">"
echo %YELLOW%[COMMAND STARTED]%RESET%

if %input%==connect (
	@echo on
	cd C:\Program Files\Microsoft Office\Office16
	cscript ospp.vbs /sethst:kms.03k.org
	cd %~dp0
	@echo off
)
if %input%==rearm (
	@echo on
	cd C:\Program Files\Microsoft Office\Office16
	cscript ospp.vbs /rearm
	cd %~dp0
	@echo off
)
if %input%==act (
	@echo on
	cd C:\Program Files\Microsoft Office\Office16
	cscript ospp.vbs /act
	cd %~dp0
	@echo off
)
if %input%==menu (
	goto START
)
echo %YELLOW%[COMMAND FINISHED]%RESET%
goto DEVI




:StartTaskWithSpinner
rem Start the external process in the background
start /B %1 %2 %3 %4 %5 %6 %7 %8 %9

rem Spinning animation in a single line, overwriting previous symbol
:spin
for %%a in (/ - \) do (
    set "spin=%%a"
    <nul set /p=!spin!
    timeout /nobreak /t 1 >nul
    <nul set /p=
)

rem Wait until the external process finishes
tasklist /FI "IMAGENAME eq %1" 2>NUL | find /I "%1" >NUL
if %ERRORLEVEL%==0 goto spin

rem Exit the method
exit /b







:OUT