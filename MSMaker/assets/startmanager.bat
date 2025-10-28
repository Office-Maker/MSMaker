:0
@echo off
chcp 65001 >nul
cls
setlocal enabledelayedexpansion
REM presetter
title MSMaker
set RED=!ESC![31m
set LIGHTRED=!ESC![91m
set CYAN=!ESC![96m
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

set hwidmissing=%1
if %hwidmissing%==true (
	set CYAN=%DARKER%
)

echo.
echo.
echo        What are you willing to activate?
echo       ╭────────────────────────────────────────────────────╮ ╭────────────────────────────────────────────────────╮
echo       │%LIGHTRED%                                                    %RESET%│ │%CYAN%                                                    %RESET%│
echo       │%LIGHTRED%                             #####-:                %RESET%│ │%CYAN%                                   ########         %RESET%│
echo       │%LIGHTRED%                   ****#####-------+                %RESET%│ │%CYAN%                         ##################         %RESET%│
echo       │%LIGHTRED%               ***********##-----------             %RESET%│ │%CYAN%               ###### #####################         %RESET%│
echo       │%LIGHTRED%           *****************-----------             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *************      -----------             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         ********           -----------             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           -----------             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           -----------             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
if %hwidmissing%==true (goto 1)
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%                                                    %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
goto 2

:1
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ##########              ############         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%RESET%                   REDOWNLOAD                       %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ##########              ############         %RESET%│

:2
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******            ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         -***               ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%              .************************             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%               ************************             %RESET%│ │%CYAN%                ##### #####################         %RESET%│
echo       │%LIGHTRED%                  *****************-                %RESET%│ │%CYAN%                          #################         %RESET%│
echo       │%LIGHTRED%                      ******+                       %RESET%│ │%CYAN%                                     ######         %RESET%│
echo       │%LIGHTRED%                                                    %RESET%│ │%CYAN%                                                    %RESET%│
if %hwidmissing%==true (goto 3)
echo       │%LIGHTRED%          (1) Microsoft 365 [Office-Maker]          %RESET%│ │%CYAN%           (2) Windows 10/11 [Win-Maker]            %RESET%│
goto 4
:3
echo       │%LIGHTRED%          (1) Microsoft 365 [Office-Maker]          %RESET%│ │%RESET%     (2) REDOWNLOAD %CYAN%Windows 10/11 [Win-Maker]       %RESET%│
:4
echo       ╰────────────────────────────────────────────────────╯ ╰────────────────────────────────────────────────────╯
echo.
choice /c 12 /n
if %hwidmissing%==true (goto 5)
goto 6
:5
if %errorlevel%==2 (
	start https://github.com/Office-Maker/MSMaker
	goto 0
)

:6
exit /b %errorlevel%
            
                                                  
                                                  
                                               
                                                  
                                                  
                
