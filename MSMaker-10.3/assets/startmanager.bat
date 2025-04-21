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





echo.
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
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%                                                    %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******.           ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         *******            ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%         -***               ***********             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%              .************************             %RESET%│ │%CYAN%       ############## #####################         %RESET%│
echo       │%LIGHTRED%               ************************             %RESET%│ │%CYAN%                ##### #####################         %RESET%│
echo       │%LIGHTRED%                  *****************-                %RESET%│ │%CYAN%                          #################         %RESET%│
echo       │%LIGHTRED%                      ******+                       %RESET%│ │%CYAN%                                     ######         %RESET%│
echo       │%LIGHTRED%                                                    %RESET%│ │%CYAN%                                                    %RESET%│
echo       ╰────────────────────────────────────────────────────╯ ╰────────────────────────────────────────────────────╯
echo       ╭───────────────────────────────────────────────────────────────────────────────────────────────────────────╮
echo       │ What are you willing to activate?                                                                         │
echo       ├─────────────────────────────────────────────────────┬─────────────────────────────────────────────────────┤
echo       │ %LIGHTBLUE%(1) Office LTSC 2024 [Office-Maker]%RESET%                 │ %LIGHTBLUE%(2) Windows 10/11 [Win-Maker]%RESET%                       │
echo       ╰─────────────────────────────────────────────────────┴─────────────────────────────────────────────────────╯
choice /c 12 /n
exit /b %errorlevel%
            
                                                  
                                                  

                                                 
                                                  
                                                  
                
