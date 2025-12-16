@echo off
setlocal

:: This script allows you to drag and drop a Houzz Excel file onto it.
:: It will ask for the Project Name and Proposal Number, then run the conversion.

:: 1. Check if a file was dragged onto the script
if "%~1"=="" (
    echo ========================================================
    echo  ERROR: No file provided.
    echo  Please drag and drop your Excel file onto this icon.
    echo ========================================================
    pause
    goto :EOF
)

:: 2. Display the detected file
echo Input File detected: "%~1"
echo.

:: 3. Ask for details
set /p project="Enter Project Name (e.g. Paxman DDS): "
set /p proposal="Enter Proposal Number (e.g. Art_12_16): "

echo.
echo Running conversion...
echo.

:: 4. Run the Python script
:: We use %~dp0 to refer to the directory where this script sits
python "%~dp0houzz_to_grist.py" "%~1" "%~dp0houzz_import.csv" "%project%" "%proposal%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo  SUCCESS! 
    echo  File created at: %~dp0houzz_import.csv
    echo ========================================================
    timeout /t 10
) else (
    echo.
    echo ========================================================
    echo  ERROR DETECTED during conversion.
    echo ========================================================
    pause
)
