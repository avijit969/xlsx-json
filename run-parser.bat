@echo off
setlocal

if "%~1"=="help" goto usage
if "%~1"=="/?" goto usage
if "%~2"=="" goto usage

set BANK_NAME=%~1
set EXCEL_FILE=%~2
set JAR_FILE=target\xlsx-to-json-parser-1.0-SNAPSHOT.jar

if not exist "%JAR_FILE%" (
    echo [INFO] JAR file not found. Building the project with Maven...
    call mvn clean package
    if errorlevel 1 (
        echo [ERROR] Maven build failed.
        exit /b 1
    )
)

echo [INFO] Running parser for Bank: "%BANK_NAME%" on File: "%EXCEL_FILE%"
java -jar "%JAR_FILE%" "%BANK_NAME%" "%EXCEL_FILE%"

if errorlevel 1 (
    echo [ERROR] Parser execution failed.
    exit /b 1
)

echo [INFO] Complete!
goto end

:usage
echo Usage: run-parser.bat ^<BankName^> ^<ExcelFilePath^>
echo Example: run-parser.bat "ICICI" "xlsx-files\Bank-Statement-Template-2-TemplateLab.xlsx"
exit /b 1

:end
endlocal
