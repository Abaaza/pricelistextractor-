@echo off
echo MJD Pricelist Extraction Setup
echo ==============================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH
    exit /b 1
)

REM Install requirements
echo Installing required packages...
pip install -r requirements.txt

echo.
echo Setup complete!
echo.

REM Ask for OpenAI API key
set /p USE_AI="Do you want to use OpenAI for keyword generation? (y/n): "
if /i "%USE_AI%"=="y" (
    set /p OPENAI_KEY="Enter your OpenAI API key: "
    set OPENAI_API_KEY=%OPENAI_KEY%
)

echo.
echo Running extraction script...
python pricelist_extractor.py

pause