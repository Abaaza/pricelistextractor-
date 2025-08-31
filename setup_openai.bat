@echo off
echo =====================================
echo OpenAI API Key Setup for Enhancement
echo =====================================
echo.

echo This script will help you set up your OpenAI API key for pricelist enhancement.
echo The API key will be used to:
echo   - Clean up fragmented descriptions
echo   - Properly categorize items
echo   - Generate relevant keywords
echo   - Ensure construction terminology consistency
echo.

set /p API_KEY="Enter your OpenAI API key (starts with sk-): "

if "%API_KEY%"=="" (
    echo No API key provided. Exiting...
    pause
    exit /b 1
)

echo.
echo Setting API key as environment variable...
setx OPENAI_API_KEY "%API_KEY%"

echo.
echo API key has been set permanently!
echo.
echo Now running the enhanced extraction...
echo.

set OPENAI_API_KEY=%API_KEY%
python pricelist_extractor_enhanced.py

pause