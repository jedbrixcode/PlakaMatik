build_orchestrator:
@echo off
echo Compiling PlakaMatik Unified Orchestrator...
echo.

color a

cd /D "%~dp0"

python -m pip install pyinstaller

rem This checks if pyinstaller actually exists before running
python -m PyInstaller --noconsole --onefile --name "orchestrator" main.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Wait, something went wrong during compilation. Check the errors above.
) else (
    echo.
    echo Compilation complete. The orchestrator.exe is located in the 'dist' folder.
)
pause