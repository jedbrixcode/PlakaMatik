@echo off
echo Compiling PlakaMatik Unified Orchestrator...
echo.

cd /D "%~dp0"
pip install pyinstaller

rem Compile main.py as a single executable with no console window
pyinstaller --noconsole --onefile --name "orchestrator" main.py

echo.
echo Compilation complete. The orchestrator.exe is located in the 'dist' folder.
pause
