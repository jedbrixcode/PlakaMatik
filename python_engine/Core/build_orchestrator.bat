@echo off
echo Compiling PlakaMatik Unified Orchestrator...
echo.

color a

:: Set the current directory to where the batch file is
cd /D "%~dp0"

:: Set the destination path variable
set "DEST_FOLDER=C:\Users\Window 10\Documents\Jed Internship\Project\plakamatic_flutterui\assets"

:: Ensure PyInstaller and required Python packages are installed
python -m pip install pyinstaller

:: Compile the Python script
:: --noconsole para invisible, --onefile para compact
python -m PyInstaller --noconsole --onefile --name "orchestrator" main.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Wait, something went wrong during compilation. Check the errors above.
    color c
) else (
    echo.
    echo Compilation complete. Moving orchestrator.exe to Assets...

    :: Check if the destination folder exists, create if not
    if not exist "%DEST_FOLDER%" mkdir "%DEST_FOLDER%"

    :: Copy/Overwrite the .exe to your Flutter Assets
    copy /Y "dist\orchestrator.exe" "%DEST_FOLDER%\orchestrator.exe"

    if %ERRORLEVEL% EQU 0 (
        echo.
        echo Success! orchestrator.exe copied to:
        echo %DEST_FOLDER%
        color a

        :: --- Cleanup: remove PyInstaller build artifacts ---
        echo Cleaning up build artifacts...
        if exist "build" (
            rmdir /S /Q "build"
            echo   build\ folder removed.
        )
        if exist "orchestrator.spec" (
            del /Q "orchestrator.spec"
            echo   orchestrator.spec removed.
        )
        if exist "dist" (
            rmdir /S /Q "dist"
            echo   dist\ folder removed.
        )
        echo Cleanup complete.
    ) else (
        echo.
        echo Error: Failed to copy the file to the assets folder. Check permissions.
        color e
    )
)

pause