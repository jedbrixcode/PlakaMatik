# PlakaMatik Production Packaging Script
# ----------------------------------------
# This script bundles the decoupled Python ecosystem utilizing PyInstaller,
# and pairs it internally with the Flutter release build.

Write-Host "Initializing PlakaMatik Compiler Pipeline..." -ForegroundColor Cyan

# 1. Ensure packages are present
Write-Host "Checking for PyInstaller..."
pip install pyinstaller

# 2. Package Python Backend Handlers
Write-Host "Packaging python automation hooks..."
Set-Location -Path ".\python_engine\Core"

pyinstaller --noconfirm --onedir --console --name "LTO_ExportManager" "main.py"
pyinstaller --noconfirm --onedir --console --name "Print_Spooler" "send_to_printer.py"

# Combine unified executables into a single isolated release bin to share DLL weights
New-Item -ItemType Directory -Force -Path "..\dist\plakamatic_engine"
Copy-Item -Path ".\dist\LTO_ExportManager\*" -Destination "..\dist\plakamatic_engine" -Recurse -Force
Copy-Item -Path ".\dist\Print_Spooler\*" -Destination "..\dist\plakamatic_engine" -Recurse -Force

Set-Location -Path "..\.."

# 3. Compile Flutter
Write-Host "Compiling Flutter Windows Release..."
flutter build windows

# 4. Integrate Pipeline
Write-Host "Merging binaries..."
Copy-Item -Path ".\python_engine\dist\plakamatic_engine" -Destination ".\build\windows\x64\runner\Release\python_engine" -Recurse -Force

Write-Host "===========================" -ForegroundColor Green
Write-Host "COMPILATION SUCCESSFUL!" -ForegroundColor Green
Write-Host "The PlakaMatik isolated application is located at:"
Write-Host ".\build\windows\x64\runner\Release\"
Write-Host "===========================" -ForegroundColor Green
