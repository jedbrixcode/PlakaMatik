# PlakaMatik — Automated UV Plate Printing System

> **LTO Protocol Plate Manufacturing Layout Maker**  
> A production desktop application for the automated generation and UV printing of Philippine LTO (Land Transportation Office) protocol plates, built with Flutter + Python (CorelDRAW COM Automation).

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [System Requirements](#system-requirements)
- [Getting Started (Developers)](#getting-started-developers)
- [Project Structure](#project-structure)
- [Flutter UI — Features](#flutter-ui--features)
- [Python Engine — How It Works](#python-engine--how-it-works)
- [Configuration Bridge](#configuration-bridge)
- [Path Resilience & Deployment Notes](#path-resilience--deployment-notes)
- [Building for Production](#building-for-production)
- [Creating the Installer](#creating-the-installer)
- [Known Issues & Troubleshooting](#known-issues--troubleshooting)
- [Deployment Checklist](#deployment-checklist)

---

## Overview

PlakaMatik is a Windows-only desktop automation system that:

1. Accepts operator input (plate identifier, middle text, plate type: **MV** or **MC**).
2. Injects the data into CorelDRAW template files (`.cdr`) via Windows COM automation.
3. Exports a silent, layer-filtered PDF from CorelDRAW.
4. Composes single or dual-plate **A3 landscape** PDFs using `pypdf`.
5. Sends the final PDF to a **Canon iX6700** UV inkjet printer via SumatraPDF CLI (GDI pipeline).

The system operates in two modes:
- **Single Plate** — one plate per job, immediate preview + print.
- **Batch Print** — queue of plates processed in configurable chunks (default: 2 per A3 sheet).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Flutter UI (Dart)                        │
│  SinglePlateView ─┐                                             │
│  BatchViewModel  ─┼─► BackendService.runOrchestrator()          │
│  SettingsViewModel│        │                                    │
│  TroubleshootingView       │  Process.run(orchestrator.exe)     │
└────────────────────────────┼────────────────────────────────────┘
                             │  [--config config.json]
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                  Python Orchestrator (PyInstaller EXE)          │
│                                                                 │
│  main.py ──► config.py ──► config_manager.py (JSON bridge)      │
│          ──► data_processor.py (parse input.txt)                │
│          ──► corel_engine.py (CorelDRAW COM via win32com)       │
│          ──► export_manager.py (pypdf A3 composition)           │
│          ──► send_to_printer.py (SumatraPDF → Canon iX6700)    │
└─────────────────────────────────────────────────────────────────┘
```

**Communication Protocol:**  
Flutter writes `config.json` and `input.txt` to the operator's Documents folder before launching `orchestrator.exe --config <path>`. The Python engine reads these files, processes the data, and writes PDFs back to the same Documents folder. Flutter polls for the output files.

---

## System Requirements

| Component | Requirement |
|-----------|-------------|
| OS | Windows 10 / 11 (64-bit) |
| CorelDRAW | CorelDRAW 2018 (COM automation required) |
| Printer | Canon iX6700 series (GDI inkjet) |
| Runtime | Flutter 3.x SDK (for development only) |
| Python | Python 3.10+ (for rebuilding the engine only) |
| Inno Setup | Inno Setup 6 (for building the installer) |
| SumatraPDF | Portable SumatraPDF.exe (placed in `Documents\PlakaMatik Files\bin\`) |

---

## Getting Started (Developers)

### 1. Clone the repository

```bash
git clone https://github.com/jedbrixcode/PlakaMatik.git
cd PlakaMatik/plakamatic_flutterui
```

### 2. Install Flutter dependencies

```bash
flutter pub get
```

### 3. Run in development mode

```bash
flutter run -d windows
```

> **Dev mode note:** When running from source, the `BackendService` fallback path points to `python_engine/Core/dist/orchestrator.exe`. You must build the Python engine first (see below).

### 4. Build the Python engine

```bash
cd python_engine/Core
build_orchestrator.bat
```

This runs PyInstaller and produces `dist/orchestrator.exe`. The Flutter build system copies this into `assets/orchestrator.exe` automatically via `pubspec.yaml`.

### 5. Build the Flutter Windows app

```bash
flutter build windows
```

Output: `build/windows/x64/runner/Release/plakamatic_flutterui.exe`

---

## Project Structure

```
plakamatic_flutterui/
├── assets/
│   ├── orchestrator.exe            # Bundled Python engine (PyInstaller)
│   ├── Templates/
│   │   └── Main Templates/
│   │       ├── MV_PLATE.cdr        # Motorcycle plate template
│   │       └── MC_PLATE.cdr        # Motor vehicle plate template
│   ├── BACKGROUND.png
│   └── BACKGROUND_DARKMODE.png
│
├── lib/
│   ├── main.dart                   # App entry point, service initialization
│   ├── services/
│   │   ├── backend_service.dart    # Asset extraction, orchestrator execution
│   │   ├── path_service.dart       # Centralized runtime path resolution
│   │   ├── cleanup_service.dart    # Session temp file cleanup
│   │   └── log_watcher_service.dart # Tails Python daily log → log console
│   ├── viewmodels/
│   │   ├── batch_viewmodel.dart    # Batch queue state, chunk orchestration
│   │   ├── navigation_viewmodel.dart
│   │   └── settings_viewmodel.dart # Persists settings to config.json bridge
│   ├── views/
│   │   ├── main_layout.dart        # Shell: sidebar + background + routing
│   │   ├── single_plate_view.dart  # Single plate form + preview
│   │   ├── multiple_plate_view.dart# Batch queue builder + chunk execution
│   │   ├── settings_view.dart      # All operator-configurable settings
│   │   ├── information_view.dart   # About / version info
│   │   └── troubleshooting_view.dart # Interactive diagnostic hub
│   ├── widgets/
│   │   ├── console_log_widget.dart # Live Python log tail (SelectableText)
│   │   └── print_countdown_dialog.dart # Print confirmation + spool dialog
│   └── utils/
│       └── input_sanitizer.dart
│
├── python_engine/
│   └── Core/
│       ├── main.py                 # Orchestrator entry point
│       ├── config.py               # Static path configuration
│       ├── config_manager.py       # Reads/writes config.json bridge
│       ├── corel_engine.py         # CorelDRAW COM automation
│       ├── export_manager.py       # PDF export + A3 pypdf composition
│       ├── send_to_printer.py      # 3-tier spooler (SumatraPDF → ShellExecute)
│       ├── data_processor.py       # Parses 4-space delimited input.txt
│       ├── text_mapper.py          # COM text field injection
│       ├── engine_logger.py        # Daily rotating log writer
│       ├── session_manager.py      # Temp file cleanup
│       └── build_orchestrator.bat  # PyInstaller build script
│
├── plakamatic_installer.iss        # Inno Setup 6 installer script
├── pubspec.yaml
└── README.md
```

---

## Flutter UI — Features

### Single Plate View
- Input fields: **Identifier** (plate number), **Middle** text, **Type** (MV/MC).
- "Generate" triggers `orchestrator.exe --config config.json`.
- On success, scans the `Outputs/` folder for the latest `*_PREVIEW.pdf` and displays it inline using `syncfusion_flutter_pdfviewer`.
- "Print" opens the countdown dialog which dispatches `--action spool`.

### Batch Print View
- Drag-and-drop queue of plates with `identifier` and `middle` fields.
- Configurable **chunk size** (default: 2 per A3 sheet).
- "Generate Next" processes the next chunk and previews the A3 composite.
- "Resume from last" bookmark — remembers the last successfully printed chunk.

### Settings View
- **Printer Name** — target printer identifier (must match Windows Printers exactly).
- **Trial Bypass Delay** — seconds to wait after CorelDRAW launches before sending `Alt+Z`.
- **CorelDRAW Visible** — show/hide CorelDRAW window during automation.
- **Global DX/DY Offsets** — millimeter adjustments applied to all plate text positions.
- **Dark/Light Mode** — background theme toggle.
- Save/Reset to `config.json`.

### Troubleshooting Hub
- **Category filters**: Hardware, Engine, Automation.
- **Flush Spooler** button — runs `net stop/start spooler` to clear stuck print jobs.
- **Asset Repair** button — force re-extracts `orchestrator.exe` and CDR templates from bundled assets.
- 11 expandable issue cards with step-by-step resolutions.

### Log Console
- Tails `Documents\PlakaMatik Files\Logs\daily_<date>.log` in real time.
- `SelectableText` — operators can select and copy full Windows paths.
- Auto-scrolls to latest entry.

---

## Python Engine — How It Works

### Execution Flow (`main.py`)

```
1. Parse --config argument → load config.json
2. Parse input.txt (4-space delimited: identifier    middle    type)
3. Connect to CorelDRAW via win32com.client.Dispatch("CorelDRAW.Application")
4. For each plate record:
   a. Open MV_PLATE.cdr or MC_PLATE.cdr (copy in temp dir)
   b. Inject identifier + middle text into "Print Layer" shapes
   c. Export PREVIEW PDF (all layers except Guides)
   d. Export PRINT PDF (Print Layer only)
5. Compose A3 PDF using pypdf merge_transformed_page
   - 1 plate → centered on A3
   - 2 plates → top half + bottom half
6. On --action spool: send PRINT.pdf to printer via SumatraPDF CLI
```

### Trial Screen Bypass
CorelDRAW 2018 trial builds show a modal dialog on launch. The Python engine:
1. Arms the bypass delay via `corel_engine.bypass_trial_screen(delay=N)`.
2. Dispatches CorelDRAW via COM (this triggers the launch).
3. Waits `N` seconds (the delay) for the trial screen to appear.
4. Sends `Alt+Z` via pyautogui to dismiss it.
5. Waits 7 seconds for the UI to stabilize.

The delay is configurable from the Flutter Settings UI (0–7+ seconds).

### Spooler Strategy (send_to_printer.py)
The Canon iX6700 is a **GDI inkjet printer** — it cannot decode raw PDF bytes. The spooler uses a 3-tier strategy:

| Tier | Method | Notes |
|------|--------|-------|
| 1 | **SumatraPDF CLI** (`-print-to -silent -exit-on-print`) | GDI rendering — primary method |
| 2 | **ShellExecute `printto`** | GDI via registered PDF handler — fallback |
| 3 | **Hard fail** | Actionable error message shown in Flutter UI |

A file-lock retry loop (10 attempts × 1s) waits for CorelDRAW to release the PDF before transmission.

---

## Configuration Bridge

`Documents\PlakaMatik Files\config.json` is the communication contract between Flutter and Python.

```json
{
  "PRINTER_NAME": "Canon iX6700 series",
  "TRIAL_BYPASS_DELAY": 5,
  "CORELDRAW_VISIBLE": false,
  "GLOBAL_OFFSETS": {
    "dx": 0.0,
    "dy": 0.0
  },
  "PLATES": [
    {
      "identifier": "ABC 1234",
      "middle": "JUAN DE LA CRUZ",
      "type": "MV"
    }
  ]
}
```

Flutter writes this file via `SettingsViewModel` and `BatchViewModel` before launching the orchestrator. Python reads it via `config_manager.load_config(path)`.

---

## Path Resilience & Deployment Notes

### The Space-in-Path Problem
Windows machines with usernames containing spaces (e.g., `Win10 PRO`) cause path parsing failures when commands are run through `cmd.exe`.

**Solution:** All `Process.run` calls use:
```dart
runInShell: false       // Calls CreateProcess() directly — bypasses cmd.exe
```
All paths are built with `p.join()` (native backslashes) and `Platform.pathSeparator` normalization.

### Asset Extraction
On first launch, `BackendService.initialize()` extracts `orchestrator.exe` from Flutter assets to:
```
Documents\PlakaMatik Files\bin\orchestrator.exe
```
This ensures the engine always runs from a writable, non-UAC-protected location.

### Template Self-Healing
`BackendService.ensureTemplates()` checks for `MV_PLATE.cdr` and `MC_PLATE.cdr` on every startup. If either is missing, it restores them from bundled assets.

### PathService
All runtime paths are resolved through `PathService.resolve()` which returns a `PlakaMatikPaths` object with pre-built, normalized paths. **Never hardcode `PlakaMatik Files` path segments** — always use `PathService`.

---

## Building for Production

### Step 1: Build the Python engine

```bash
cd python_engine/Core
build_orchestrator.bat
```

Verify: `dist/orchestrator.exe` exists and is > 15 MB.

### Step 2: Copy SumatraPDF portable
Download [SumatraPDF portable](https://www.sumatrapdfreader.org/download-free-pdf-viewer) and place `SumatraPDF.exe` in:
```
python_engine/Core/dist/SumatraPDF.exe
```
It will be bundled by the Inno Setup installer into the operator's bin/ folder.

### Step 3: Build Flutter

```bash
flutter build windows
```

### Step 4: Package with Inno Setup

```bash
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" plakamatic_installer.iss
```

Output: `Output\PlakaMatik_Setup.exe`

---

## Creating the Installer

The `plakamatic_installer.iss` script:
- Installs the Flutter app to `C:\Program Files\PlakaMatik` (or `Program Files (x86)`).
- Copies CDR templates to `{userdocs}\PlakaMatik Files\CorelDRAW Templates\Main Templates\`.
- Creates Desktop and Start Menu shortcuts.
- Bundles VC++ runtime DLLs (`vcruntime140.dll`, `msvcp140.dll`).

**Important:** The installer requires admin privileges for `Program Files` installation. Templates are written to `{userdocs}` so they are writable by the application without UAC prompts.

---

## Known Issues & Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| Blank page printed | Canon GDI driver can't read raw PDF | Ensure `SumatraPDF.exe` is in `bin/` folder |
| "Access is Denied" on spool | Printer not shared | Enable printer sharing in Windows Settings |
| COM Error: RPC unavailable | CorelDRAW closed during automation | Kill CorelDRAW in Task Manager, retry |
| Trial screen not bypassed | `bypass_trial_screen` fires too early | Increase Trial Bypass Delay to 7–10s in Settings |
| Template not found | CDR files deleted from Documents | Click "Asset Repair" in Troubleshooting Hub |
| ProcessException: Invalid directory | Mixed path slashes | Already fixed — paths use `p.join()` and `runInShell: false` |
| 0-byte PDF exported | CorelDRAW ran out of memory | Restart CorelDRAW and regenerate |
| Ñ character garbled | Non-UTF-8-BOM input.txt | Already fixed — Flutter writes BOM header automatically |

---

## Deployment Checklist

Before deploying to a production machine:

- [ ] CorelDRAW 2018 is installed and activated (or trial mode configured)
- [ ] Canon iX6700 driver is installed and printer is set as "Shared"
- [ ] `SumatraPDF.exe` (portable) is placed in `Documents\PlakaMatik Files\bin\`
- [ ] `PlakaMatik_Setup.exe` is run as Administrator
- [ ] On first launch, verify the log console shows "orchestrator.exe extracted" and "Template OK" messages
- [ ] Test single plate generation before batch
- [ ] Set Trial Bypass Delay in Settings to match the target machine's CorelDRAW load time

---

## License

Internal use — LTO Plate Manufacturing Operations.  
Developed by **JedImsonDev** for the LTO internship project, 2026.
