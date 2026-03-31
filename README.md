# Print

Windows desktop app for designing and printing product labels.

## What It Does

- Loads label data from Excel files
- Supports barcode-based labels
- Lets you adjust layout settings
- Prints to Windows printers
- Packages into a standalone `.exe`

## Tech Stack

- Python
- PySide6
- `openpyxl`
- `python-barcode`
- Pillow
- `pywin32`
- PyInstaller
- Inno Setup

## Run Locally

Create and activate a virtual environment, then install the dependencies:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install PySide6 openpyxl python-barcode pillow pywin32 pyinstaller
```

Start the app:

```powershell
python app.py
```

## Build The EXE

```powershell
pyinstaller Print.spec
```

The built executable will be created under `dist/`.

## Build The Installer

Use the included Inno Setup script:

```text
PrintInstaller.iss
```

Update any machine-specific paths in that file before building the installer.

## Project Files

- `app.py` - main application
- `layout.json` - label layout settings
- `settings.json` - app settings
- `Print.spec` - PyInstaller build config
- `PrintInstaller.iss` - Inno Setup installer config

## Notes

- This project is Windows-only because it depends on the Windows print stack.
- The repository ignores local build output, virtual environment files, and editor settings.
