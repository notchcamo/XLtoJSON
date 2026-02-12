# XLtoJSON - Build & Deployment Guide

## Overview
XLtoJSON is a standalone executable for bidirectional conversion between Excel (.xlsx) and JSON files.

## Executable Information
- **Location**: `dist/XLtoJSON.exe`
- **Size**: ~31 MB
- **Platform**: Windows 64-bit
- **Dependencies**: All bundled (no Python required)

## Building from Source

### Prerequisites
- Python 3.14+
- Virtual environment activated

### Build Steps

1. **Install dependencies**:
   ```bash
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Build executable** (Option 1 - Automated):
   ```bash
   build.bat
   ```

3. **Build executable** (Option 2 - Manual):
   ```bash
   venv\Scripts\activate
   pyinstaller --onefile --name XLtoJSON --clean __main__.py
   ```

4. **Output**:
   - Executable: `dist\XLtoJSON.exe`
   - Build artifacts: `build\` (can be deleted)
   - Spec file: `XLtoJSON.spec` (PyInstaller configuration)

## Usage

### Excel to JSON
```bash
XLtoJSON.exe data.xlsx output_folder\
```

### JSON to Excel
```bash
XLtoJSON.exe data.json output_folder\
```

### Command Format
```bash
XLtoJSON.exe <source_file> <output_directory>
```

**Arguments**:
- `source_file`: Path to .xlsx or .json file
- `output_directory`: Folder where converted file will be saved

## Deployment

### Standalone Distribution
The `XLtoJSON.exe` file can be distributed and run on any Windows machine without requiring Python installation.

### Recommended Deployment
1. Copy `dist\XLtoJSON.exe` to target location
2. Optionally add to system PATH for command-line access
3. Create a desktop shortcut if needed

### File Size
The executable is ~31 MB because it includes:
- Python runtime
- pandas library
- openpyxl library
- All dependencies

## Supported Features
- Simple scalar columns
- Flat arrays (numbered columns: `tags_0`, `tags_1`)
- Structured columns (merged headers with nested objects)
- Non-ASCII character preservation
- Automatic deduplication when appending JSON to existing Excel

## Troubleshooting

### Build Warnings
- "Hidden import 'jinja2' not found" - Can be safely ignored (jinja2 is an optional pandas dependency)

### Antivirus False Positives
Some antivirus software may flag PyInstaller executables. This is a known issue with packaged Python applications. You can:
1. Add exception in antivirus software
2. Submit the file to antivirus vendor for whitelisting
3. Build with code signing certificate (for production)

### File Permissions
Ensure XLtoJSON.exe has execute permissions and write access to the output directory.

## Advanced Build Options

### Optimize for Size
```bash
pyinstaller --onefile --name XLtoJSON --clean --strip --upx-dir=<upx_path> __main__.py
```
(Requires UPX compressor)

### Add Custom Icon
```bash
pyinstaller --onefile --name XLtoJSON --icon=icon.ico __main__.py
```

### Debug Build
```bash
pyinstaller --onefile --name XLtoJSON --debug=all __main__.py
```

## Clean Build Artifacts

```bash
# Windows
rmdir /s /q build dist
del XLtoJSON.spec

# Linux/Mac
rm -rf build dist XLtoJSON.spec
```

## Version Info
- PyInstaller: 6.18.0
- Python: 3.14.2
- Build date: 2026-02-12
