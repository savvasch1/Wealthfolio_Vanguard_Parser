# Vanguard Transactions Converter

A small Tkinter desktop app that converts Vanguard transaction exports into normalized CSV files. 

## What it does

- Opens a GUI with a single **Convert EXCEL file** button. 
- Loads Vanguard Excel workbooks with `openpyxl`. 
- Removes the `Summary` worksheet when present, deletes the first four rows of each remaining worksheet, and trims rows starting at `Balance`. 
- Normalizes transaction data into these output columns: `Date`, `Symbol`, `Quantity`, `ActivityType`, `unitPrice`, `currency`, `fee`, and `Amount`. 
- Writes one CSV per worksheet to the current user's Desktop. 

## Requirements

- Python 3.10+ recommended.
- Python packages: `pandas`, `numpy`, and `openpyxl`. Use the versions in the requirements file
- `tkinter` must be available in your Python installation because the app uses a Tkinter GUI. 

## Run from source

```bash
python main.py
```

## Build with PyInstaller

PyInstaller can package Python programs into standalone executables, and `--onefile` creates a single distributable file. For a Tkinter GUI app, `--windowed` suppresses the console window on Windows and macOS. Do not supress it as you need it for logs

### 1. Install dependencies
If using local python

```bash
pip install -r requirements.txt
```

### 2. Build a single-file executable

```bash
pyinstaller --noconfirm --onefile --windowed --name VanguardConverter main.py
```

### 3. Built output

- Windows: `dist/VanguardConverter.exe` after the build completes.
- macOS/Linux: `dist/VanguardConverter` after the build completes. 

## Useful build options

- `--onefile` packages the app into a single file.
- `--windowed` disables the console window for GUI apps. 
- `--clean` clears cached build files before packaging. 


Example:
with window suppresed
```bash
pyinstaller --noconfirm --clean --onefile --windowed --name VanguardConverter --icon app.ico main.py
```

with window 
```bash
pyinstaller --noconfirm --clean --onefile --name VanguardConverter --icon app.ico main.py
```



## Notes

- The current GUI exposes only the Excel conversion flow. 
- Output files are named after each worksheet and saved to the Desktop. 
- PyInstaller builds for the current platform, so you should build on the same operating system you plan to distribute for. [web:51]