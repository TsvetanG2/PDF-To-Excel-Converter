# Build Instructions - PDF to Excel Converter

## Prerequisites

1. **Python 3.8+** with pip
2. **Java Runtime Environment (JRE)** - required for tabula-py
3. **Inno Setup 6.x** - download from https://jrsoftware.org/isdl.php

## Step 1: Install Dependencies

```bash
cd ConverterApp
pip install -r requirements.txt
pip install pyinstaller
```

## Step 2: Build Executable

Run the build script:

```bash
build.bat
```

Or manually:

```bash
cd ConverterApp
pyinstaller --noconfirm --onedir --console --name "PDF-To-Excel-Converter" ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --add-data "pdftoexcel.py;." ^
    --hidden-import=pdfplumber ^
    --hidden-import=tabula ^
    --hidden-import=fitz ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=flask ^
    --hidden-import=werkzeug ^
    --collect-all pdfplumber ^
    --collect-all tabula ^
    --collect-all PyMuPDF ^
    launcher.py
```

Output will be in: `ConverterApp\dist\PDF-To-Excel-Converter\`

## Step 3: Test the Executable

1. Navigate to `ConverterApp\dist\PDF-To-Excel-Converter\`
2. Run `PDF-To-Excel-Converter.exe`
3. Browser should open automatically at `http://127.0.0.1:5000`
4. Test with a PDF file

## Step 4: Add Icon (Optional)

Create or add an icon file:
- Place `icon.ico` in `ConverterApp\static\`
- Or remove the `SetupIconFile` line from `installer.iss`

## Step 5: Create Installer

1. Open **Inno Setup Compiler**
2. File → Open → select `installer.iss`
3. Build → Compile (or press F9)
4. Installer will be created in `installer_output\` folder

## Output Files

After building:

```
PDF-To-Excel-Converter/
├── ConverterApp/
│   └── dist/
│       └── PDF-To-Excel-Converter/    <- Standalone app folder
│           ├── PDF-To-Excel-Converter.exe
│           ├── templates/
│           ├── static/
│           ├── uploads/
│           ├── logs/
│           └── ... (dependencies)
│
└── installer_output/
    └── PDF-To-Excel-Converter-Setup-0.2.0.exe   <- Installer
```

## Troubleshooting

### "Java not found" error
- Install Java JRE from https://java.com
- Make sure `java` is in system PATH

### PyInstaller errors
- Try: `pip install --upgrade pyinstaller`
- Check for missing hidden imports in the error message

### App doesn't start
- Run from command line to see error messages
- Check `logs\app.log` for details

## Notes

- The app runs a local web server on port 5000
- Browser opens automatically when the app starts
- Console window shows server logs
- Close the console window to stop the server
