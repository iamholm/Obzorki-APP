# Obzorki APP

Desktop application for generating overview reports from UII source lists.

## Requirements

- Windows
- Python 3.11+ (tested on 3.14)
- Install dependencies:

```powershell
python -m pip install PySide6 python-docx openpyxl
```

## Run

```powershell
python app_uii.py
```

## Build portable

```powershell
python -m PyInstaller --noconfirm --clean --windowed --name AutoObzorki --add-data "Test.docx;." app_uii.py
```
