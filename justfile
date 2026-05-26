# use PowerShell instead of sh:
set shell := ["powershell.exe", "-c"]

build:
    uv run pyinstaller --name "excelmerger" src/app.py