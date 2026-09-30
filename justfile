# use PowerShell instead of sh:
set shell := ["powershell.exe", "-c"]

build:
    uv run pyinstaller excelmerger.spec --noconfirm