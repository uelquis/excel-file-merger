from pathlib import Path
import typer
from excel_merger import ExcelMerger

app = typer.Typer()

@app.command()
def merge(folder_path: str, o: str = "./merged.xlsx"):
    """Merge multiple Excel files into a single file.
    
    - folder_path: path to folder containing .xlsx files
    - o: output file path for merged workbook (default: ./merged.xlsx)"""

    merger = ExcelMerger()
    merged_wb = merger.merge_files(folder_path)

    if(Path(o).exists()):
        typer.confirm(f"File {o} already exists. Do you want to overwrite it?", abort=True)

    Path(o).unlink(missing_ok=True)
    merged_wb.save(o)
    print(f"Merged file saved as {o}")

if __name__ == "__main__":
    app()
