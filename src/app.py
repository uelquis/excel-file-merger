from formatter import Formatter
import typer
from excel_merger import ExcelMerger

app = typer.Typer()

@app.command()
def merge(folder_path: str, style_path: str, o: str = "./merged.xlsx", overwrite: bool = False):
    """Merge multiple Excel files into a single file.
    
    - folder_path: path to folder containing .xlsx files
    - style_path: path to YAML file containing formatting configuration
    - o: output file path for merged workbook (default: ./merged.xlsx)
    - overwrite: whether to overwrite existing output file (default: False)"""

    merger = ExcelMerger(folder_path=folder_path, output_path=o, overwrite=overwrite)
    
    merger.merge_files() \
        .format(Formatter(config_path=style_path)) \
        .save()

if __name__ == "__main__":
    app()
