from pathlib import Path

from config import file_name, sheet_names
from readers import WorkbookReader

if __name__ == "__main__":
    path_excel = Path("data") / file_name
    workbook_reader = WorkbookReader(path_excel, sheet_names)
