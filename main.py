from pathlib import Path

import matplotlib.pyplot as plt

from config import file_name, sheet_name
from readers import WorksheetReader

if __name__ == "__main__":
    path_excel = Path("data") / file_name
    cpu_worksheet_reader = WorksheetReader(sheet_name, path_excel)
    table_names_string = "\n".join(cpu_worksheet_reader.data_frames.keys())
    print(f'Found the following tables on worksheet "{sheet_name}":\n\n{table_names_string}')
    for df_name, df in cpu_worksheet_reader.data_frames.items():
        df.plot(title=df_name)
        plt.savefig(f"data/plots/{df_name}")
