"""Build the messy supplier workbook from the text fixture."""

from pathlib import Path

import pandas as pd


def csv_to_xlsx(csv_path: Path, xlsx_path: Path) -> Path:
    frame = pd.read_csv(csv_path, dtype=object, keep_default_na=False)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(xlsx_path, index=False, engine="openpyxl")
    return xlsx_path
