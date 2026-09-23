#!/usr/bin/env python3
"""Run the included synthetic catalogs and write example outputs."""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from sample_workbook import csv_to_xlsx  # noqa: E402
from transform import process  # noqa: E402

FIXTURES = ROOT / "fixtures"
EXAMPLES = ROOT / "examples"
MAPPING = FIXTURES / "mapping.json"
RULES = FIXTURES / "rules.json"

CASES = [
    ("clean", [FIXTURES / "clean_supplier.csv"]),
    ("messy", [FIXTURES / "messy_supplier.xlsx"]),
    ("variants", [FIXTURES / "variants_supplier.csv"]),
    ("multi", [FIXTURES / "clean_supplier.csv", FIXTURES / "supplier_b.csv"]),
    ("failure", [FIXTURES / "failure_supplier.csv"]),
]


def main() -> None:
    csv_to_xlsx(FIXTURES / "messy_rows.csv", FIXTURES / "messy_supplier.xlsx")
    results = []
    for name, files in CASES:
        out = EXAMPLES / name
        result = process(files, MAPPING, RULES, out)
        result["case"] = name
        result["import_path"] = f"examples/{name}/shopify_import.csv"
        results.append(result)
        print(f"{name}: accepted={result['accepted']} rejected={result['rejected']} products={result['products']}")
    (EXAMPLES / "demo_results.json").write_text(json.dumps(results, indent=2) + "\n")


if __name__ == "__main__":
    main()
