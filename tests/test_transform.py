import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from sample_workbook import csv_to_xlsx  # noqa: E402
from transform import CLASSIC_HEADERS, process  # noqa: E402

FIXTURES = ROOT / "fixtures"
MAPPING = FIXTURES / "mapping.json"
RULES = FIXTURES / "rules.json"


def run(tmp_path, files):
    return process([FIXTURES / name for name in files], MAPPING, RULES, tmp_path)


def test_clean_groups_one_product(tmp_path):
    result = run(tmp_path, ["clean_supplier.csv"])
    assert result["accepted"] == 3
    assert result["rejected"] == 0
    assert result["products"] == 1
    frame = pd.read_csv(tmp_path / "shopify_import.csv", dtype=str).fillna("")
    assert list(frame.columns) == CLASSIC_HEADERS
    assert frame.loc[0, "Title"] == "Classic Tee"
    assert frame.loc[1, "Title"] == ""
    assert frame.loc[0, "Variant Price"] == "24.99"
    assert frame.loc[0, "Handle"] == frame.loc[2, "Handle"] == "classic-tee"


def test_messy_rejects_unsafe_rows(tmp_path):
    xlsx = tmp_path / "messy_supplier.xlsx"
    csv_to_xlsx(FIXTURES / "messy_rows.csv", xlsx)
    out = tmp_path / "out"
    result = process([xlsx], MAPPING, RULES, out)
    assert (result["accepted"], result["rejected"], result["products"]) == (2, 3, 2)
    rejected = pd.read_csv(out / "rejected_rows.csv", dtype=str).fillna("")
    reasons = ";".join(rejected["reasons"].tolist())
    assert "missing_title" in reasons
    assert "unparseable_price:eighteen" in reasons
    assert "duplicate_sku" in reasons
    prices = pd.read_csv(out / "shopify_import.csv", dtype=str)["Variant Price"].tolist()
    assert prices == ["26.40", "20.90"]


def test_variants_share_handle(tmp_path):
    result = run(tmp_path, ["variants_supplier.csv"])
    assert (result["accepted"], result["rejected"], result["products"]) == (4, 0, 1)
    frame = pd.read_csv(tmp_path / "shopify_import.csv", dtype=str).fillna("")
    assert set(frame["Handle"]) == {"hood-100"}
    assert list(frame["Title"]) == ["Zip Hoodie", "", "", ""]


def test_two_supplier_files(tmp_path):
    result = run(tmp_path, ["clean_supplier.csv", "supplier_b.csv"])
    assert (result["accepted"], result["rejected"], result["products"]) == (5, 0, 3)


def test_failure_file_invents_nothing(tmp_path):
    result = run(tmp_path, ["failure_supplier.csv"])
    assert (result["accepted"], result["rejected"], result["products"]) == (0, 3, 0)
    frame = pd.read_csv(tmp_path / "shopify_import.csv")
    assert frame.empty
    rejected = pd.read_csv(tmp_path / "rejected_rows.csv", dtype=str)
    assert set(rejected["reasons"]) == {
        "missing_title;blank_price",
        "negative_price",
        "unparseable_price:ask",
    }


def test_plain_item_header_is_not_auto_mapped(tmp_path):
    source = tmp_path / "plain_item.csv"
    source.write_text("Item,Retail\nLamp,10\n", encoding="utf-8")
    mapping = tmp_path / "mapping.json"
    rules = tmp_path / "rules.json"
    mapping.write_text("{}\n", encoding="utf-8")
    rules.write_text(RULES.read_text(), encoding="utf-8")
    out = tmp_path / "out"
    result = process([source], mapping, rules, out)
    assert result["accepted"] == 0
    assert result["rejected"] == 1
    rejected = pd.read_csv(out / "rejected_rows.csv", dtype=str)
    assert "missing_title" in rejected.loc[0, "reasons"]


def test_xls_is_rejected(tmp_path):
    source = tmp_path / "old.xls"
    source.write_text("not a workbook", encoding="utf-8")
    try:
        process([source], MAPPING, RULES, tmp_path / "out")
    except ValueError as exc:
        assert ".xls" in str(exc)
        assert ".xlsx" in str(exc)
    else:
        raise AssertionError("legacy .xls should be rejected")
