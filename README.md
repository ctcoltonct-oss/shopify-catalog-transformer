# Shopify catalog transformer

Config-driven Python pipeline that turns supplier CSV or XLSX files into a Shopify product import CSV, plus a rejected-row file and a QA report.

This is an AI-assisted practice project. It shows workflow design, validation, and deterministic tests. It is not a store integration, and it has not been imported into a live Shopify admin.

## What problem this solves

Supplier sheets do not match Shopify's import columns. Prices are messy, titles go missing, and the same SKU can appear twice. Loading that file as-is creates bad products. This tool maps the sheet, groups variants, and refuses rows it cannot price or title safely.

## What goes in

- One or more `.csv` or `.xlsx` files.
- `mapping.json` — optional per-file column map. If a file has no entry, known headers are matched exactly (`sku`, `item #`, `item number`, `product name`, `upc`, `colour`, and so on). A bare header named `Item` is not guessed. Map it explicitly.
- `rules.json` — vendor default, publish/status flags, and an optional cost markup used only when the row has no sell price.

Synthetic inputs are in [fixtures](fixtures).

## What comes out

Each run writes four files:

| File | Contents |
|---|---|
| `shopify_import.csv` | Classic Products → Import headers (`Handle`, `Title`, `Body (HTML)`, `Variant SKU`, `Variant Price`, …) |
| `rejected_rows.csv` | Source file, source row, SKU, title, reasons |
| `QA_report.csv` | Row counts, product count, duplicate-SKU count, header check |
| `transformation_summary.txt` | One-screen totals |

Nothing is filled in from outside the source row, except the vendor default and the markup rule when those are configured.

## How validation works

A row is rejected when:

- It has no title and no parent/style key (`missing_title`).
- The sell price is blank, negative, or not a number.
- A cost used as the price source cannot be parsed (`unparseable_price:...`).
- Quantity is present but not a number.
- The same SKU was already accepted (`duplicate_sku`). The first copy is kept.

Variant rows that share a parent/style get one handle. Later rows of that handle leave Title, body, vendor, and status blank, which matches Shopify's multi-row variant layout.

## Before and after

Input, three tee rows in [fixtures/clean_supplier.csv](fixtures/clean_supplier.csv):

```
SKU,Product Name,Wholesale,Retail,Color,Size
TEE-BLK-S,Classic Tee,8.5,24.99,Black,S
TEE-BLK-M,Classic Tee,8.5,24.99,Black,M
TEE-WHT-S,Classic Tee,8.5,24.99,White,S
```

Output (trimmed):

| Handle | Title | Option1 | Option2 | Variant SKU | Variant Price |
|---|---|---|---|---|---|
| classic-tee | Classic Tee | Color / Black | Size / S | TEE-BLK-S | 24.99 |
| classic-tee |  | Color / Black | Size / M | TEE-BLK-M | 24.99 |
| classic-tee |  | Color / White | Size / S | TEE-WHT-S | 24.99 |

The retail column is the sell price. Wholesale is cost. Markup is not applied when a sell price is already present.

Messy rows are in [fixtures/messy_rows.csv](fixtures/messy_rows.csv). `python run_demo.py` writes `fixtures/messy_supplier.xlsx` from that file, then converts it. The workbook is not stored in git because it is binary. Five rows, no sell-price column. Cost is marked up by 2.2 from [fixtures/rules.json](fixtures/rules.json):

| Result | What happened |
|---|---|
| Desk Lamp, first SKU | Accepted at 26.40 (12.00 × 2.2) |
| Same SKU again | Rejected: `duplicate_sku;duplicate_barcode` |
| Blank title | Rejected: `missing_title;blank_price` |
| Cost value `eighteen` | Rejected: `unparseable_price:eighteen` |
| Clamp Lamp | Accepted at 20.90 (9.50 × 2.2) |

Full files are in [examples](examples).

## How to run

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m pytest
python run_demo.py
```

One catalog:

```bash
python src/transform.py out fixtures/mapping.json fixtures/rules.json fixtures/clean_supplier.csv
```

## What this project practices

- Reading CSV and Excel without changing the source.
- Keeping mapping and business rules out of the code.
- Rejecting bad rows instead of guessing.
- Checking the result with fixed synthetic files, not a one-off manual run.

## Limitations

- Local checks only. No Shopify store was created, and this CSV was not uploaded through Shopify admin.
- CSV and XLSX only. Legacy `.xls` is not supported.
- A repeated barcode is named in the reject reason when that row is already a duplicate SKU. A repeated barcode with unique SKUs is not rejected by itself.
- Image cells must already be public URLs. The tool does not download or invent images.
- At most three options (color, size, and one extra).
- Markup never replaces a sell price that parsed successfully.

## Tests

`python -m pytest` runs five cases:

| Case | Accepted | Rejected | Products |
|---|---:|---:|---:|
| Clean tee | 3 | 0 | 1 |
| Messy workbook | 2 | 3 | 2 |
| Hoodie size/color | 4 | 0 | 1 |
| Two supplier files | 5 | 0 | 3 |
| Failure file | 0 | 3 | 0 |
