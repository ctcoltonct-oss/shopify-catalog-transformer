#!/usr/bin/env python3
"""Config-driven supplier catalog → Shopify product CSV pack.

Outputs classic Shopify import headers (Handle, Title, Body (HTML), Variant SKU,
Variant Price, ...) which the Products → Import flow still accepts.

Does not invent prices, barcodes, images, or inventory.
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

CLASSIC_HEADERS = [
    "Handle",
    "Title",
    "Body (HTML)",
    "Vendor",
    "Product Category",
    "Type",
    "Tags",
    "Published",
    "Option1 Name",
    "Option1 Value",
    "Option2 Name",
    "Option2 Value",
    "Option3 Name",
    "Option3 Value",
    "Variant SKU",
    "Variant Grams",
    "Variant Inventory Tracker",
    "Variant Inventory Qty",
    "Variant Inventory Policy",
    "Variant Fulfillment Service",
    "Variant Price",
    "Variant Compare At Price",
    "Variant Requires Shipping",
    "Variant Taxable",
    "Variant Barcode",
    "Image Src",
    "Image Position",
    "Image Alt Text",
    "Gift Card",
    "SEO Title",
    "SEO Description",
    "Variant Image",
    "Variant Weight Unit",
    "Cost per item",
    "Status",
]

HEADER_ALIASES = {
    "sku": "sku",
    "item #": "sku",
    "item number": "sku",
    "product sku": "sku",
    "style": "parent_sku",
    "style #": "parent_sku",
    "parent sku": "parent_sku",
    "parent": "parent_sku",
    "product id": "parent_sku",
    "title": "title",
    "product name": "title",
    "name": "title",
    "description": "body",
    "product description": "body",
    "desc": "body",
    "long description": "body",
    "vendor": "vendor",
    "brand": "vendor",
    "manufacturer": "vendor",
    "type": "type",
    "product type": "type",
    "category": "type",
    "price": "price",
    "retail": "price",
    "retail price": "price",
    "map": "compare_at",
    "msrp": "compare_at",
    "compare at": "compare_at",
    "compare-at": "compare_at",
    "cost": "cost",
    "wholesale": "cost",
    "unit cost": "cost",
    "barcode": "barcode",
    "upc": "barcode",
    "ean": "barcode",
    "gtin": "barcode",
    "color": "option_color",
    "colour": "option_color",
    "size": "option_size",
    "variant": "option_variant",
    "image": "image",
    "image url": "image",
    "image src": "image",
    "photo": "image",
    "qty": "qty",
    "quantity": "qty",
    "stock": "qty",
    "inventory": "qty",
    "weight": "grams",
    "weight grams": "grams",
    "grams": "grams",
    "tags": "tags",
}


def slug_handle(text: str) -> str:
    s = str(text).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s[:80] or "untitled"


def parse_money(val) -> tuple[str | None, str | None]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None, "blank_price"
    s = str(val).strip()
    if s == "" or s.lower() in {"n/a", "na", "none", "-"}:
        return None, "blank_price"
    s = s.replace("$", "").replace("£", "").replace(",", "").replace("USD", "").strip()
    try:
        n = float(s)
    except ValueError:
        return None, f"unparseable_price:{val}"
    if n < 0:
        return None, "negative_price"
    return f"{n:.2f}", None


def normalize_header(h) -> str:
    return re.sub(r"\s+", " ", str(h).strip().lower())


def auto_map(columns: list[str]) -> dict[str, str]:
    mapping = {}
    used = set()
    for col in columns:
        key = normalize_header(col)
        field = HEADER_ALIASES.get(key)
        if field and field not in used:
            mapping[col] = field
            used.add(field)
    return mapping


def read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(path, dtype=object, engine="openpyxl")
    if suffix == ".csv":
        return pd.read_csv(path, dtype=object, encoding="utf-8-sig")
    raise ValueError(f"unsupported file type {suffix}; use .csv or .xlsx")


def apply_markup(cost: str | None, markup: dict | None) -> str | None:
    if not markup or not cost:
        return None
    try:
        c = float(cost)
    except ValueError:
        return None
    kind = markup.get("type")
    amt = float(markup.get("amount", 0))
    if kind == "multiplier":
        return f"{c * amt:.2f}"
    if kind == "percent":
        return f"{c * (1 + amt / 100):.2f}"
    if kind == "fixed":
        return f"{c + amt:.2f}"
    return None


def process(sources: list[Path], mapping_path: Path, rules_path: Path, out_dir: Path) -> dict:
    mapping_cfg = json.loads(mapping_path.read_text()) if mapping_path.exists() else {}
    rules = json.loads(rules_path.read_text()) if rules_path.exists() else {}
    vendor_default = rules.get("vendor", "")
    published = str(rules.get("published", True)).lower()
    status = rules.get("status", "active")
    inventory_policy = rules.get("inventory_policy", "deny")
    fulfillment = rules.get("fulfillment_service", "manual")
    taxable = str(rules.get("taxable", True)).lower()
    shipping = str(rules.get("requires_shipping", True)).lower()
    markup = rules.get("markup")
    group_by = rules.get("group_by", "parent_or_title")

    frames = []
    for src in sources:
        df = read_table(src)
        df.columns = [str(c).strip() for c in df.columns]
        file_map = mapping_cfg.get(src.name) or mapping_cfg.get("default") or auto_map(list(df.columns))
        frames.append((src.name, df, file_map))

    shopify_rows = []
    rejected = []
    seen_sku = Counter()
    seen_barcode = Counter()
    handle_titles = {}

    for src_name, df, fmap in frames:
        inv = {v: k for k, v in fmap.items()}

        def col(field):
            c = inv.get(field)
            return None if c is None else df[c]

        for idx, row in df.iterrows():
            src_row = int(idx) + 2
            raw = {field: (row[inv[field]] if field in inv else None) for field in set(fmap.values())}

            title = raw.get("title")
            title = "" if title is None or (isinstance(title, float) and pd.isna(title)) else str(title).strip()
            sku_raw = raw.get("sku")
            sku = "" if sku_raw is None or (isinstance(sku_raw, float) and pd.isna(sku_raw)) else str(sku_raw).strip()
            parent = raw.get("parent_sku")
            parent = "" if parent is None or (isinstance(parent, float) and pd.isna(parent)) else str(parent).strip()

            price, price_err = parse_money(raw.get("price"))
            if price is None and markup:
                cost_val, cost_err = parse_money(raw.get("cost"))
                price = apply_markup(cost_val, markup)
                if price:
                    price_err = None
                elif cost_err and cost_err != "blank_price":
                    price_err = cost_err

            compare, _ = parse_money(raw.get("compare_at"))
            cost, _ = parse_money(raw.get("cost"))

            reasons = []
            if not title and not parent:
                reasons.append("missing_title")
            if price is None:
                reasons.append(price_err or "missing_price")
            if sku:
                seen_sku[sku] += 1

            barcode = raw.get("barcode")
            barcode = "" if barcode is None or (isinstance(barcode, float) and pd.isna(barcode)) else str(barcode).strip()
            if barcode:
                seen_barcode[barcode] += 1

            if reasons:
                rejected.append(
                    {
                        "source_file": src_name,
                        "source_row": src_row,
                        "sku": sku,
                        "title": title,
                        "reasons": ";".join(reasons),
                    }
                )
                continue

            if group_by == "parent_or_title" and parent:
                handle = slug_handle(parent)
                product_title = title or parent
            else:
                handle = slug_handle(title)
                product_title = title

            if handle in handle_titles and handle_titles[handle] != product_title:
                handle = slug_handle(f"{product_title}-{sku or src_row}")
            handle_titles.setdefault(handle, product_title)

            opt1_name = opt1_val = opt2_name = opt2_val = opt3_name = opt3_val = ""
            color = raw.get("option_color")
            size = raw.get("option_size")
            variant = raw.get("option_variant")
            opts = []
            if color is not None and not (isinstance(color, float) and pd.isna(color)) and str(color).strip():
                opts.append(("Color", str(color).strip()))
            if size is not None and not (isinstance(size, float) and pd.isna(size)) and str(size).strip():
                opts.append(("Size", str(size).strip()))
            if variant is not None and not (isinstance(variant, float) and pd.isna(variant)) and str(variant).strip() and not opts:
                opts.append(("Title", str(variant).strip()))
            if not opts:
                opts.append(("Title", "Default Title"))
            if len(opts) > 0:
                opt1_name, opt1_val = opts[0]
            if len(opts) > 1:
                opt2_name, opt2_val = opts[1]
            if len(opts) > 2:
                opt3_name, opt3_val = opts[2]

            vendor = raw.get("vendor")
            vendor = vendor_default if vendor is None or (isinstance(vendor, float) and pd.isna(vendor)) or not str(vendor).strip() else str(vendor).strip()
            body = raw.get("body")
            body = "" if body is None or (isinstance(body, float) and pd.isna(body)) else str(body).strip()
            ptype = raw.get("type")
            ptype = "" if ptype is None or (isinstance(ptype, float) and pd.isna(ptype)) else str(ptype).strip()
            tags = raw.get("tags")
            tags = "" if tags is None or (isinstance(tags, float) and pd.isna(tags)) else str(tags).strip()
            image = raw.get("image")
            image = "" if image is None or (isinstance(image, float) and pd.isna(image)) else str(image).strip()
            qty = raw.get("qty")
            qty_s = ""
            if qty is not None and not (isinstance(qty, float) and pd.isna(qty)) and str(qty).strip() != "":
                try:
                    qty_s = str(int(float(str(qty).replace(",", ""))))
                except ValueError:
                    rejected.append(
                        {
                            "source_file": src_name,
                            "source_row": src_row,
                            "sku": sku,
                            "title": title,
                            "reasons": f"bad_qty:{qty}",
                        }
                    )
                    continue
            grams = raw.get("grams")
            grams_s = ""
            if grams is not None and not (isinstance(grams, float) and pd.isna(grams)) and str(grams).strip() != "":
                try:
                    grams_s = str(int(float(str(grams).replace(",", ""))))
                except ValueError:
                    grams_s = ""

            shopify_rows.append(
                {
                    "Handle": handle,
                    "Title": product_title,
                    "Body (HTML)": body,
                    "Vendor": vendor,
                    "Product Category": "",
                    "Type": ptype,
                    "Tags": tags,
                    "Published": published,
                    "Option1 Name": opt1_name,
                    "Option1 Value": opt1_val,
                    "Option2 Name": opt2_name,
                    "Option2 Value": opt2_val,
                    "Option3 Name": opt3_name,
                    "Option3 Value": opt3_val,
                    "Variant SKU": sku,
                    "Variant Grams": grams_s,
                    "Variant Inventory Tracker": "shopify" if qty_s != "" else "",
                    "Variant Inventory Qty": qty_s,
                    "Variant Inventory Policy": inventory_policy,
                    "Variant Fulfillment Service": fulfillment,
                    "Variant Price": price,
                    "Variant Compare At Price": compare or "",
                    "Variant Requires Shipping": shipping,
                    "Variant Taxable": taxable,
                    "Variant Barcode": barcode,
                    "Image Src": image,
                    "Image Position": "1" if image else "",
                    "Image Alt Text": product_title if image else "",
                    "Gift Card": "false",
                    "SEO Title": "",
                    "SEO Description": "",
                    "Variant Image": "",
                    "Variant Weight Unit": "g" if grams_s else "",
                    "Cost per item": cost or "",
                    "Status": status,
                    "_source_file": src_name,
                    "_source_row": src_row,
                }
            )

    # post-pass duplicate SKU / barcode flags (keep first, reject later dupes if rule set)
    sku_counts = Counter(r["Variant SKU"] for r in shopify_rows if r["Variant SKU"])
    barcode_counts = Counter(r["Variant Barcode"] for r in shopify_rows if r["Variant Barcode"])
    kept = []
    seen_live_sku = set()
    for r in shopify_rows:
        extra = []
        if r["Variant SKU"] and sku_counts[r["Variant SKU"]] > 1:
            if r["Variant SKU"] in seen_live_sku:
                extra.append("duplicate_sku")
            else:
                seen_live_sku.add(r["Variant SKU"])
                extra.append("duplicate_sku_kept_first")
        if r["Variant Barcode"] and barcode_counts[r["Variant Barcode"]] > 1:
            extra.append("duplicate_barcode")
        if "duplicate_sku" in extra:
            rejected.append(
                {
                    "source_file": r["_source_file"],
                    "source_row": r["_source_row"],
                    "sku": r["Variant SKU"],
                    "title": r["Title"],
                    "reasons": ";".join(extra),
                }
            )
            continue
        kept.append(r)

    # blank Title on subsequent variant rows of same handle
    seen_handles = set()
    out_rows = []
    for r in kept:
        row = {h: r.get(h, "") for h in CLASSIC_HEADERS}
        if r["Handle"] in seen_handles:
            row["Title"] = ""
            row["Body (HTML)"] = ""
            row["Vendor"] = ""
            row["Type"] = ""
            row["Tags"] = ""
            row["Published"] = ""
            row["Status"] = ""
        else:
            seen_handles.add(r["Handle"])
        out_rows.append(row)

    out_dir.mkdir(parents=True, exist_ok=True)
    import_path = out_dir / "shopify_import.csv"
    reject_path = out_dir / "rejected_rows.csv"
    qa_path = out_dir / "QA_report.csv"
    summary_path = out_dir / "transformation_summary.txt"

    pd.DataFrame(out_rows, columns=CLASSIC_HEADERS).to_csv(import_path, index=False, encoding="utf-8")
    pd.DataFrame(rejected).to_csv(reject_path, index=False, encoding="utf-8")

    qa = []
    handles = [r["Handle"] for r in out_rows]
    qa.append({"check": "output_rows", "value": len(out_rows)})
    qa.append({"check": "rejected_rows", "value": len(rejected)})
    qa.append({"check": "unique_handles", "value": len(set(handles))})
    qa.append({"check": "duplicate_skus_detected", "value": sum(1 for v in sku_counts.values() if v > 1)})
    qa.append({"check": "headers_ok", "value": list(pd.read_csv(import_path, nrows=0).columns) == CLASSIC_HEADERS})
    pd.DataFrame(qa).to_csv(qa_path, index=False)

    summary = (
        f"sources={','.join(s.name for s in sources)}\n"
        f"accepted_rows={len(out_rows)}\n"
        f"rejected_rows={len(rejected)}\n"
        f"products={len(set(handles))}\n"
        f"markup={markup}\n"
        f"vendor_default={vendor_default}\n"
        f"note=Images must already be public URLs. No values invented.\n"
    )
    summary_path.write_text(summary)
    return {
        "accepted": len(out_rows),
        "rejected": len(rejected),
        "products": len(set(handles)),
        "import_path": str(import_path),
    }


def main():
    # usage: engine.py out_dir mapping.json rules.json file1 [file2...]
    out_dir = Path(sys.argv[1])
    mapping = Path(sys.argv[2])
    rules = Path(sys.argv[3])
    files = [Path(p) for p in sys.argv[4:]]
    result = process(files, mapping, rules, out_dir)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
