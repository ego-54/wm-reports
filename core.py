"""
Shared logic for report generation and inventory import CSV building.
Works with in-memory file objects (Streamlit uploads) instead of file paths.
"""

import csv
import io
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── helpers ────────────────────────────────────────────────────────────────

def to_int(val):
    v = (val or "").strip()
    if not v or v.lower() == "not stocked":
        return 0
    try:
        return int(float(v))
    except Exception:
        return 0

def to_float(val):
    try:
        return float((val or "").strip())
    except Exception:
        return 0.0


# ─── loaders ────────────────────────────────────────────────────────────────

def load_products_from_files(files):
    """files: list of uploaded file objects (products_export*.csv)"""
    products = {}
    for f in files:
        text = f.read().decode("utf-8")
        reader = csv.DictReader(io.StringIO(text))
        for row in reader:
            h = row.get("Handle", "").strip()
            if not h or h in products:
                continue
            products[h] = {
                "title": row.get("Title", "").strip(),
                "price": row.get("Variant Price", "").strip(),
                "tags":  row.get("Tags", "").strip(),
            }
    return products

def load_inventory_from_files(files):
    """files: list of uploaded file objects (inventory_export*.csv)"""
    inv = {}
    for f in files:
        text = f.read().decode("utf-8")
        reader = csv.DictReader(io.StringIO(text))
        for row in reader:
            h   = row.get("Handle", "").strip()
            sku = row.get("SKU", "").strip()
            if not h:
                continue
            wmc   = to_int(row.get("Wine More Cellars", ""))
            chad  = to_int(row.get("WM Chadstone", ""))
            title = row.get("Title", "").strip()
            if h not in inv:
                inv[h] = {"wmc": 0, "chad": 0, "skus": [], "title": title, "_rows": []}
            inv[h]["wmc"]   += wmc
            inv[h]["chad"]  += chad
            if sku:
                inv[h]["skus"].append(sku)
            inv[h]["_rows"].append(row)
    return inv

def load_supplier_handles_from_csv(file):
    """Returns set of handles from a supplier tagged-products CSV."""
    text = file.read().decode("utf-8")
    handles = set()
    for row in csv.DictReader(io.StringIO(text)):
        h = row.get("Handle", "").strip()
        if h:
            handles.add(h)
    return handles

def load_joval_quantities_xlsx(file):
    """Read joval.xlsx: headers row 8, SKU col A, Qty. In Stock col G."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    qty = {}
    for r in range(9, ws.max_row + 1):
        sku = ws.cell(r, 1).value
        q   = ws.cell(r, 7).value
        if sku and q is not None:
            try:
                qty[str(sku).strip()] = max(0, int(q))
            except (ValueError, TypeError):
                pass
    return qty

def load_qty_csv(file):
    """Generic CSV with SKU and Quantity columns."""
    text   = file.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(text))
    fields = [c.strip().lower() for c in (reader.fieldnames or [])]
    orig   = reader.fieldnames or []
    sku_col = next((orig[i] for i, c in enumerate(fields)
                    if c in ("sku", "item no.", "item no", "item_no")), None)
    qty_col = next((orig[i] for i, c in enumerate(fields)
                    if c in ("quantity", "qty", "qty. in stock", "stock", "on hand")), None)
    if not sku_col or not qty_col:
        return {}, f"Cannot find SKU/Quantity columns (found: {orig})"
    qty = {}
    for row in reader:
        sku = row.get(sku_col, "").strip()
        q   = row.get(qty_col, "").strip()
        if sku and q:
            try:
                qty[sku] = max(0, int(float(q)))
            except (ValueError, TypeError):
                pass
    return qty, None


# ─── report builder ──────────────────────────────────────────────────────────

SUPPLIER_COLORS = {
    "joval":    "FFF2CC",
    "bibendum": "E2EFDA",
    "dws":      "DDEBF7",
    "sss":      "FCE4D6",
}

def build_report(products, inventory, supplier_map):
    """Returns an openpyxl Workbook with Inventory + Summary sheets."""
    all_handles = set(products) | set(inventory)
    rows = []
    for h in all_handles:
        inv  = inventory.get(h, {"wmc": 0, "chad": 0, "skus": [], "title": ""})
        prod = products.get(h, {})
        wmc  = inv["wmc"]
        chad = inv["chad"]
        total = wmc + chad
        skus  = inv.get("skus", [])
        title = prod.get("title") or inv.get("title", "")

        if wmc > 0 and chad > 0:
            location = "Both"
        elif wmc > 0:
            location = "Wine More Cellars"
        elif chad > 0:
            location = "WM Chadstone"
        else:
            location = "—"

        rows.append({
            "SKU":               skus[0] if skus else "",
            "Name":              title,
            "Price":             to_float(prod.get("price", "")),
            "Wine More Cellars": wmc,
            "WM Chadstone":      chad,
            "Total Stock":       total,
            "Location":          location,
            "Supplier":          supplier_map.get(h, "—"),
            "zero":              total == 0,
        })
    rows.sort(key=lambda r: (r["zero"], r["Name"].lower()))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventory"

    col_keys   = ["SKU", "Name", "Price", "Wine More Cellars", "WM Chadstone",
                  "Total Stock", "Location", "Supplier"]
    col_widths = [18, 55, 10, 20, 16, 14, 22, 20]

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    zero_fill   = PatternFill("solid", fgColor="F2F2F2")
    grey_font   = Font(color="909090", size=10)
    normal_font = Font(size=10)
    thin_border = Border(bottom=Side(style="thin", color="D0D0D0"))

    ws.append(col_keys)
    for i in range(1, len(col_keys) + 1):
        c = ws.cell(1, i)
        c.fill      = header_fill
        c.font      = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(col_keys))}1"

    for r in rows:
        ws.append([
            r["SKU"], r["Name"],
            r["Price"] if r["Price"] else "",
            r["Wine More Cellars"], r["WM Chadstone"], r["Total Stock"],
            r["Location"], r["Supplier"],
        ])
        rn  = ws.max_row
        sup = r["Supplier"].split(",")[0].strip().lower()
        sup_color = SUPPLIER_COLORS.get(sup)
        for ci in range(1, len(col_keys) + 1):
            cell = ws.cell(rn, ci)
            cell.border = thin_border
            if r["zero"]:
                cell.fill = zero_fill
                cell.font = grey_font
            else:
                cell.font = normal_font
                if ci == len(col_keys) and sup_color:
                    cell.fill = PatternFill("solid", fgColor=sup_color)
        ws.cell(rn, 3).number_format = '"$"#,##0.00'

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    stocked    = [r for r in rows if not r["zero"]]
    zeroed     = [r for r in rows if r["zero"]]
    total_wmc  = sum(r["Wine More Cellars"] for r in rows)
    total_chad = sum(r["WM Chadstone"] for r in rows)

    sup_stats = defaultdict(lambda: {"variants": 0, "stocked": 0, "zero": 0, "units": 0})
    for r in rows:
        for s in r["Supplier"].split(","):
            s = s.strip() or "No Supplier"
            sup_stats[s]["variants"] += 1
            sup_stats[s]["units"]    += r["Total Stock"]
            sup_stats[s]["stocked" if not r["zero"] else "zero"] += 1

    summary = [
        ["OVERVIEW", ""],
        ["Report generated", datetime.now().strftime("%d %b %Y %H:%M")],
        ["Total variants",   len(rows)],
        ["Stocked",          len(stocked)],
        ["Zero stock",       len(zeroed)],
        ["Total WMC units",  total_wmc],
        ["Total Chadstone",  total_chad],
        ["Combined units",   total_wmc + total_chad],
        [],
        ["SUPPLIER BREAKDOWN", "Variants", "Stocked", "Zero Stock", "Units"],
    ]
    for s, d in sorted(sup_stats.items(), key=lambda x: -x[1]["units"]):
        summary.append([s, d["variants"], d["stocked"], d["zero"], d["units"]])

    for row in summary:
        ws2.append(row)
    ws2["A1"].font = Font(bold=True, size=11)
    ws2["A10"].font = Font(bold=True)
    for col in ["A", "B", "C", "D", "E"]:
        ws2.column_dimensions[col].width = 18

    stats = {
        "total": len(rows),
        "stocked": len(stocked),
        "zero": len(zeroed),
        "wmc_units": total_wmc,
        "chad_units": total_chad,
    }
    return wb, stats


# ─── inventory import CSV builder ────────────────────────────────────────────

SHOPIFY_COLS = [
    "Handle", "Title", "Option1 Name", "Option1 Value",
    "Option2 Name", "Option2 Value", "Option3 Name", "Option3 Value",
    "SKU", "HS Code", "COO", "Location", "Bin name",
    "Incoming (not editable)", "Unavailable (not editable)",
    "Committed (not editable)", "Available (not editable)", "On hand (new)",
]

def build_inventory_import(inv_rows_raw, supplier_qty, location="Wine More Cellars"):
    """
    inv_rows_raw: flat list of raw CSV row dicts from inventory export
    supplier_qty: {sku: quantity}
    Returns (csv_bytes, stats_dict, unmatched_skus)
    """
    matched   = []
    skipped   = 0
    not_found = set(supplier_qty.keys())

    for row in inv_rows_raw:
        sku = row.get("SKU", "").strip()
        if sku not in supplier_qty:
            skipped += 1
            continue
        not_found.discard(sku)
        out = {col: "" for col in SHOPIFY_COLS}
        out.update({
            "Handle":                     row.get("Handle", ""),
            "Title":                      row.get("Title", ""),
            "Option1 Name":               row.get("Option1 Name", ""),
            "Option1 Value":              row.get("Option1 Value", ""),
            "Option2 Name":               row.get("Option2 Name", ""),
            "Option2 Value":              row.get("Option2 Value", ""),
            "Option3 Name":               row.get("Option3 Name", ""),
            "Option3 Value":              row.get("Option3 Value", ""),
            "SKU":                        sku,
            "Location":                   location,
            "Incoming (not editable)":    "0",
            "Unavailable (not editable)": "0",
            "Committed (not editable)":   "0",
            "Available (not editable)":   "0",
            "On hand (new)":              str(supplier_qty[sku]),
        })
        matched.append(out)

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=SHOPIFY_COLS)
    writer.writeheader()
    writer.writerows(matched)

    stats = {
        "matched": len(matched),
        "skipped": skipped,
        "not_found": len(not_found),
    }
    return buf.getvalue().encode("utf-8"), stats, sorted(not_found)


def load_flat_inventory_rows(files):
    """Returns flat list of all raw CSV rows (not aggregated by handle)."""
    rows = []
    for f in files:
        text = f.read().decode("utf-8")
        for row in csv.DictReader(io.StringIO(text)):
            if row.get("SKU", "").strip():
                rows.append(row)
    return rows
