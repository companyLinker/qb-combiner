"""QuickBooks export parser. Supports two variants:

  VARIANT A (single-year): one amount column on the right. Header on row 4 names
    the period (e.g., "Jan - Dec 25" or "Dec 31, 25"). Used by AP NE filings.

  VARIANT B (multi-year side-by-side): three amount columns on the right —
    Current Year, Prior Year, $ Change. Header on row 5 names them.
    Used by Popeyes IL filings ("Dec 31, 25" | "Dec 31, 24" | "$ Change").

Auto-detected per file. Each parsed row carries amount_cy, amount_py, change.
Downstream code reads amount_cy by default; legacy `amount` is an alias for CY.
"""

import io
import os
import re
import openpyxl


def entity_from_filename(fn):
    base = os.path.basename(fn)
    m = re.match(r"^\d+\.\s*(.+?)(?:\s+Balance Sheet.*|\s+\d{4}-\d{4}|\.xlsx).*$", base, re.IGNORECASE)
    if m:
        return m.group(1).strip().rstrip(".")
    m = re.match(r"^\d+\.\s*(.+)\.xlsx$", base)
    return m.group(1).strip() if m else base.replace(".xlsx", "")


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _detect_amount_columns(ws):
    period_pat = re.compile(r"(Jan\s*-\s*Dec|Dec\s*31|Jan\s*\d|Period|Change)", re.IGNORECASE)
    for r in range(3, 8):
        if r > ws.max_row:
            break
        hits = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and period_pat.search(v):
                hits.append((c, v.strip()))
        if hits:
            if len(hits) >= 3:
                cols = sorted([c for c, _ in hits])
                return ("triple", cols[0], cols[1], cols[2], r, r + 1)
            return ("single", hits[0][0], None, None, r, r + 1)
    return ("single", ws.max_column, None, None, 4, 5)


def parse_sheet(ws):
    variant, col_cy, col_py, col_change, header_row, data_start = _detect_amount_columns(ws)
    period_cy = ws.cell(row=header_row, column=col_cy).value if col_cy else None
    period_py = ws.cell(row=header_row, column=col_py).value if col_py else None
    period_cy = period_cy.strip() if isinstance(period_cy, str) else period_cy
    period_py = period_py.strip() if isinstance(period_py, str) else period_py

    rows = []
    hierarchy = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx < data_start:
            continue

        amt_cy = amt_py = change = None
        if variant == "triple":
            if col_cy - 1 < len(row): amt_cy = row[col_cy - 1]
            if col_py - 1 < len(row): amt_py = row[col_py - 1]
            if col_change - 1 < len(row): change = row[col_change - 1]
            if not is_number(amt_cy): amt_cy = None
            if not is_number(amt_py): amt_py = None
            if not is_number(change): change = None
            label_search_end = min(col_cy - 1, col_py - 1, col_change - 1)
        else:
            label_search_end = len(row)
            for ci in range(len(row) - 1, -1, -1):
                if is_number(row[ci]):
                    amt_cy = row[ci]
                    label_search_end = ci
                    break

        lbl_idx, lbl = None, None
        for ci in range(label_search_end - 1, -1, -1):
            v = row[ci]
            if isinstance(v, str) and v.strip():
                lbl_idx, lbl = ci, v.strip()
                break

        if lbl is None:
            continue

        lbl_lower = lbl.lower()
        is_total = (lbl_lower.startswith("total ")
                    or lbl_lower in ("net income", "net ordinary income", "net other income")
                    or "total" in lbl_lower[:6])
        is_section_only = (amt_cy is None and amt_py is None)

        while hierarchy and hierarchy[-1][0] >= lbl_idx:
            hierarchy.pop()

        parents = [p[1] for p in hierarchy]
        rows.append({
            "indent": lbl_idx,
            "label": lbl,
            "amount": amt_cy,
            "amount_cy": amt_cy,
            "amount_py": amt_py,
            "change": change,
            "is_total": is_total,
            "is_section_only": is_section_only,
            "breadcrumb": " > ".join(parents),
            "parents": parents,
            "row_idx": r_idx,
        })

        if not is_total:
            hierarchy.append((lbl_idx, lbl))

    return variant, period_cy, period_py, rows


def parse_uploaded_files(uploaded_files):
    all_data = {}
    for uf in uploaded_files:
        if hasattr(uf, "name"):
            fname = uf.name
            file_bytes = uf.getvalue() if hasattr(uf, "getvalue") else uf.read()
        else:
            fname, file_bytes = uf

        entity = entity_from_filename(fname)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        pnl_sheet = bs_sheet = None
        for sn in wb.sheetnames:
            sn_l = sn.lower()
            if any(k in sn_l for k in ["profit", "income statement", "income state"]):
                pnl_sheet = sn
            elif "balance" in sn_l:
                bs_sheet = sn

        if not pnl_sheet or not bs_sheet:
            all_data[entity] = {
                "file": fname,
                "error": "Missing P&L or BS sheet. Found: " + str(wb.sheetnames),
                "pnl_rows": [], "bs_rows": [],
                "pnl_period_cy": None, "pnl_period_py": None,
                "bs_period_cy": None,  "bs_period_py": None,
                "variant": "unknown",
            }
            continue

        pnl_variant, pnl_cy, pnl_py, pnl_rows = parse_sheet(wb[pnl_sheet])
        bs_variant,  bs_cy,  bs_py,  bs_rows  = parse_sheet(wb[bs_sheet])
        variant = "triple" if "triple" in (pnl_variant, bs_variant) else "single"

        all_data[entity] = {
            "file": fname,
            "variant": variant,
            "pnl_period_cy": pnl_cy, "pnl_period_py": pnl_py,
            "bs_period_cy":  bs_cy,  "bs_period_py":  bs_py,
            "pnl_period": pnl_cy,
            "bs_period":  bs_cy,
            "pnl_rows": pnl_rows,
            "bs_rows":  bs_rows,
        }
    return all_data
