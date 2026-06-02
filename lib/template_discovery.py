"""Discover the structure of an arbitrary target combination workbook at runtime.

Why: templates evolve. Sheet names get typos, header rows shift, entity columns
move, new years get added, line items get renamed. Hard-coding row numbers and
column letters breaks every time. Discovery introspects the actual workbook
each run and produces a structural plan the builder consumes.
"""

import re
import sys
from typing import List, Optional, Tuple


YEAR_RE = re.compile(r"(20\d{2})")
_BS_RE = re.compile(r"(balance\s*sheet|balacnce|\bbs\b|\bbs\s*\d|\bbs(?:20\d{2})?$|^bs(?:20\d{2})?$)", re.IGNORECASE)
_IS_RE = re.compile(r"(income\s*statement|income\s*state|profit|p\s*&\s*l|\bis\b|\bis\s*\d|\bis(?:20\d{2})?$|^is(?:20\d{2})?$|p&l)", re.IGNORECASE)
SUMMARY_KW = {
    "total": "total", "subtotal": "total",
    "adj": "adj", "adjustment": "adj", "elimination": "adj",
    "parent": "parent",
    "grand": "final",
}


def norm_label(s) -> str:
    if not isinstance(s, str):
        return ""
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s)


def _classify_sheet(name: str):
    n = name.strip()
    yr_m = YEAR_RE.search(n)
    year = int(yr_m.group(1)) if yr_m else None
    # IS first because "P&l" with year 2024 could otherwise be mis-matched
    if _IS_RE.search(n):
        return "IS", year
    if _BS_RE.search(n):
        return "BS", year
    return "OTHER", year


def _find_header_row(ws, max_scan: int = 15):
    best_row, best_count = 1, 0
    for r in range(1, min(max_scan, ws.max_row) + 1):
        cnt = 0
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() and not v.startswith("="):
                cnt += 1
        if cnt > best_count:
            best_count, best_row = cnt, r
    return best_row


def _classify_summary_col(header_text: str) -> Optional[str]:
    n = norm_label(header_text)
    for kw, role in SUMMARY_KW.items():
        if kw in n:
            return role
    return None


_RANGE_RE = re.compile(r"([A-Z]+)(\d+):([A-Z]+)(\d+)")
_CELL_RE = re.compile(r"([A-Z]+)(\d+)")


def _formula_spans_multiple_rows(formula: str, current_row: int) -> bool:
    """True iff this formula references rows other than the current row.
    Same-row range like =SUM(B10:G10) on row 10 → False (it's a per-row total).
    Multi-row range like =SUM(B10:B20) → True (it's a subtotal across rows)."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    for m in _RANGE_RE.finditer(formula):
        r1, r2 = int(m.group(2)), int(m.group(4))
        if r1 != current_row or r2 != current_row:
            return True
    for m in _CELL_RE.finditer(formula):
        r = int(m.group(2))
        if r != current_row:
            return True
    return False


def _is_cross_sheet_ref(value) -> bool:
    return isinstance(value, str) and value.startswith("=") and "!" in value


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


class RowInfo:
    __slots__ = ("row_idx", "label", "label_norm", "role", "existing_formula")

    def __init__(self, row_idx, label, label_norm, role, existing_formula=None):
        self.row_idx = row_idx
        self.label = label
        self.label_norm = label_norm
        self.role = role
        self.existing_formula = existing_formula

    def __repr__(self):
        return f"RowInfo(R{self.row_idx} {self.role} {self.label[:30]!r})"


class YearSheet:
    __slots__ = ("sheet_name", "statement", "year", "header_row", "label_col",
                 "entity_cols", "summary_cols", "data_start", "data_end", "rows")

    def __init__(self, sheet_name, statement, year, header_row, label_col,
                 entity_cols, summary_cols, data_start, data_end, rows):
        self.sheet_name = sheet_name
        self.statement = statement
        self.year = year
        self.header_row = header_row
        self.label_col = label_col
        self.entity_cols = entity_cols
        self.summary_cols = summary_cols
        self.data_start = data_start
        self.data_end = data_end
        self.rows = rows


def discover_sheet(ws) -> Optional[YearSheet]:
    stmt, year = _classify_sheet(ws.title)
    if stmt == "OTHER":
        return None

    header_row = _find_header_row(ws)
    label_col = 1
    data_start = header_row + 1

    entity_cols = []
    summary_cols = []
    for c in range(label_col + 1, ws.max_column + 1):
        hdr = ws.cell(row=header_row, column=c).value
        if hdr is None or (isinstance(hdr, str) and not hdr.strip()):
            continue
        role = _classify_summary_col(str(hdr))
        if role:
            summary_cols.append((c, role))
        else:
            entity_cols.append((c, str(hdr).strip()))

    rows = []
    last_data_row = data_start
    for r in range(data_start, ws.max_row + 1):
        label = ws.cell(row=r, column=label_col).value
        if label is None or (isinstance(label, str) and not label.strip()):
            continue
        label = str(label).strip()
        ln = norm_label(label)

        any_cross_ref = False
        any_multi_row_sum = False
        any_same_row_sum = False
        any_static = False
        first_formula = None

        for c, _ in entity_cols:
            v = ws.cell(row=r, column=c).value
            if _is_cross_sheet_ref(v):
                any_cross_ref = True
                first_formula = first_formula or v
            elif _is_formula(v):
                if _formula_spans_multiple_rows(v, r):
                    any_multi_row_sum = True
                else:
                    any_same_row_sum = True
                first_formula = first_formula or v
            elif v is not None and v != "" and v != 0:
                any_static = True

        if any_cross_ref:
            role = "cross_ref"
        elif any_multi_row_sum:
            role = "subtotal"
        elif any_static and not any_same_row_sum:
            role = "preloaded"
        else:
            role = "data"

        rows.append(RowInfo(row_idx=r, label=label, label_norm=ln,
                            role=role, existing_formula=first_formula))
        last_data_row = r

    return YearSheet(
        sheet_name=ws.title, statement=stmt, year=year,
        header_row=header_row, label_col=label_col,
        entity_cols=entity_cols, summary_cols=summary_cols,
        data_start=data_start, data_end=last_data_row, rows=rows,
    )


def discover_template(wb_formula) -> List[YearSheet]:
    out = []
    for sn in wb_formula.sheetnames:
        ws = wb_formula[sn]
        ys = discover_sheet(ws)
        if ys:
            out.append(ys)
    return out


def summarize(sheets: List[YearSheet]) -> dict:
    return {
        "sheets": [
            {
                "name": s.sheet_name,
                "statement": s.statement,
                "year": s.year,
                "header_row": s.header_row,
                "n_entities": len(s.entity_cols),
                "entities": [name for _, name in s.entity_cols],
                "summary_cols": [{"col": c, "role": r} for c, r in s.summary_cols],
                "n_data": sum(1 for r in s.rows if r.role == "data"),
                "n_subtotal": sum(1 for r in s.rows if r.role == "subtotal"),
                "n_cross_ref": sum(1 for r in s.rows if r.role == "cross_ref"),
                "n_preloaded": sum(1 for r in s.rows if r.role == "preloaded"),
            }
            for s in sheets
        ]
    }
