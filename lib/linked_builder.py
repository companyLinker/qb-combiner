"""Template-agnostic SUMIFS-linked workbook builder.

The builder reads the structure of the target template via `template_discovery`
and writes SUMIFS only into data-role cells. It preserves:
  • subtotal rows (formulas that sum across other rows)
  • cross-sheet references (equity rollforward chain)
  • preloaded rows (cells already containing static numbers — by default skipped)

For multi-year templates (BS 2025 / BS 2024 / IS 2025 / IS 2024 / …), each sheet's
SUMIFS pulls from a `QB_Data` sheet filtered by year and statement.

How rows are matched between QB accounts and template line items:
  1. Profile mappings (saved overrides) win.
  2. Auto-rules (mapping_rules.map_pnl / map_bs) fill remaining.
  3. Each mapped account → target_line. The template's column-A label IS used as
     the SUMIFS lookup key. If a template label changes spelling, the user
     re-runs and the engine fuzzy-matches old target_lines to new template
     labels via rapidfuzz (or simple normalize-and-equal if rapidfuzz absent).
"""

import io
import re
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .mapping_rules import map_pnl, map_bs
from .master_builder import write_pivot_sheets
from .template_discovery import (
    discover_template, summarize, norm_label, YearSheet, RowInfo,
)

try:
    # pyrefly: ignore [missing-import]
    from rapidfuzz import fuzz, process as rf_process
    HAVE_RAPIDFUZZ = True
except ImportError:
    HAVE_RAPIDFUZZ = False


HDR_FILL = PatternFill("solid", fgColor="1F3864")
WHITE = Font(color="FFFFFF", bold=True)


def style_header(ws, row=1):
    for c in ws[row]:
        c.font = WHITE
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def fmt_money(c):
    c.number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'


def fit(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _best_template_row_for(target_line: str, candidates: List[str], threshold: int = 75) -> Optional[str]:
    """Given a target_line (from the mapping rules), find the closest template
    label among the candidates. Uses rapidfuzz when available; falls back to
    exact-normalized equality."""
    if not target_line or not candidates:
        return None
    tnorm = norm_label(target_line)
    if not tnorm:
        return None
    cand_norm = {norm_label(c): c for c in candidates}
    if tnorm in cand_norm:
        return cand_norm[tnorm]
    if HAVE_RAPIDFUZZ:
        best = rf_process.extractOne(tnorm, list(cand_norm.keys()),
                                      scorer=fuzz.token_set_ratio,
                                      score_cutoff=threshold)
        if best:
            return cand_norm[best[0]]
    return None


def compute_mapping(qb_data, mapping_overrides=None, entity_mapping_overrides=None):
    """Run auto-rules on every leaf QB account, allow overrides. Returns:
        mapping: dict[(statement, entity, breadcrumb, qb_account)] = (target_line, source)
        pnl_leaves, bs_leaves: dict[(breadcrumb, label)] = {amounts: {entity: $}}

    Lookup priority per entity+account:
      1. Entity-specific override  (key: 'E|stmt|entity|bc|lbl')
      2. Generic override          (key: 'stmt|bc|lbl')
      3. Auto-rules
    """
    overrides        = mapping_overrides or {}
    entity_overrides = entity_mapping_overrides or {}
    pnl_leaves: Dict = {}
    bs_leaves:  Dict = {}
    mapping:    Dict = {}

    for entity, info in qb_data.items():
        for r in info.get("pnl_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            k = (r["breadcrumb"], r["label"])
            d = pnl_leaves.setdefault(k, {"amounts_cy": {}, "amounts_py": {},
                                          "breadcrumb": r["breadcrumb"], "label": r["label"]})
            d["amounts_cy"][entity] = d["amounts_cy"].get(entity, 0) + (r.get("amount_cy") or r.get("amount") or 0)
            if r.get("amount_py") is not None:
                d["amounts_py"][entity] = d["amounts_py"].get(entity, 0) + (r.get("amount_py") or 0)
        for r in info.get("bs_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            k = (r["breadcrumb"], r["label"])
            d = bs_leaves.setdefault(k, {"amounts_cy": {}, "amounts_py": {},
                                         "breadcrumb": r["breadcrumb"], "label": r["label"]})
            d["amounts_cy"][entity] = d["amounts_cy"].get(entity, 0) + (r.get("amount_cy") or r.get("amount") or 0)
            if r.get("amount_py") is not None:
                d["amounts_py"][entity] = d["amounts_py"].get(entity, 0) + (r.get("amount_py") or 0)

    # Build per-entity mapping — entity-specific override wins
    for entity, info in qb_data.items():
        for r in info.get("pnl_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            bc, lbl = r["breadcrumb"], r["label"]
            entity_key  = f"E|P&L|{entity}|{bc}|{lbl}"
            generic_key = f"P&L|{bc}|{lbl}"
            mkey = ("P&L", entity, bc, lbl)
            if entity_key in entity_overrides and entity_overrides[entity_key]:
                mapping[mkey] = (entity_overrides[entity_key], "entity")
            elif generic_key in overrides and overrides[generic_key]:
                mapping[mkey] = (overrides[generic_key], "manual")
            else:
                t, c = map_pnl(bc, lbl)
                mapping[mkey] = (t if t and t != "__SKIP__" else "", c)

        for r in info.get("bs_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            bc, lbl = r["breadcrumb"], r["label"]
            entity_key  = f"E|BS|{entity}|{bc}|{lbl}"
            generic_key = f"BS|{bc}|{lbl}"
            mkey = ("BS", entity, bc, lbl)
            if entity_key in entity_overrides and entity_overrides[entity_key]:
                mapping[mkey] = (entity_overrides[entity_key], "entity")
            elif generic_key in overrides and overrides[generic_key]:
                mapping[mkey] = (overrides[generic_key], "manual")
            else:
                mapping[mkey] = map_bs(bc, lbl)

    return mapping, pnl_leaves, bs_leaves


def _write_qb_data_sheet(wb, qb_data, mapping):
    """Hidden source sheet: every leaf row × CY/PY × entity, with resolved target_line."""
    sheet_name = "QB_Data"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 0)
    ws.append(["Year", "Entity", "Statement", "Breadcrumb", "QB Account",
               "Target Line", "Mapping Key", "Amount"])
    style_header(ws)
    ws.freeze_panes = "A2"

    for entity, info in qb_data.items():
        # Write a "CY" row and (if available) a "PY" row for each leaf
        for r in info.get("pnl_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            # Look up entity-specific target first, then generic
            tgt, _ = mapping.get(("P&L", entity, r["breadcrumb"], r["label"]),
                       mapping.get(("P&L", r["breadcrumb"], r["label"]), ("", "REVIEW")))
            key = f"P&L|{r['breadcrumb']}|{r['label']}"
            cy = r.get("amount_cy") or r.get("amount") or 0
            py = r.get("amount_py")
            ws.append(["CY", entity, "P&L", r["breadcrumb"], r["label"], tgt, key, cy])
            fmt_money(ws.cell(row=ws.max_row, column=8))
            if py is not None:
                ws.append(["PY", entity, "P&L", r["breadcrumb"], r["label"], tgt, key, py])
                fmt_money(ws.cell(row=ws.max_row, column=8))
        for r in info.get("bs_rows", []):
            if r["is_section_only"] or r["is_total"]:
                continue
            tgt, _ = mapping.get(("BS", entity, r["breadcrumb"], r["label"]),
                       mapping.get(("BS", r["breadcrumb"], r["label"]), ("", "REVIEW")))
            key = f"BS|{r['breadcrumb']}|{r['label']}"
            cy = r.get("amount_cy") or r.get("amount") or 0
            py = r.get("amount_py")
            ws.append(["CY", entity, "BS", r["breadcrumb"], r["label"], tgt, key, cy])
            fmt_money(ws.cell(row=ws.max_row, column=8))
            if py is not None:
                ws.append(["PY", entity, "BS", r["breadcrumb"], r["label"], tgt, key, py])
                fmt_money(ws.cell(row=ws.max_row, column=8))
    fit(ws, [6, 28, 6, 50, 35, 35, 60, 16])
    # Hide it from casual view
    ws.sheet_state = "hidden"


def _write_template_map_sheet(wb, year_sheets, mapping):
    """Audit sheet that records, per template year-sheet, every data-row label
    and which QB target_lines route to it. Useful for diagnosing why a cell is 0."""
    if "Template_Map" in wb.sheetnames:
        del wb["Template_Map"]
    ws = wb.create_sheet("Template_Map", 1)
    ws.append(["Sheet", "Row", "Template Label", "Matched Target Line", "Source QB Accounts"])
    style_header(ws)
    ws.freeze_panes = "A2"

    # Build reverse index: target_line → list of QB account keys
    by_target = {}
    for (stmt, entity, bc, lbl), (tgt, src) in mapping.items():
        if tgt:
            by_target.setdefault(tgt, []).append((stmt, bc, lbl))

    for ys in year_sheets:
        for row in ys.rows:
            if row.role != "data":
                continue
            # Find which target_line(s) match this template label
            matched = _best_template_row_for(row.label, list(by_target.keys()))
            qbs = by_target.get(matched, []) if matched else []
            qb_summary = "; ".join(f"{stmt}: {lbl}" for stmt, bc, lbl in qbs[:5])
            if len(qbs) > 5:
                qb_summary += f"  … (+{len(qbs)-5} more)"
            ws.append([ys.sheet_name, row.row_idx, row.label, matched or "(no match)", qb_summary])
    fit(ws, [22, 6, 40, 40, 80])


def _resolve_row_to_target(template_label: str, target_lines: List[str]) -> Optional[str]:
    """Fuzzy-match a template label against known target_lines."""
    return _best_template_row_for(template_label, target_lines)


def build_linked_workbook(
    qb_data,
    target_bytes,
    mapping_overrides=None,
    entity_mapping_overrides=None,
    overwrite_preloaded: bool = False,
    selected_sheets=None,
    entity_col_mapping=None,
):
    """Build the linked combination workbook. Returns (BytesIO, mapping, year_sheets, report).

    Args:
        selected_sheets: Optional list of template sheet names to process.
            If None, all IS/BS-classified sheets are processed.
        entity_col_mapping: Optional dict mapping template column header name
            to the QB entity name whose data should flow into that column.
            A value of None/missing means use the column header verbatim
            (auto-matching by name). Set a column's value to None to skip it.
    """
    wb = openpyxl.load_workbook(io.BytesIO(target_bytes))
    year_sheets = discover_template(wb)

    # Filter to selected sheets only
    if selected_sheets:
        year_sheets = [ys for ys in year_sheets if ys.sheet_name in selected_sheets]

    if not year_sheets:
        raise RuntimeError(
            f"No IS/BS sheets found in the selection. "
            f"Sheets present: {wb.sheetnames}. "
            f"Selected: {selected_sheets}"
        )

    mapping, pnl_leaves, bs_leaves = compute_mapping(
        qb_data, mapping_overrides, entity_mapping_overrides
    )
    
    # Delete old pivot/QB_Data sheets if they exist in the uploaded template
    for suffix in ["CY", "PY", "Change"]:
        for base in ["P&L Pivot", "BS Pivot"]:
            sn = f"{base} {suffix}"
            if sn in wb.sheetnames:
                del wb[sn]
    if "QB_Data" in wb.sheetnames:
        del wb["QB_Data"]

    # Write fresh pivot sheets containing the Target Line column (Column C)
    write_pivot_sheets(wb, qb_data, mapping_overrides, entity_mapping_overrides)
    _write_template_map_sheet(wb, year_sheets, mapping)

    # Available target_lines (from mapping output)
    pnl_targets = sorted({t for (s, _, _, _), (t, _) in mapping.items() if s == "P&L" and t})
    bs_targets = sorted({t for (s, _, _, _), (t, _) in mapping.items() if s == "BS" and t})

    report = {
        "n_year_sheets": len(year_sheets),
        "cells_written": 0,
        "cells_skipped_preloaded": 0,
        "cells_skipped_subtotal": 0,
        "cells_skipped_crossref": 0,
        "rows_unmapped": [],
    }

    entities = list(qb_data.keys())

    for ys in year_sheets:
        ws = wb[ys.sheet_name]
        is_bs = ys.statement == "BS"
        year_filter = "CY" if (ys.year is None or _is_cy_year(ys, year_sheets)) else "PY"
        targets = bs_targets if is_bs else pnl_targets

        for row in ys.rows:
            if row.role == "subtotal":
                report["cells_skipped_subtotal"] += len(ys.entity_cols)
                continue
            if row.role == "cross_ref":
                report["cells_skipped_crossref"] += len(ys.entity_cols)
                continue
            if row.role == "preloaded" and not overwrite_preloaded:
                report["cells_skipped_preloaded"] += len(ys.entity_cols)
                continue
            if row.role not in ("data", "preloaded"):
                continue

            # Match this template label to a target_line
            matched_target = _resolve_row_to_target(row.label, targets)
            if not matched_target:
                report["rows_unmapped"].append({
                    "sheet": ys.sheet_name, "row": row.row_idx, "label": row.label,
                })
                continue

            # Write SUMIFS into each entity column
            for col, entity_header in ys.entity_cols:
                cell = ws.cell(row=row.row_idx, column=col)
                # If preloaded and not overwriting, skip per-cell (re-check)
                if row.role == "preloaded" and not overwrite_preloaded:
                    continue

                # Resolve the matching QB entity and year filter using the smart resolver
                qb_entity, resolved_year_filter = resolve_qb_source(
                    ys, year_sheets, qb_data, entity_col_mapping, entity_header
                )
                
                if qb_entity is None:
                    # Skipped or not found: leave cell untouched
                    continue

                if qb_entity in entities:
                    idx = entities.index(qb_entity)
                    col_letter = get_column_letter(6 + idx)  # Column F onwards (6 = F)
                else:
                    # Entity not uploaded/found: leave cell untouched
                    continue

                sheet_base = "BS Pivot" if ys.statement == "BS" else "P&L Pivot"
                pivot_sheet_name = f"{sheet_base} {resolved_year_filter}"
                formula = f"=SUMIFS('{pivot_sheet_name}'!${col_letter}:${col_letter}, '{pivot_sheet_name}'!$C:$C, \"{matched_target}\")"
                
                cell.value = formula
                fmt_money(cell)
                report["cells_written"] += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, mapping, year_sheets, report


def _is_cy_year(ys: YearSheet, all_sheets: List[YearSheet]) -> bool:
    """The newest year among sheets of the same statement is the CY."""
    years = [s.year for s in all_sheets if s.statement == ys.statement and s.year]
    if not years:
        return True
    return ys.year == max(years)


def resolve_qb_source(
    ys: YearSheet,
    all_sheets: List[YearSheet],
    qb_data: dict,
    entity_col_mapping: Optional[dict],
    entity_header: str,
) -> Tuple[Optional[str], str]:
    """Resolves the correct QB entity key and the year filter ('CY' or 'PY') for a given template column and sheet.

    Returns:
        (resolved_qb_entity_key, year_filter)
    """
    is_cy = _is_cy_year(ys, all_sheets)

    # 1. Check manual mapping first
    mapped_entity = None
    if entity_col_mapping is not None:
        if entity_header in entity_col_mapping:
            mapped_entity = entity_col_mapping[entity_header]
            if mapped_entity is None:
                # Explicitly mapped to "Skip"
                return None, "CY"

    # 2. Extract metadata and clean names for all uploaded QB entities
    candidates = []
    for ent_key, info in qb_data.items():
        # Detect year from period strings or filename
        ent_year = None
        for p in [info.get("pnl_period_cy"), info.get("bs_period_cy"), info.get("pnl_period"), info.get("bs_period")]:
            if p:
                p_str = str(p).strip()
                m = re.search(r"\b(20\d{2})\b", p_str)
                if m:
                    ent_year = int(m.group(1))
                    break
                m2 = re.search(r"\b(\d{2})$", p_str)
                if m2:
                    val = int(m2.group(1))
                    if 0 <= val <= 99:
                        ent_year = 2000 + val if val < 50 else 1900 + val
                        break
                # Look for a 2 digit year with a preceding boundary, e.g. - 24, / 24, space 24
                m3 = re.search(r"\b(\d{2})\b", p_str)
                if m3:
                    val = int(m3.group(1))
                    if 15 <= val <= 35:
                        ent_year = 2000 + val
                        break
        
        if not ent_year:
            m_fn = re.search(r"\b(20\d{2})\b", info.get("file", ""))
            if m_fn:
                ent_year = int(m_fn.group(1))
            else:
                m_fn2 = re.search(r"\b(\d{2})\b", info.get("file", ""))
                if m_fn2:
                    val = int(m_fn2.group(1))
                    if 20 <= val <= 35:
                        ent_year = 2000 + val

        # Clean the name: normalize and strip any year numbers out
        clean_ent = norm_label(ent_key)
        clean_ent_no_year = clean_ent
        if ent_year:
            yr_str = str(ent_year)
            yr_short = str(ent_year % 100)
            clean_ent_no_year = re.sub(r"\b(" + yr_str + r"|" + yr_short + r")\b", "", clean_ent)
            clean_ent_no_year = re.sub(r"[\s\-_]+", " ", clean_ent_no_year).strip()

        candidates.append({
            "key": ent_key,
            "clean_name": clean_ent_no_year,
            "year": ent_year,
            "variant": info.get("variant", "single")
        })

    # If manual mapping is specified, find that entity key
    if mapped_entity:
        cand = next((c for c in candidates if c["key"] == mapped_entity), None)
        if cand:
            if cand["variant"] == "triple":
                return cand["key"], ("CY" if is_cy else "PY")
            else:
                return cand["key"], "CY"

    # Otherwise, auto-resolve by comparing normalized name with normalized header
    clean_header = norm_label(entity_header)
    
    # Try exact match on clean name first
    matches = [c for c in candidates if c["clean_name"] == clean_header or norm_label(c["key"]) == clean_header]
    # Fallback to substring matching if needed
    if not matches:
        matches = [c for c in candidates if clean_header in c["clean_name"] or clean_header in norm_label(c["key"])]

    if not matches:
        return None, "CY"

    # Case A: If one of the matching files is a multi-year (triple) file, use it
    triple_match = next((c for c in matches if c["variant"] == "triple"), None)
    if triple_match:
        return triple_match["key"], ("CY" if is_cy else "PY")

    # Case B: Match the single-year file that corresponds to the template sheet year
    if ys.year is not None:
        year_match = next((c for c in matches if c["year"] == ys.year), None)
        if year_match:
            return year_match["key"], "CY"

    # Default fallback to first match
    return matches[0]["key"], "CY"
