"""Auto-mapping rules for QuickBooks accounts → target combination workbook line items.

Two main functions:
    map_pnl(breadcrumb, label) → (target_line, confidence)
    map_bs(breadcrumb, label)  → (target_line, confidence)

confidence is 'auto' if a rule fired, 'REVIEW' if no rule matched.
target_line == '__SKIP__' means this is a structural pseudo-leaf (e.g., 'Gross Profit'
appears as a leaf in some QB exports because of a display setting) — should not flow.

The target line strings match the exact text in the target template column (label col).
Leading/trailing whitespace matches template indentation — do NOT strip these.

Coverage observed on the May 2026 dataset (22 entities, FY 2025):
  - P&L: 157/158 auto-mapped (99%)
  - BS:  465/499 auto-mapped (93%)
  
AP Illinois MGMT LLC 2025 template additions — June 2026:
  - IS2025: ~30 new line items added including KIOSK FEES, WOTC, adjustments
  - BS2025: ~40 new/renamed line items including ACCUM. AMORT. sub-items
"""

import re
from functools import lru_cache


@lru_cache(maxsize=8192)
def norm(s):
    """Normalize a label for fuzzy matching: lowercase, alphanumeric+space only."""
    if not s:
        return ""
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s)


@lru_cache(maxsize=8192)
def map_pnl(breadcrumb, label):
    """Return (target_line, confidence) for a P&L leaf account."""
    bc_n = norm(breadcrumb)
    lbl_n = norm(label)
    full_n = f"{bc_n} {lbl_n}".strip()

    # ── Structural pseudo-rows — skip ────────────────────────────────────────
    if lbl_n == "gross profit":
        return "__SKIP__", "auto"
    if "operating expenses other" in lbl_n:
        return "    MISC EXP", "auto"
    if "labor cost other" in lbl_n:
        return "   KITCHEN LABOR", "auto"

    # ============ INCOME ============
    if (
        "food sales" in bc_n
        or bc_n.endswith(" income")
        or bc_n == "ordinary income expense > income"
    ):
        if any(
            x in lbl_n
            for x in [
                "non taxable sales", "taxable sales",
                "delivery sales refund", "merchant card",
                "online sales", "delivery sale refund",
            ]
        ):
            return "RETAIL SALES", "auto"

    if any(x in lbl_n for x in ["retail sales", "food sales", "food sale", "taxable sales"]):
        return "RETAIL SALES", "auto"

    if "rebate" in lbl_n:
        return "REBATE INCOME", "auto"

    if "management fees" in lbl_n or "management fee" in lbl_n:
        bc_parts = [p.strip().lower() for p in breadcrumb.split(" > ")]
        is_income = False
        for part in bc_parts:
            if "other income" in part:
                is_income = True
            elif "income" in part and "expense" not in part:
                is_income = True
        if any("expense" in part for part in bc_parts):
            is_income = False
        if is_income:
            return "MANAGEMENT FEES INCOME", "auto"
        else:
            return "    MANAGEMENT FEES", "auto"

    if "rental income" in lbl_n or ("rent" in lbl_n and "income" in lbl_n):
        return "RENTAL INCOME", "auto"

    if "cam charges collected" in lbl_n or "cam charge" in lbl_n and "income" in bc_n:
        return "    CAM CHARGES COLLECTED", "auto"

    if "income from partnership" in lbl_n or "k 1 income" in lbl_n or "k1 income" in lbl_n:
        return "    INCOME FROM PARTNERSHIP / K-1", "auto"

    if "income from other" in lbl_n and "entit" in lbl_n:
        return "    INCOME FROM OTHER ENTITIES", "auto"

    if "erc" in lbl_n:
        if "interest" in lbl_n:
            return "ERC Interest Income", "auto"
        return "ERC Credit Income", "auto"

    if "ppp" in lbl_n and "income" in lbl_n:
        return "Non Taxable PPP Income", "auto"

    if "eidl grant" in lbl_n:
        return "    EIDL GRANT INCOME", "auto"

    if "discount" in lbl_n and ("receive" in lbl_n or "recievable" in lbl_n):
        return "REBATE INCOME", "auto"

    if "misc" in lbl_n and "income" in lbl_n:
        return "REBATE INCOME", "auto"

    if "interest" in lbl_n and "income" in lbl_n and "expense" not in lbl_n:
        return "INTEREST INCOME", "auto"

    # ============ COGS — Franchise Fees ============
    if "franchise fees" in bc_n:
        if any(x in lbl_n for x in ["ad fund", "advt", "royalty", "advertising"]):
            return " FRANCHISE & ADVT FEES", "auto"
        if "preservation" in lbl_n:
            return "    FRANCHISE PRESERVATION FEES", "auto"

    # ============ COGS — Labor ============
    if "labor cost" in bc_n:
        if "regular hour" in lbl_n:
            return " SALARIES & WAGES", "auto"
        if "ot hour" in lbl_n or "overtime" in lbl_n:
            return " SALARIES & WAGES", "auto"
        if lbl_n == "salary" or lbl_n.endswith(" salary") or lbl_n.startswith("salary"):
            return " SALARIES & WAGES", "auto"
        if "covid" in lbl_n:
            return "  COVID CARE", "auto"
        if "kitchen" in lbl_n:
            return "   KITCHEN LABOR", "auto"
        if "training" in lbl_n:
            return "   TRAINING COST", "auto"
        if "bonus" in lbl_n and "performance" not in lbl_n:
            return "    BONUS", "auto"
        if "performance" in lbl_n:
            return "  PERFORMANCE BONUS", "auto"
        if "worker" in lbl_n or "workmen" in lbl_n:
            return " INSURANCE & WORKERS' COMPENSATION", "auto"
        if (
            "payroll tax" in bc_n
            or "social security" in lbl_n
            or "unemployment" in lbl_n
            or "fica" in lbl_n
        ):
            return "    PAYROLL TAXES", "auto"
        if "fed" in lbl_n and "unemployment" in lbl_n:
            return "    PAYROLL TAXES", "auto"
        if "profit sharing" in lbl_n:
            return "    PROFIT SHARING EXPENSE", "auto"
        if "salary" in lbl_n or "wage" in lbl_n:
            return " SALARIES & WAGES", "auto"
        # Catch-all labor
        return " SALARIES & WAGES", "auto"

    # Training standalone
    if "training" in lbl_n and "labor" not in bc_n and "cost" in lbl_n:
        return "   TRAINING COST", "auto"

    # ============ COGS — Purchases / Inventory ============
    if "purchases" in bc_n or "purchase" in bc_n:
        if "opening" in lbl_n:
            return "Opening Inventory", "auto"
        if "closing" in lbl_n or "ending" in lbl_n:
            return "Ending Inventory", "auto"
        return "  PURCHASE", "auto"

    # ============ DELIVERY / AGGREGATORS ============
    if "delivery charges" in bc_n or any(
        x in lbl_n
        for x in [
            "doordash", "grubhub", "ubereats", "uber eats", "otter",
            "ezcarter", "ez carter",
        ]
    ):
        if "income" not in bc_n and "income" not in lbl_n:
            return "DELIVERY FEES EXPENSE", "auto"

    if "gift card charges" in lbl_n:
        return "    GIFT CARD CHARGES", "auto"

    # ============ EXPENSES — broad-category short-circuit ============
    if "auto expenses" in bc_n or ("auto" in lbl_n and "travel" in lbl_n):
        return "    AUTO EXPENSES AND TRAVEL", "auto"
    if "cleaning expenses" in bc_n or "cleaning" in lbl_n:
        return "    CLEANING EXPENSES", "auto"
    if "insurance" in bc_n.lower() or (
        "insurance" in lbl_n and "worker" in lbl_n
    ):
        return "    INSURANCE & WORKERS' COMPENSATION", "auto"
    if "insurance" in lbl_n and "worker" not in lbl_n and "income" not in lbl_n:
        return "    INSURANCE & WORKERS' COMPENSATION", "auto"

    if "occupancy" in bc_n:
        bc_parts = [p.strip().lower() for p in breadcrumb.split(" > ")]
        last_part = bc_parts[-1] if bc_parts else ""
        if "rent" in last_part or "rent" in lbl_n:
            return "    RENT & CAM CHARGES", "auto"
        if "repair" in bc_n or "repair" in lbl_n or "maint" in lbl_n:
            return "    REPAIRS AND MAINTENANCE", "auto"
        if "utilities" in bc_n or any(
            x == lbl_n or lbl_n.startswith(x)
            for x in ["electricity", "water", "gas", "heating", "cooling", "sewer"]
        ):
            return "    UTILITIES", "auto"

    if "uniform" in lbl_n:
        return "    UNIFORMS", "auto"

    if any(
        x in lbl_n
        for x in [
            "restaurant supplies", "restaurant supply",
            "restaurant supp", "kitchen smallware", "smallware",
        ]
    ):
        return "    RESTAURANT SUPPLIES", "auto"

    # ============ SPECIFIC LINE ITEMS (high-priority exact matches) ============
    if "digital transaction" in lbl_n or "ordering tech fee" in lbl_n or "ordering technology" in lbl_n:
        return "    POPEYES ORDERING TECHNOLOGY FEE", "auto"
    if "guest care fee" in lbl_n or "guest care" in lbl_n:
        return "    POPEYES GUEST CARE FEE", "auto"
    if "service check" in lbl_n and ("popeye" in lbl_n or "popeye" in bc_n):
        return "    POPEYES SERVICE CHECK", "auto"
    if "food" in lbl_n and "employee" in lbl_n:
        return "    FOOD FOR EMPLOYEES", "auto"
    if "loan fee" in lbl_n or "loan fees" in lbl_n:
        return "    LOAN FEES", "auto"
    if "kiosk" in lbl_n:
        return "    KIOSK FEES", "auto"
    if "dues" in lbl_n or "subscription" in lbl_n:
        return "    DUES AND SUBSCRIPTIONS", "auto"
    if "cash handling" in lbl_n:
        return "    CASH HANDLING SERVICE", "auto"
    if "house charge" in lbl_n:
        return "    HOUSE CHARGE", "auto"
    if "refinance" in lbl_n:
        return "    REFINANCE CHARGES", "auto"
    if "guaranteed payment" in lbl_n:
        return "    GUARANTEED PAYMENT", "auto"
    if "brokerage" in lbl_n:
        return "    BROKERAGE FEES", "auto"
    if "development rights written off" in lbl_n:
        return "    DEVELOPMENT RIGHTS WRITTEN OFF", "auto"
    if "franchise preservation" in lbl_n:
        return "    FRANCHISE PRESERVATION FEES", "auto"
    if "gain" in lbl_n and ("loss" in lbl_n or "sale" in lbl_n) and "asset" in lbl_n:
        return "    GAIN / (LOSS) ON SALE OF ASSETS", "auto"
    if "non deductible" in lbl_n:
        return "    NON-DEDUCTIBLE EXPENSE", "auto"
    if "wotc" in lbl_n:
        if "25" in lbl_n or "wage" in lbl_n:
            return "    WAGES ELIGIBLE @ 25%", "auto"
        if "40" in lbl_n:
            return "    WAGES ELIGIBLE @ 40%", "auto"
        if "50" in lbl_n:
            return "    WAGES ELIGIBLE @ 50%", "auto"
        return "    WOTC NON-TAXABLE EXPENSE", "auto"
    if "empowerment zone" in lbl_n:
        return "    EMPOWERMENT ZONE CREDIT @ 20%", "auto"
    if "donation from k" in lbl_n or "donation from k1" in lbl_n or "passthrough" in lbl_n:
        return "    DONATION FROM K-1 / PASSTHROUGH", "auto"
    if "1231 loss" in lbl_n:
        return "    1231 LOSS FROM K-1", "auto"
    if "profit" in lbl_n and "transfer" in lbl_n and "management" in lbl_n:
        return "    PROFIT / LOSS TRANSFER TO MANAGEMENT ENTITY", "auto"

    # ============ OPERATING EXPENSES — catch-all bucket ============
    if "operating expenses" in bc_n.lower():
        if "credit card" in lbl_n or "merchant fee" in lbl_n:
            return "    CREDIT CARD CHARGES", "auto"
        if "license" in lbl_n or "permit" in lbl_n:
            return "    LICENSES AND PERMITS", "auto"
        if "accounting" in lbl_n:
            return "    PROFESSIONAL FEES", "auto"
        if "bank service" in lbl_n or "bank charge" in lbl_n:
            return "    BANK SERVICE CHARGES", "auto"
        if "401" in lbl_n or "401k" in lbl_n:
            return "    401(K) EXPENSES", "auto"
        if "payroll" in lbl_n and "expense" in lbl_n:
            return "    PAYROLL EXPENSE", "auto"
        if "payroll processing" in lbl_n:
            return "    PAYROLL PROCESSING", "auto"
        if "shortage" in lbl_n or "cash short" in lbl_n or "over" in lbl_n:
            return "    SHORTAGE AND OVERS", "auto"
        if "office" in lbl_n:
            return "    OFFICE SUPPLIES AND EXPENSE", "auto"
        if "security" in lbl_n or "alarm" in lbl_n:
            return "    ALARM AND SECURITY", "auto"
        if "donation" in lbl_n or "charity" in lbl_n:
            return "    DONATION", "auto"
        if "equipment rental" in lbl_n:
            return "    EQUIPMENT RENTAL", "auto"
        if "popeyes" in lbl_n or "popeye" in lbl_n:
            return "    POPEYES SERVICE CHECK", "auto"
        if "professional" in lbl_n or "legal" in lbl_n:
            return "    PROFESSIONAL FEES", "auto"
        if "convention" in lbl_n or "convension" in lbl_n:
            return "    CONVENTION EXPENSE", "auto"
        if "computer" in lbl_n or "software" in lbl_n:
            return "    SOFTWARE EXPENSE", "auto"
        if "advertis" in lbl_n or "promotion" in lbl_n:
            return "    ADVERTISING AND PROMOTION", "auto"
        if "non recurring" in lbl_n:
            return "  Non recurring expenses", "auto"
        if "penalty" in lbl_n or "interest" in lbl_n:
            return "    PENALTY AND INTEREST", "auto"
        if "annual report" in lbl_n:
            return "    ANNUAL REPORT FEES", "auto"
        if "filing fee" in lbl_n or "state filing" in lbl_n:
            return "    STATE FILING FEES", "auto"
        if "misc" in lbl_n and "tax" not in lbl_n:
            return "    MISCELLANEOUS EXPENSES", "auto"
        if "suspense" in lbl_n:
            return "    MISCELLANEOUS EXPENSES", "auto"
        # Catch-all
        return "    MISCELLANEOUS EXPENSES", "auto"

    # ============ TAXES (standalone) ============
    if "taxes" in bc_n.lower() or "tax" in lbl_n.lower():
        if "real estate" in lbl_n or "property tax" in lbl_n:
            if "maryland" in lbl_n:
                return "    -PROPERTY TAX-MARYLAND", "auto"
            return "    REAL ESTATE TAX", "auto"
        if "annual report" in lbl_n:
            return "    ANNUAL REPORT FEES", "auto"
        if "filing fee" in lbl_n or "filing fees" in lbl_n:
            return "    STATE FILING FEES", "auto"
        if "nj" in lbl_n or "new jersey" in lbl_n:
            return "    STATE FILING FEES", "auto"
        if "restaurant" in lbl_n or "meal" in lbl_n:
            return "    MEAL TAXES", "auto"
        if "misc" in lbl_n or "other" in lbl_n:
            return "    MISCELLANEOUS EXPENSES", "auto"
        if "state" in lbl_n or "illinois replacement" in lbl_n:
            return "    STATE TAXES / ILLINOIS REPLACEMENT TAX", "auto"
        if "mercantile" in lbl_n or "local" in lbl_n or "bpt" in lbl_n or "school" in lbl_n:
            return "   LOCAL MERCANTILE TAX", "auto"
        if "gross receipt" in lbl_n:
            return "    Gross receipt tax", "auto"

    # ============ INTEREST EXPENSES ============
    if (
        ("interest" in bc_n.lower() and "expense" in bc_n.lower())
        or "interest expenses" in lbl_n
        or "interest to bank" in lbl_n
        or "interest to other" in lbl_n
    ):
        return "        INTEREST TO BANK" if "bank" in lbl_n else "        INTEREST TO OTHERS", "auto"

    if "interest expense" in lbl_n:
        return "        INTEREST TO OTHERS", "auto"

    # ============ NON-CASH / OTHER ============
    if "depreciation" in lbl_n and "accumulated" not in lbl_n:
        return "    DEPRECIATION", "auto"
    if "amortization" in lbl_n or "amortisation" in lbl_n:
        return "    AMORTIZATION", "auto"

    if "profit sharing" in bc_n.lower() or "profit sharing" in lbl_n.lower():
        return "    PROFIT SHARING EXPENSE", "auto"
    if "profit transfer" in lbl_n or ("transfer" in lbl_n and "north" in lbl_n):
        return "Profit Transfer to AP North", "auto"

    # ============ FALLBACK — try to match by keyword ============
    if "rent" in lbl_n:
        return "    RENT & CAM CHARGES", "auto"
    if "utility" in lbl_n or "utilities" in lbl_n:
        return "    UTILITIES", "auto"
    if "repair" in lbl_n or "maintenance" in lbl_n:
        return "    REPAIRS AND MAINTENANCE", "auto"
    if "credit card" in lbl_n:
        return "    CREDIT CARD CHARGES", "auto"
    if "bank" in lbl_n and "charge" in lbl_n:
        return "    BANK SERVICE CHARGES", "auto"
    if "payroll" in lbl_n and "processing" in lbl_n:
        return "    PAYROLL PROCESSING", "auto"
    if "advertis" in lbl_n:
        return "    ADVERTISING AND PROMOTION", "auto"
    if "alarm" in lbl_n or "security" in lbl_n:
        return "    ALARM AND SECURITY", "auto"
    if "license" in lbl_n or "permit" in lbl_n:
        return "    LICENSES AND PERMITS", "auto"
    if "software" in lbl_n:
        return "    SOFTWARE EXPENSE", "auto"
    if "convention" in lbl_n:
        return "    CONVENTION EXPENSE", "auto"
    if "professional" in lbl_n:
        return "    PROFESSIONAL FEES", "auto"
    if "legal" in lbl_n:
        return "    LEGAL FEES", "auto"
    if "donation" in lbl_n:
        return "    DONATION", "auto"
    if "office" in lbl_n:
        return "    OFFICE SUPPLIES AND EXPENSE", "auto"
    if "shortage" in lbl_n or "overage" in lbl_n:
        return "    SHORTAGE AND OVERS", "auto"
    if "penalty" in lbl_n:
        return "    PENALTY AND INTEREST", "auto"
    if "annual report" in lbl_n:
        return "    ANNUAL REPORT FEES", "auto"
    if "state filing" in lbl_n or "filing fee" in lbl_n:
        return "    STATE FILING FEES", "auto"
    if "auto" in lbl_n and "travel" in lbl_n:
        return "    AUTO EXPENSES AND TRAVEL", "auto"
    if "cleaning" in lbl_n:
        return "    CLEANING EXPENSES", "auto"
    if "delivery" in lbl_n and "fee" in lbl_n:
        return "DELIVERY FEES EXPENSE", "auto"

    return None, "REVIEW"


def resolve_target_entries(stmt_kind, entity, breadcrumb, label,
                            overrides=None, entity_overrides=None):
    """Resolve the list of (target_line, source, pivot_override) for one
    entity's QB leaf account.

    Normally returns exactly one entry (entity override > generic override >
    auto-rule, same priority as before). If the entity has "duplicate" rows
    (created via the Variants & Analysis "Duplicate row" control), one extra
    entry is appended per duplicate with a non-blank target line, so the same
    account's amount fans out into multiple Target Lines.

    entity_overrides[entity_key] is a list of
    {"dup_id": str, "target_line": str, "pivot_override": str} dicts, where
    the primary/original mapping has dup_id == "".
    """
    overrides = overrides or {}
    entity_overrides = entity_overrides or {}
    entity_key  = f"E|{stmt_kind}|{entity}|{breadcrumb}|{label}"
    generic_key = f"{stmt_kind}|{breadcrumb}|{label}"

    entries = entity_overrides.get(entity_key) or []
    primary = next((e for e in entries if not e.get("dup_id")), None)
    dups    = [e for e in entries if e.get("dup_id")]

    results = []
    if primary and (primary.get("target_line") or "").strip():
        results.append((
            (primary["target_line"] or "").strip(),
            "entity",
            (primary.get("pivot_override") or "").strip(),
        ))
    elif generic_key in overrides and overrides[generic_key]:
        results.append((overrides[generic_key].strip(), "manual", ""))
    else:
        fn = map_pnl if stmt_kind == "P&L" else map_bs
        t, c = fn(breadcrumb, label)
        results.append(((t or "").strip() if t != "__SKIP__" else "", c, ""))

    for d in dups:
        tl = (d.get("target_line") or "").strip()
        if tl:
            results.append((tl, "entity-dup", (d.get("pivot_override") or "").strip()))

    return results


@lru_cache(maxsize=8192)
def map_bs(breadcrumb, label):
    """Return (target_line, confidence) for a Balance Sheet leaf account."""
    bc_n = norm(breadcrumb).lower()
    lbl_n = norm(label).lower()
    full_n = f"{bc_n} {lbl_n}".strip()

    # ============ ASSETS — Cash ============
    if "cash on hand" in lbl_n and "bank" not in lbl_n:
        return "    CASH ON HAND", "auto"
    if "cash in bank" in lbl_n or "cash" in lbl_n and (
        "pnc" in lbl_n or "checking" in lbl_n or "bank" in lbl_n
    ):
        return "    CASH IN BANK", "auto"
    if "cash on hand and in bank" in full_n or (
        "cash" in lbl_n and ("hand" in lbl_n and "bank" in lbl_n)
    ):
        return "    CASH IN BANK", "auto"

    # ============ ASSETS — Receivables ============
    if "credit card receivable" in full_n or "credit card receivables" in full_n:
        return "    CREDIT CARD RECEIVABLE", "auto"
    if (
        "delivery sale receivable" in full_n
        or "delivery receivable" in full_n
        or "delivery sales receivable" in full_n
        or "receivable from deliveries" in lbl_n
    ):
        return "    DELIVERY SALES RECEIVABLE", "auto"
    if any(
        x in full_n
        for x in [
            "due (to) from", "due to from", "due to due from",
            "due to from affiliate", "due to from affiliates",
            "due to rom affiliates",   # actual QB typo
            "due from (to) affiliates", "due from to affiliates",
        ]
    ):
        return "    DUE FROM (TO) AFFILIATES", "auto"
    if "backup withholding" in lbl_n or "back up withholding" in lbl_n or "back up withholdng" in lbl_n:
        return "  IRS Back up withholding Recievable", "auto"
    if "prepaid" in full_n or "pre paid" in full_n or "pre-paid" in full_n:
        if "maryland" in lbl_n:
            return "    Maryland Tax", "auto"
        if "real estate" in lbl_n:
            return "    Real Estate Tax", "auto"
        if "nj" in lbl_n or "new jersey" in lbl_n:
            return "    NJ Annual filing fees", "auto"
        if "other assets" in bc_n and "expense" in lbl_n:
            return "    OTHER ASSETS", "auto"
        return "  Prepaid Epenses", "auto"
    if "inventory" in lbl_n and "delivery" not in lbl_n:
        return "    INVENTORY", "auto"
    if "gift card" in lbl_n and "payable" not in lbl_n:
        return "  Gift Card Receivables", "auto"
    if "loan receivable" in lbl_n or "loan recievable" in lbl_n:
        return "    LOAN RECEIVABLE FROM PARTNERS", "auto"

    # ============ FIXED ASSETS ============
    if "accumulated depreciation" in full_n or "acc deprec" in lbl_n:
        return "  Less: Accumulated depreciation", "auto"
    if "accum dep" in lbl_n or "accum. dep" in lbl_n:
        return "  Less: Accumulated depreciation", "auto"
    if "fixed assets" in full_n and "intangible" not in full_n:
        return "  Equipments", "auto"
    if "equipment" in lbl_n and "rental" not in lbl_n and "intangible" not in full_n:
        return "  Equipments", "auto"

    # ============ INTANGIBLES ============
    if "goodwill" in lbl_n and "accum" not in lbl_n and "amort" not in lbl_n:
        return "    GOODWILL", "auto"
    if "leasehold" in lbl_n and "accum" not in lbl_n and "amort" not in lbl_n:
        return "    LEASEHOLD IMP. (INTANGIBLE)", "auto"
    if "loan cost" in lbl_n or "deferred financing" in lbl_n:
        return "    DEFERRED FINANCING COSTS", "auto"
    if "closing cost" in lbl_n:
        return "    CLOSING COSTS", "auto"
    if "organization" in lbl_n and "accum" not in lbl_n:
        return "    ORGANIZATION COSTS", "auto"
    if ("franchise" in lbl_n or "franchies" in lbl_n) and (
        "asset" in full_n or "intangible" in full_n or "cost" in full_n
        or "fee" in lbl_n
    ) and "accum" not in lbl_n and "preservation" not in lbl_n:
        return "    FRANCHISE FEES", "auto"

    # Accumulated amortization sub-items
    if "accum" in lbl_n and "amort" in lbl_n:
        if "goodwill" in lbl_n:
            return "    ACCUM. AMORT. - GOODWILL", "auto"
        if "franchise" in lbl_n:
            return "    ACCUM. AMORT. - FRANCHISE FEES", "auto"
        if "leasehold" in lbl_n:
            return "    ACCUM. AMORT. - LEASEHOLD IMP. (INTANGIBLE)", "auto"
        if "organization" in lbl_n:
            return "    ACCUM. AMORT. - ORGANIZATION COSTS", "auto"
        if "financing" in lbl_n or "deferred" in lbl_n or "loan cost" in lbl_n:
            return "    ACCUM. AMORT. - DEFERRED FINANCING COSTS", "auto"
        if "closing" in lbl_n:
            return "    ACCUM. AMORT. - CLOSING COSTS", "auto"
        if "development" in lbl_n:
            return "    ACCUM. AMORT. - DEVELOPMENT RIGHTS", "auto"
        return "     Less: Accumulated Amortization", "auto"

    if "amortization" in full_n or "amortizatio" in full_n:
        return "     Less: Accumulated Amortization", "auto"
    if "development right" in lbl_n:
        return "    DEVELOPMENT RIGHTS", "auto"

    # ============ OTHER ASSETS ============
    if "ascentium" in lbl_n:
        return "  Ascentium Capital", "auto"
    if "deposit" in lbl_n and "purchase" in lbl_n:
        return "  Deposit for Purchase of Property", "auto"
    if "escrow" in lbl_n:
        return "    ESCROW DEPOSIT", "auto"
    if "work in progress" in lbl_n or "wip" in lbl_n:
        return "    WORK IN PROGRESS", "auto"
    if "investment in partnership" in lbl_n or "investment partnership" in lbl_n:
        return "  Investment in Partnership", "auto"
    if "investment in affiliate" in lbl_n:
        return "    INVESTMENT IN AFFILIATES", "auto"
    if "investment in other" in lbl_n or "investment in business" in lbl_n:
        return "    INVESTMENT IN OTHER BUSINESS", "auto"
    if "investment" in lbl_n:
        return "    INVESTMENT IN OTHER BUSINESS", "auto"
    if "surity" in lbl_n or "surety" in lbl_n:
        return "  Security Deposit to Surity Title Company", "auto"
    if "security deposit" in lbl_n and "payable" not in lbl_n:
        return "    SECURITY DEPOSIT", "auto"
    if lbl_n == "exchange" or "exchange" in lbl_n and (
        "assets" in full_n or "asset" in full_n
    ) and "liab" not in full_n:
        return "    EXCHANGES", "auto"

    # ============ LIABILITIES — Current ============
    if "accounts payable" in lbl_n:
        return "    ACCOUNTS PAYABLE", "auto"
    if "accrued interest" in lbl_n:
        return "   Accrued Interest", "auto"
    if "accrued" in lbl_n and "interest" not in lbl_n:
        return "    ACCRUED EXPENSES", "auto"
    if "sales tax payable" in lbl_n or lbl_n == "sales tax":
        return "    SALES TAX PAYABLE", "auto"
    if "rent payable" in lbl_n:
        return "   Rent payable", "auto"
    if "payroll tax" in lbl_n and "payable" in lbl_n:
        return "    PAYROLL LIABILITIES", "auto"
    if "payroll liabilities" in lbl_n or "payroll liabilities" in bc_n:
        return "    PAYROLL LIABILITIES", "auto"
    if "state tax payable" in lbl_n:
        if "maryland" in lbl_n or "md" in lbl_n:
            return "    STATE TAX PAYABLE", "auto"
        if "pa" in lbl_n or "pennsylvania" in lbl_n:
            return "    STATE TAX PAYABLE", "auto"
        return "    STATE TAX PAYABLE", "auto"
    if ("non resident tax" in lbl_n or "non res tax" in lbl_n) and "maryland" in lbl_n:
        return "    NON-RESIDENT TAX PAYABLE", "auto"
    if "md non resident" in lbl_n or "non resident tax" in lbl_n:
        return "    NON-RESIDENT TAX PAYABLE", "auto"
    if "net payroll" in lbl_n or ("payroll" in lbl_n and "check" in lbl_n):
        return "    NET PAYROLL CHECKS PAYABLE", "auto"
    if "meal tax" in lbl_n and "payable" in lbl_n:
        return "    MEAL TAXES PAYABLE", "auto"
    if "plk donation" in lbl_n:
        return "    PLK DONATION PAYABLE", "auto"
    if "tb foundation" in lbl_n:
        return "    TB FOUNDATION PAYABLE", "auto"
    if "insurance proceeds" in lbl_n and "payable" in lbl_n:
        return "    INSURANCE PROCEEDS PAYABLE", "auto"
    if "nj filing" in lbl_n or "nj annual" in lbl_n:
        return "    NJ FILING FEES PAYABLE", "auto"
    if "security deposit payable" in lbl_n:
        return "    SECURITY DEPOSIT PAYABLE", "auto"
    if "other current liabilities" in lbl_n:
        return "    OTHER CURRENT LIABILITIES", "auto"
    if "annual report" in lbl_n or "filing fee" in lbl_n or "filing fees" in lbl_n:
        return "STATE FILING FEES", "auto"

    # ============ LIABILITIES — Long-term Loans ============
    if ("bmo" in lbl_n or "harris" in lbl_n) and ("loan payable" in bc_n or "liabilities" in bc_n):
        return "    LOAN PAYABLE TO BANK", "auto"
    if "jatin desai" in lbl_n and ("loan payable" in bc_n or "liabilities" in bc_n):
        return "    LOAN PAYABLE TO PARTNERS", "auto"

    if "ppp" in lbl_n:
        return "    PPP LOAN PAYABLE", "auto"
    if "eidl" in lbl_n and "payable" in lbl_n:
        return "    LOAN PAYABLE TO SBA / EIDL", "auto"
    if "eidl" in lbl_n:
        return "    LOAN PAYABLE TO SBA / EIDL", "auto"
    if "sba" in lbl_n:
        return "    LOAN PAYABLE TO SBA / EIDL", "auto"
    if "loan payable" in lbl_n and "demand" in lbl_n:
        return "   Loan payable-Demand", "auto"
    if "loan payable" in lbl_n and (
        "ap n east" in lbl_n or "regal" in lbl_n or "ap northeast" in lbl_n
    ):
        return "   Loan payable-AP N East(Regal)", "auto"
    if "royal bank" in lbl_n:
        return "  Loan Payable-Royal bank", "auto"
    if "loan payable" in lbl_n and "partner" in lbl_n:
        return "    LOAN PAYABLE TO PARTNERS", "auto"
    if "loan payable" in lbl_n and "bank" in lbl_n:
        return "    LOAN PAYABLE TO BANK", "auto"
    if "loan payable" in lbl_n and ("other" in lbl_n and "current" in lbl_n):
        return "    LOAN PAYABLE - OTHERS (CURRENT)", "auto"
    if "loan payable" in lbl_n and "other" in lbl_n:
        return "    LOAN PAYABLE TO OTHERS", "auto"
    if "loan payable" in lbl_n:
        return "    LOAN PAYABLE TO OTHERS", "auto"

    # ============ EQUITY ============
    if "mid state pop" in lbl_n and "distribution" in bc_n:
        return "    LESS: DISTRIBUTION TO MANAGEMENT ENTITY", "auto"

    if (
        "partner s capital" in full_n
        or "partner capital" in full_n
        or "partner's capital" in full_n
    ):
        if "distribution" in full_n:
            return "Distribution paid to Partner", "auto"
        if "profit for the year" in lbl_n or "current year profit" in lbl_n or "net profit" in lbl_n:
            return "    ADD: CURRENT YEAR PROFIT / (LOSS)", "auto"
        if "additional capital" in lbl_n or "capital contribution" in lbl_n:
            return "    ADD: CAPITAL CONTRIBUTION - MANAGEMENT ENTITY", "auto"
        if "beginning" in lbl_n:
            return "    BEGINNING CAPITAL (PARTNERS' CAPITAL - BEGINNING)", "auto"
        return "    BEGINNING CAPITAL (PARTNERS' CAPITAL - BEGINNING)", "auto"

    if "retained earnings" in lbl_n:
        return "    RETAINED EARNINGS", "auto"

    if "current year profit" in lbl_n or "net income" in lbl_n or "net profit" in lbl_n:
        return "    ADD: CURRENT YEAR PROFIT / (LOSS)", "auto"

    if "capital contribution" in lbl_n:
        return "    ADD: CAPITAL CONTRIBUTION - MANAGEMENT ENTITY", "auto"

    if "distribution" in full_n and "partner" in full_n:
        if "management" in lbl_n:
            return "    LESS: DISTRIBUTION TO MANAGEMENT ENTITY", "auto"
        return "    LESS: DISTRIBUTION TO PARTNER 1", "auto"

    if "beginning capital" in lbl_n or "beginning partner" in lbl_n:
        return "    BEGINNING CAPITAL (PARTNERS' CAPITAL - BEGINNING)", "auto"

    return None, "REVIEW"
