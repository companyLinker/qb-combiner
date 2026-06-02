"""Auto-mapping rules for QuickBooks accounts → target combination workbook line items.

Two main functions:
    map_pnl(breadcrumb, label) → (target_line, confidence)
    map_bs(breadcrumb, label)  → (target_line, confidence)

confidence is 'auto' if a rule fired, 'REVIEW' if no rule matched.
target_line == '__SKIP__' means this is a structural pseudo-leaf (e.g., 'Gross Profit'
appears as a leaf in some QB exports because of a display setting) — should not flow.

The target line strings include intentional leading/trailing whitespace that matches
the exact text in the target template column A. Do NOT strip these.

Coverage observed on the May 2026 dataset (22 entities, FY 2025):
  - P&L: 157/158 auto-mapped (99%)
  - BS:  465/499 auto-mapped (93%)

Items that consistently stay REVIEW:
  - Individual partner names under loan accounts
  - LLC-specific sub-accounts under PARTNER'S CAPITAL
  - One-off items (CLOSING COST, INSURANCE PROCEEDS RECEIVABLE, RETAINED EARNINGS)
"""

import re


def norm(s):
    """Normalize a label for fuzzy matching: lowercase, alphanumeric+space only."""
    if not s:
        return ""
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s)


def map_pnl(breadcrumb, label):
    """Return (target_line, confidence) for a P&L leaf account."""
    bc_n = norm(breadcrumb)
    lbl_n = norm(label)

    # Skip "Gross Profit" leaf — it's actually a subtotal pseudo-row in some QB layouts
    if lbl_n == "gross profit":
        return "__SKIP__", "auto"
    if "operating expenses other" in lbl_n:
        return "    MISC EXP", "auto"
    if "labor cost other" in lbl_n:
        return "   KITCHEN LABOR", "auto"

    # ============ INCOME ============
    if "food sales" in bc_n or bc_n.endswith(" income") or bc_n == "ordinary income expense > income":
        if any(x in lbl_n for x in ["non taxable sales", "taxable sales",
                                    "delivery sales refund", "merchant card",
                                    "online sales", "delivery sale refund"]):
            return "RETAIL SALES", "auto"
    if "rebate" in lbl_n:
        return "Rebate Income", "auto"
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
            return "Management fees  Income", "auto"
        else:
            return " MANAGEMENT FEES", "auto"
    if "rental income" in lbl_n or ("rent" in lbl_n and "income" in lbl_n):
        return "Rental Income", "auto"
    if "erc" in lbl_n:
        if "interest" in lbl_n:
            return "ERC Interest Income", "auto"
        return "ERC Credit Income", "auto"
    if "ppp" in lbl_n:
        return "Non Taxable PPP Income", "auto"
    if "discount" in lbl_n and ("receive" in lbl_n or "recievable" in lbl_n):
        return "Rebate Income", "auto"
    if "misc" in lbl_n and "income" in lbl_n:
        return "Rebate Income", "auto"
    if "interest" in lbl_n and "income" in lbl_n and "expense" not in lbl_n:
        return "INTEREST INCOME", "auto"

    # ============ COGS ============
    if "franchise fees" in bc_n:
        if any(x in lbl_n for x in ["ad fund", "advt", "royalty", "advertising"]):
            return " FRANCHISE & ADVT FEES", "auto"

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
            return "   TRAINING", "auto"
        if "bonus" in lbl_n and "performance" not in lbl_n:
            return "   BONUS", "auto"
        if "performance" in lbl_n:
            return "  PERFORMANCE BONUS", "auto"
        if "worker" in lbl_n or "workmen" in lbl_n:
            return " INSURANCE EXP & WORKERS COMP", "auto"
        if "payroll tax" in bc_n or "social security" in lbl_n or "unemployment" in lbl_n or "fica" in lbl_n:
            return "   PAYROLL TAXES", "auto"
        if "fed" in lbl_n and "unemployment" in lbl_n:
            return "   PAYROLL TAXES", "auto"

    if "purchases" in bc_n or "purchase" in bc_n:
        if "opening" in lbl_n:
            return "Opening Inventory", "auto"
        if "closing" in lbl_n or "ending" in lbl_n:
            return "Ending Inventory", "auto"
        return "  PURCHASE", "auto"

    # ============ EXPENSES ============
    if "auto expenses" in bc_n:
        return " AUTO EXPENSE", "auto"
    if "cleaning expenses" in bc_n:
        return " CLEANING EXP", "auto"
    if "insurance" in bc_n.lower():
        return " INSURANCE EXP & WORKERS COMP", "auto"
    if "occupancy" in bc_n:
        bc_parts = [p.strip().lower() for p in breadcrumb.split(" > ")]
        last_part = bc_parts[-1] if bc_parts else ""
        if "rent" in last_part or "rent" in lbl_n:
            return " RENT ", "auto"
        if "repair" in bc_n or "repair" in lbl_n or "maint" in lbl_n:
            return " REPAIR & MAINT", "auto"
        if "utilities" in bc_n or any(x == lbl_n or lbl_n.startswith(x)
                                       for x in ["electricity", "water", "gas", "heating", "cooling", "sewer"]):
            return " UTILITIES", "auto"
    if "uniform" in lbl_n:
        return " UNIFORM ", "auto"
    if any(x in lbl_n for x in ["restaurant supplies", "restaurant supply",
                                 "restaurant supp", "kitchen smallware", "smallware"]):
        return " RESTAURANT SUPP", "auto"

    # Delivery aggregators
    if "delivery charges" in bc_n or any(x in lbl_n for x in [
        "doordash", "grubhub", "ubereats", "uber eats", "otter",
        "ezcarter", "ez carter", "gift card charges"
    ]):
        if "income" not in bc_n and "income" not in lbl_n:
            return "DELIVERY CHARGES", "auto"

    # Specific common operating expense items matching regardless of exact breadcrumb
    if "digital transaction" in lbl_n or "ordering tech fee" in lbl_n or "guest care fee" in lbl_n:
        return "POPEYES DIGITAL TRANSACTION FEE", "auto"
    if "food" in lbl_n and "employee" in lbl_n:
        return "FOOD FOR EMPLOYEES", "auto"
    if "loan fee" in lbl_n or "loan fees" in lbl_n:
        return "LOAN FEES", "auto"

    if "operating expenses" in bc_n.lower():
        if "credit card" in lbl_n or "merchant fee" in lbl_n or "guest care" in lbl_n:
            return "    CREDIT CARD CHARGES", "auto"
        if "license" in lbl_n or "permit" in lbl_n:
            return "    LICENCE AND PERMITS", "auto"
        if "accounting" in lbl_n:
            return "    ACCOUNTING FEES", "auto"
        if "bank service" in lbl_n or "bank charge" in lbl_n:
            return "    BANK CHARGES", "auto"
        if "401" in lbl_n or "401k" in lbl_n:
            return "    PAYROLL EXPENSE", "auto"
        if "payroll" in lbl_n and "expense" in lbl_n:
            return "    PAYROLL EXPENSE", "auto"
        if "shortage" in lbl_n or "cash short" in lbl_n:
            return "    SHORTAGES", "auto"
        if "office" in lbl_n:
            return "    Office Expenses", "auto"
        if "security" in lbl_n or "alarm" in lbl_n:
            return "    Security and Alarm", "auto"
        if "donation" in lbl_n or "charity" in lbl_n:
            return "    DONATIONS", "auto"
        if "equipment rental" in lbl_n:
            return "    EQUIPMENT RENTAL", "auto"
        if "popeyes" in lbl_n or "service check" in lbl_n:
            return "   POPEYES SERVICE CHECK", "auto"
        if "professional" in lbl_n or "legal" in lbl_n:
            return "   PROFESSIONAL FEES", "auto"
        if "convention" in lbl_n or "convension" in lbl_n:
            return "   CONVENSION EXPENSES", "auto"
        if "computer" in lbl_n or "software" in lbl_n:
            return "   COMPUTER EXPENSES", "auto"
        if "advertis" in lbl_n or "promotion" in lbl_n:
            return "   Advertising and Promotion", "auto"
        if "non recurring" in lbl_n:
            return "  Non recurring expenses", "auto"
        if "penalty" in lbl_n or "annual report" in lbl_n or "suspense" in lbl_n:
            return "    MISC EXP", "auto"
        if "misc" in lbl_n and "tax" not in lbl_n:
            return "    MISC EXP", "auto"
        # Catch-all in operating expenses bucket
        return "    MISC EXP", "auto"

    # ============ TAXES ============
    if "taxes" in bc_n.lower() or "tax" in lbl_n.lower():
        if "real estate" in lbl_n or "property tax" in lbl_n:
            if "maryland" in lbl_n:
                return "    -PROPERTY TAX-MARYLAND", "auto"
            return "     RE TAX AND PROPRTY TAX", "auto"
        if "annual report" in lbl_n or "filing fee" in lbl_n or "filing fees" in lbl_n or "nj" in lbl_n or "new jersey" in lbl_n:
            return "STATE FILING FEES", "auto"
        if "restaurant" in lbl_n or "misc" in lbl_n or "other" in lbl_n:
            return "Misc Taxes", "auto"
        if "state" in lbl_n:
            return "     STATE TAXES", "auto"
        if "mercantile" in lbl_n or "local" in lbl_n or "bpt" in lbl_n or "school tax" in lbl_n:
            return "   LOCAL MERCANTILE TAX", "auto"
        if "gross receipt" in lbl_n:
            return "    Gross receipt tax", "auto"

    # ============ INTEREST EXPENSES ============
    if ("interest" in bc_n.lower() and "expense" in bc_n.lower()) or "interest expenses" in lbl_n:
        return "INTEREST EXPENSE", "auto"

    # ============ NON-CASH / OTHER ============
    if "depreciation" in lbl_n:
        return "DEPRECIATION", "auto"
    if "amortization" in lbl_n or "amortisation" in lbl_n:
        return "AMORTIZATION", "auto"

    if "profit sharing" in bc_n.lower() or "profit sharing" in lbl_n.lower():
        return "Profit Transfer to AP North", "auto"
    if "profit transfer" in lbl_n or ("transfer" in lbl_n and "north" in lbl_n):
        return "Profit Transfer to AP North", "auto"

    return None, "REVIEW"


def map_bs(breadcrumb, label):
    """Return (target_line, confidence) for a Balance Sheet leaf account."""
    bc_n = norm(breadcrumb).lower()
    lbl_n = norm(label).lower()
    full_n = f"{bc_n} {lbl_n}".strip()

    # ============ ASSETS ============
    if "cash on hand and in bank" in full_n or (
        "cash" in lbl_n and ("hand" in lbl_n or "pnc" in lbl_n or "bank" in lbl_n or "checking" in lbl_n)
    ):
        return "  Cash on hand and in bank", "auto"
    if "credit card receivable" in full_n or "credit card receivables" in full_n:
        return "  Credit card receivable", "auto"
    if "delivery sale receivable" in full_n or "delivery receivable" in full_n or "delivery sale receivables" in full_n:
        return "  Receivable from Deliveries", "auto"
    if any(x in full_n for x in [
        "due (to) from", "due to from", "due to due from",
        "due to from affiliate", "due to from affiliates",
        "due to rom affiliates"  # actual QB typo
    ]):
        return "  Due from/(Due to) Affiliates", "auto"
    if "backup withholding" in lbl_n or "back up withholding" in lbl_n or "back up withholdng" in lbl_n:
        return "  IRS Back up withholding Recievable", "auto"
    if "prepaid" in full_n:
        if "maryland" in lbl_n:
            return "    Maryland Tax", "auto"
        if "real estate" in lbl_n:
            return "    Real Estate Tax", "auto"
        if "nj" in lbl_n or "new jersey" in lbl_n:
            return "    NJ Annual filing fees", "auto"
        return "  Prepaid Epenses", "auto"
    if "inventory" in lbl_n and "delivery" not in lbl_n:
        return "  Inventory", "auto"
    if "gift card" in lbl_n:
        return "  Gift Card Receivables", "auto"
    if lbl_n == "exchange" or (
        "exchange" in lbl_n and ("assets" in full_n or "asset" in full_n) and "liab" not in full_n
    ):
        return "  Exchange", "auto"
    if "loan receivable" in lbl_n or "loan recievable" in lbl_n:
        return "  Loan receivable", "auto"

    # ============ FIXED ASSETS & INTANGIBLES ============
    if "accumulated depreciation" in full_n or "acc deprec" in lbl_n:
        return "  Less: Accumulated depreciation", "auto"
    if "amortization" in full_n or "amortizatio" in full_n:
        return "     Less: Accumulated Amortization", "auto"
    if "fixed assets" in full_n and "intangible" not in full_n:
        return "  Equipments", "auto"

    # ============ INTANGIBLES ============
    if "goodwill" in lbl_n:
        return "     Goodwill", "auto"
    if "leasehold" in lbl_n:
        return "     Leasehold Improvemrnts", "auto"
    if "loan cost" in lbl_n:
        return "     Loan Cost", "auto"
    if "organization" in lbl_n:
        return "     Organization expenses", "auto"
    if ("franchise" in lbl_n or "franchies" in lbl_n) and (
        "asset" in full_n or "intangible" in full_n or "cost" in full_n
    ):
        return "     Franchise fees", "auto"

    # ============ OTHER ASSETS ============
    if "ascentium" in lbl_n:
        return "  Ascentium Capital", "auto"
    if "deposit" in lbl_n and "purchase" in lbl_n:
        return "  Deposit for Purchase of Property", "auto"
    if "investment in partnership" in lbl_n or "investment partnership" in lbl_n:
        return "  Investment in Partnership", "auto"
    if "surity" in lbl_n or "surety" in lbl_n:
        return "  Security Deposit to Surity Title Company", "auto"
    if "security deposit" in lbl_n:
        return "  Security Deposit ", "auto"
    if "development right" in lbl_n:
        return "  Development Rights", "auto"

    # ============ LIABILITIES ============
    if "accounts payable" in lbl_n:
        return "    Accounts Payable", "auto"
    if "accrued interest" in lbl_n:
        return "   Accrued Interest", "auto"
    if "accrued" in lbl_n and "interest" not in lbl_n:
        return "   Accrued expenses", "auto"
    if "state tax payable" in lbl_n:
        if "maryland" in lbl_n or "md" in lbl_n:
            return "   State tax payable-Maryland", "auto"
        if "pa" in lbl_n or "pennsylvania" in lbl_n:
            return "   State tax payable-PA", "auto"
    if ("non resident tax" in lbl_n or "non res tax" in lbl_n) and "maryland" in lbl_n:
        return "   State tax payable-Maryland", "auto"
    if "md non resident" in lbl_n:
        return "   State tax payable-Maryland", "auto"
    if "net payroll" in lbl_n or ("payroll" in lbl_n and "check" in lbl_n):
        return "   Net payroll checks payable", "auto"
    if "sales tax payable" in lbl_n or lbl_n == "sales tax":
        return "   Sales tax payable", "auto"
    if "rent payable" in lbl_n:
        return "   Rent payable", "auto"
    if "payroll tax" in lbl_n and "payable" in lbl_n:
        return "   Payroll taxes payable", "auto"
    if "payroll liabilities" in lbl_n:
        return "   Payroll taxes payable", "auto"
    if "annual report" in lbl_n or "filing fee" in lbl_n or "filing fees" in lbl_n or "nj" in lbl_n or "new jersey" in lbl_n:
        return "STATE FILING FEES", "auto"
    if "ppp" in lbl_n:
        return "  Loan payable-SBA PPP loan", "auto"
    if "eidl" in lbl_n:
        return "  Loan payable-EIDL loan", "auto"
    if "loan payable" in lbl_n and "demand" in lbl_n:
        return "   Loan payable-Demand", "auto"
    if "loan payable" in lbl_n and ("ap n east" in lbl_n or "regal" in lbl_n or "ap northeast" in lbl_n):
        return "   Loan payable-AP N East(Regal)", "auto"
    if "royal bank" in lbl_n:
        return "  Loan Payable-Royal bank", "auto"
    if "loan payable" in lbl_n and "partner" in lbl_n:
        return "Loan Payable to Partners", "auto"

    # ============ EQUITY (Partner's Capital) ============
    if "partner s capital" in full_n or "partner capital" in full_n or "partner's capital" in full_n:
        if "distribution" in full_n:
            return "Distribution paid to Partner", "auto"
        if "profit for the year" in lbl_n or "current year profit" in lbl_n or "net profit" in lbl_n:
            return "CURRENT YEAR PROFIT", "auto"
        if "additional capital" in lbl_n:
            return "ADDITIONAL CAPITAL", "auto"
        # Everything else under PARTNER'S CAPITAL rolls to opening balance
        return " Partner's capital-beging", "auto"
    if "distribution" in full_n:
        return "Distribution paid to Partner", "auto"

    return None, "REVIEW"
