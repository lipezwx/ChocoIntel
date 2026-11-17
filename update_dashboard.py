"""
Stock Intelligence - dashboard updater
======================================
Reads two Excel workbooks (stock control + warehouse positions), recalculates every KPI,
table, chart dataset and the daily transfer plan, and rewrites them into index.html.

Quick start (fictional demo data):
    python generate_sample_data.py
    python update_dashboard.py --as-of 2026-11-17

Your own data: point --xlsx / --sat-xlsx at workbooks with the same sheet layout
(see README) and adjust the business rules in section 1 below.
"""
import argparse
import datetime as dt
import html as htmllib
import json
import math
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from collections import defaultdict

math_inf = float("inf")

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is missing.  Run:  pip install openpyxl")

# =============================================================================
# 1. SETTINGS  -  edit here (paths + business rules)
# =============================================================================
XLSX_PATH = r"sample_data/Stock Control.xlsx"
HTML_PATH = r"index.html"
# Warehouse-position workbook ("Position Saturation"). Optional: when the file is
# missing, the position data already inside the HTML is kept as it is.
SAT_XLSX_PATH = r"sample_data/Position Saturation.xlsx"
SAT_SHEET = None             # None = first sheet

SHEET_DC = "In Stock - DC"
SHEET_EXT = "In Stock - Externals"
SHEET_USAGE = "Usage Material"
SHEET_ORDERS = "Order Info"

DC_SLOC = "1000"             # storage location that IS the main distribution centre (DC)
DC_QTY_COLUMN = "total stock"          # or "available stock"
# Which date defines "idle" for Slow Movers: "last movement" or "gr date"
# Using "last movement" by default; fallback to "Bin" column raw data if dates are invalid (future etc.)
IDLE_DATE_COLUMN = "last movement"
# Externals sheet: stock buckets that count as stock (all of them by default)
EXT_QTY_COLUMNS = ["unrestricted", "transit and transfer", "in quality insp",
                   "restricteduse stock", "blocked", "returns"]
EXT_VALUE_COLUMNS = ["value unrestricted", "val in transtfr", "value in qualinsp",
                     "value restricted", "value blockedstock"]
# Only these external storage locations feed the DC warehouse.
# 3001 = North Hub, 3002 = East Hub.  Other locations (e.g. 3009) are reported as
# source risks instead of being silently mixed into the daily transfer.
TRANSFER_FROM_SLOCS = ["3001", "3002"]

# --- Transfer logistics ------------------------------------------------------
# Display names used by the dashboard for the external sources.
TRANSFER_SOURCE_NAMES = {"3001": "North Hub", "3002": "East Hub"}
# Two free transfers per day (morning + afternoon); each normal trip has 24
# pallet positions.  Anything above that needs a paid extra trip (28 positions).
TRANSFER_NORMAL_SPOTS = 24
TRANSFER_EXTRA_SPOTS = 28
TRANSFER_FREE_TRIPS = 2

# Keep these minimum safety levels in DC, in pallets.  The target is
# "3-day demand + safety pallets" so the safety buffer still exists after the
# next three days of consumption.
SAFETY_STOCK_PALLETS = {
    "2980": 8,
    "3034": 8,
    "1141": 5,
    "2578": 5,
    "2888": 2,
    "1567": 2,
    "2802": 2,
    "3100": 2,
}

# --- Business rules (weeks of cover = stock / weekly demand) ----------------
DEMAND_WEEKS = 4             # weekly demand = requirements in the next N weeks / N
CRITICAL_BELOW = 1           # weeks  < 1          -> Critical
LOW_BELOW = 4                # weeks  1 .. 4       -> Low
OVERSTOCK_ABOVE = 12         # weeks  > 12         -> Overstocked  (4..12 = Adequate)
TARGET_WEEKS = 12            # availability score = min(100, weeks / TARGET * 100)
NO_DEMAND_WEEKS = 999        # weeks shown when a material has no demand
TRANSFER_DAYS = 3            # "3d transfer" window (today + next 2 days)
SLOW_MOVER_DAYS = 60         # idle longer than this = slow mover
ACTIVE_MATERIALS = "dc"   # "dc"  = materials with stock in DC warehouse
                             # "any"    = DC + external stock
                             # "stock+demand" = anything with stock or demand

# --- Table / chart sizes (match the current dashboard) -----------------------
TOP_CRITICAL = 30
TOP_OVERSTOCK = 30
TOP_AVAIL = 50
TOP_DEMAND = 10
TOP_TRANSFER = 15
TOP_RISK = 10
TOP_SLOW = 10
TOP_CHART = 10

# --- Material exclusions (applied to KPIs, charts and every table) ----------
# Finished/semi-finished products use codes with letters or "-" (FG-BAR-120, FG-EZ-210);
# raw materials are numeric (1141, 70010133).  This also removes packaging/films
# whose codes carry a suffix.
EXCLUDE_FINISHED_GOODS = True
EXCLUDE_CODE_RE = re.compile(r"[A-Za-z-]")
# Melted materials live in tanks and are not part of the pallet flow.  They can
# have numeric codes or alphanumeric codes (MLT100).  Any description
# containing the word MELT / MELTED is removed as well.
EXCLUDE_MELTING_CODES = True
MELTING_CODES = {"MLT100"}     # explicit codes, on top of the description rule
MELTING_DESC_RE = re.compile(r"\bMELT(?:ED)?\b", re.I)
EXCLUDE_MATERIALS = {
    "90000001",   # pallets - handled with a different logic
    "1999",       # bulk sweetener - follows a completely different logic
}                 # also removed from the warehouse-position map and its saturation %

# --- Warehouse positions (saturation workbook) -------------------------------
POOL_LABELS = ["drive-in bulk", "drive-in cases", "roller lanes", "racks"]   # shared positions
BLOCKED_LABEL = "blocked"      # excluded from dedicated saturation (as in the sheet formula)
IGNORE_POSITION_LABELS = {"total", "totals", "sum", "subtotal", "grand total"}   # sheet rows that are not materials
DEFAULT_KG_PER_POSITION = 1000    # used when a material has no kg-per-position of its own
TREND_POINTS = 40                 # last N saturation readings shown in the trend chart
NAME_MAX = 28                     # description length shown in the position map

BACKUP = True                # keep <dashboard>.bak of the previous version


# =============================================================================
# 2. HELPERS
# =============================================================================
def norm_header(v):
    """'Val. in Trans./Tfr' -> 'val in transtfr' (lowercase, only letters/digits/space)."""
    s = str(v or "").replace("\n", " ").lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_material(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if re.fullmatch(r"\d+(\.0+)?", s):
        s = s.split(".")[0].lstrip("0") or "0"
    return s


def is_excluded(m, description=""):
    """True for materials that must not appear anywhere in the dashboard."""
    if EXCLUDE_FINISHED_GOODS and EXCLUDE_CODE_RE.search(m):
        return True
    if EXCLUDE_MELTING_CODES:
        if m in MELTING_CODES:
            return True
        if description and MELTING_DESC_RE.search(str(description)):
            return True
    return m in EXCLUDE_MATERIALS


def to_num(v):
    """Numbers from Excel (real numbers, or text in BR/US format)."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y", "%m/%d/%Y"):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def find_sheet(wb, name):
    """Worksheet by name, ignoring case and stray spaces.  Resets the stored dimensions: ERP/BI exports
    often declare a wrong used-range and read-only mode would then silently stop after a few rows."""
    match = next((n for n in wb.sheetnames if n.strip().lower() == name.strip().lower()), None)
    if match is None:
        raise KeyError(f"Sheet '{name}' not found. Sheets in file: {wb.sheetnames}")
    ws = wb[match]
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()
    return ws


def open_workbook_copy(path):
    """Open a private copy (the original is usually open in Excel / locked on a share).
    Returns (workbook, temp_dir); the caller closes the workbook and removes the folder."""
    tmp_dir = tempfile.mkdtemp(prefix="stock_dash_")
    try:
        tmp = os.path.join(tmp_dir, "copy.xlsx")
        shutil.copy2(path, tmp)
        return openpyxl.load_workbook(tmp, read_only=True, data_only=True), tmp_dir
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise


def read_sheet(wb, name):
    """Return list of dicts keyed by normalised header. Header row = first non-empty row."""
    ws = find_sheet(wb, name)
    rows = ws.iter_rows(values_only=True)
    header = None
    for r in rows:
        if r and any(c is not None for c in r):
            header = [norm_header(c) for c in r]
            break
    if header is None:
        return []
    out = []
    for r in rows:
        if not r or all(c is None for c in r):
            continue
        out.append({h: c for h, c in zip(header, r) if h})
    return out


def col(row, *names):
    for n in names:
        if n in row:
            return row[n]
    return None


def require_columns(rows, sheet, needed):
    if not rows:
        print(f"  ! Sheet '{sheet}' is empty")
        return
    missing = [n for n in needed if n not in rows[0]]
    if missing:
        sys.exit(f"Sheet '{sheet}' is missing column(s): {missing}\n"
                 f"Columns found: {sorted(rows[0].keys())}")


def fmt_int(v):
    return f"{int(round(v)):,}"


def fmt_weeks(v):
    """Weeks of cover: text when there is no demand, one decimal below 10 (so 0.4 is not shown as 0)."""
    if v is None:
        return "No demand"
    return f"{v:.1f}" if v < 10 else fmt_int(v)


# English month names written by hand: strftime('%B') follows the PC language (pt-BR -> 'outubro')
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


def fmt_date_long(d):
    return f"{d.day} {MONTHS[d.month - 1]} {d.year}"


def fmt_date_short(d):
    return f"{d.day} {MONTHS[d.month - 1][:3]}"


def esc(s):
    return htmllib.escape(str(s), quote=False)


def plain(s):
    """Lowercase, accents removed: 'Shared Saturation' -> 'shared saturation'."""
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


def to_pct(v):
    """Percent from 0.634, 63.4, '63,4%' or '63.4%'  ->  63.4 (None when empty)."""
    if v is None or v == "":
        return None
    if isinstance(v, str):
        x = to_num(v.replace("%", ""))
    else:
        x = float(v)
        if abs(x) <= 1.5:
            x *= 100
    return round(x, 2)


# =============================================================================
# 3. LOAD + CALCULATE  (Stock Control workbook)
# =============================================================================
def load_and_calculate(xlsx_path, today):
    wb, tmp_dir = open_workbook_copy(xlsx_path)
    try:
        return _calculate(wb, today)
    finally:
        wb.close()                                   # always release the file (Windows keeps it locked otherwise)
        shutil.rmtree(tmp_dir, ignore_errors=True)


def fefo_picks(ext_batches, batch_expiry, m, need):
    """Pick external batches for one material, first-expired-first-out.

    The ideal source is the external location that can cover the whole need
    with its oldest stock; when no single location covers it, the location
    with the most stock is used.  That source is consumed first, then the
    other locations top up the remainder in FEFO order.  Batches with a
    known expiry come first; unknown expiry falls back to batch number.
    """
    batches = []
    for (sloc, batch), qty in ext_batches.get(m, {}).items():
        batches.append(dict(sloc=sloc, batch=batch, qty=qty,
                            exp=batch_expiry.get((m, batch))))

    def bkey(b):
        return (0, int(b), "") if b.isascii() and b.isdigit() else (1, 0, b)

    batches.sort(key=lambda b: (b["exp"] is None, b["exp"] or dt.date.max,
                                bkey(b["batch"]), b["sloc"]))
    src_total = defaultdict(float)
    src_first = {}
    for b in batches:
        src_total[b["sloc"]] += b["qty"]
        src_first.setdefault(b["sloc"], b)

    covering = [s for s, q in src_total.items() if q >= need]
    if covering:                       # earliest oldest stock that can do the whole job
        ideal = min(covering, key=lambda s: (src_first[s]["exp"] is None,
                                             src_first[s]["exp"] or dt.date.max,
                                             bkey(src_first[s]["batch"]), s))
    elif src_total:                    # otherwise the biggest source
        ideal = min(src_total, key=lambda s: (-src_total[s],
                                              src_first[s]["exp"] is None,
                                              src_first[s]["exp"] or dt.date.max, s))
    else:
        ideal = None

    ordered = [b for b in batches if b["sloc"] == ideal] + \
              [b for b in batches if b["sloc"] != ideal]
    picks, remaining, ideal_qty = [], need, 0.0
    for b in ordered:
        if remaining <= 0:
            break
        take = min(b["qty"], remaining)
        picks.append(dict(batch=b["batch"], sloc=b["sloc"], qty=round(take, 1),
                          available=round(b["qty"], 1),
                          exp=b["exp"].isoformat() if b["exp"] else None))
        if b["sloc"] == ideal:
            ideal_qty += take
        remaining -= take
    return picks, ideal, round(ideal_qty, 1)


def _calculate(wb, today):
    dc_rows = read_sheet(wb, SHEET_DC)
    ext_rows = read_sheet(wb, SHEET_EXT)
    use_rows = read_sheet(wb, SHEET_USAGE)
    try:
        ord_rows = read_sheet(wb, SHEET_ORDERS)
    except KeyError:
        ord_rows = []
    

    require_columns(dc_rows, SHEET_DC, ["material", DC_QTY_COLUMN])
    require_columns(ext_rows, SHEET_EXT, ["material", "storage location"])
    require_columns(use_rows, SHEET_USAGE, ["material", "requirements date", "difference quantity"])

    desc = {}

    def remember(mat, d):
        if mat and d and mat not in desc:
            desc[mat] = str(d).strip()

    # ---- DC warehouse ----------------------------------------------------
    # The data rows in "In Stock - DC" are one column to the left of the
    # header after "Material Description" (the Plant column is not filled).
    # Verified against the workbook:
    #   0 Material | 1 Description | 2 Batch | 3 Special Stock | 4 Storage Location
    #   5 Storage Type | 6 Storage Bin | 7 GR Date | 8 Stock Category
    #   9 Available stock | 10 Total Stock | 11 Number of quants | 12 Last movement
    #   13 Time | 14 Base unit | 15 Expiry date
    dc_stock = defaultdict(float)
    fefo_by_mat = defaultdict(list)          # batch-level stock currently in DC
    idle_weight = defaultdict(lambda: [0.0, 0.0])   # material -> [days x kg, kg]
    batch_expiry = {}                        # (material, batch) -> expiry date
    qty_idx = 9 if "available" in DC_QTY_COLUMN.lower() else 10      # Available stock | Total Stock
    n_rows = n_qty_ok = 0
    for r in find_sheet(wb, SHEET_DC).iter_rows(values_only=True):
        m = norm_material(r[0] if len(r) > 0 else None)
        if not m or m.lower() == "material":               # blank row / header row
            continue
        n_rows += 1
        q_cell = r[qty_idx] if len(r) > qty_idx else None
        if isinstance(q_cell, (int, float)) or (q_cell is not None and re.fullmatch(r"[-\d.,\s]+", str(q_cell))):
            n_qty_ok += 1
        if is_excluded(m, r[1] if len(r) > 1 else ""):
            continue
        qty = to_num(q_cell)
        dc_stock[m] += qty
        batch = str(r[2] or "").strip() if len(r) > 2 else ""
        sloc = norm_material(r[4] if len(r) > 4 else None)
        gr = to_date(r[7] if len(r) > 7 else None)
        exp = to_date(r[15] if len(r) > 15 else None)
        if batch and exp:
            batch_expiry[(m, batch)] = exp
        if qty <= 0:
            continue
        fefo_by_mat[m].append(dict(batch=batch, sloc=sloc, qty=qty,
                                   gr=gr.isoformat() if gr else None,
                                   exp=exp.isoformat() if exp else None))
        # Idle = days since the last movement (fallback: GR date).  Expiry
        # dates must never be used here: they are in the future.
        d = to_date(r[12] if len(r) > 12 else None) or gr
        if d is not None:
            days = (today - d).days
            if days >= 0:
                idle_weight[m][0] += days * qty
                idle_weight[m][1] += qty

    if n_rows and n_qty_ok / n_rows < 0.8:
        print(f"  ! '{SHEET_DC}': only {n_qty_ok} of {n_rows} rows have a number in column {qty_idx + 1}. "
              f"The column layout may have changed - check the comment above (stock was probably read from the wrong column).")

    # ---- Externals (ERP stock list, all storage locations) ----------------------
    ext_stock = defaultdict(float)       # stock outside DC that can be transferred
    ext_at_dc_sloc = defaultdict(float)
    ext_all = defaultdict(float)         # all stock outside DC
    ext_batches = defaultdict(lambda: defaultdict(float))   # material -> (sloc, batch) -> qty
    ext_qty_unit = {}                    # material -> base unit seen in the externals sheet
    price_val = defaultdict(float)       # $per material (for its own price)
    price_qty = defaultdict(float)
    kg_val = defaultdict(float)          # KG lines only, for the fallback price per kg
    kg_qty = defaultdict(float)
    value_total = 0.0
    price_qty_cols = [c for c in EXT_QTY_COLUMNS if c != "returns"]   # value columns have no 'returns'
    for r in ext_rows:
        m = norm_material(r.get("material"))
        if not m or is_excluded(m, r.get("material description")):
            continue
        remember(m, r.get("material description"))
        sloc = norm_material(r.get("storage location"))
        qty = sum(to_num(r.get(c)) for c in EXT_QTY_COLUMNS)
        val = sum(to_num(r.get(c)) for c in EXT_VALUE_COLUMNS)
        value_total += val
        pq = sum(to_num(r.get(c)) for c in price_qty_cols)      # same buckets as the value columns (no 'returns')
        price_val[m] += val
        price_qty[m] += pq
        if str(r.get("base unit of measure") or "KG").strip().upper() == "KG":
            kg_val[m] += val
            kg_qty[m] += pq
        ext_qty_unit.setdefault(m, str(r.get("base unit of measure") or "").strip().upper())
        if sloc == DC_SLOC:
            ext_at_dc_sloc[m] += qty
        else:
            ext_all[m] += qty
            if (not TRANSFER_FROM_SLOCS or sloc in TRANSFER_FROM_SLOCS) and qty > 0:
                ext_stock[m] += qty
                batch = str(r.get("batch") or "").strip()
                ext_batches[m][(sloc, batch)] += qty

    # DC-side stock: WM sheet when the material is there, else MB52 qty at 9999
    dc_side = dict(dc_stock)
    for m, q in ext_at_dc_sloc.items():
        dc_side.setdefault(m, q)

    # ---- Demand -------------------------------------------------------------
    total_cons = defaultdict(float)
    window_cons = defaultdict(float)
    cons_3d = defaultdict(float)
    win_end = today + dt.timedelta(days=DEMAND_WEEKS * 7)
    tr_end = today + dt.timedelta(days=TRANSFER_DAYS)
    for r in use_rows:
        if str(r.get("item deleted") or "").strip():
            continue
        if str(r.get("phantom item") or "").strip():
            continue
        m = norm_material(r.get("material"))
        d = to_date(r.get("requirements date"))
        q = to_num(r.get("difference quantity"))
        if not m or d is None or is_excluded(m, r.get("material description")):
            continue
        remember(m, r.get("material description"))
        if d >= today:
            total_cons[m] += q
            if d < win_end:
                window_cons[m] += q
            if d < tr_end:
                cons_3d[m] += q
    for r in ord_rows:
        remember(norm_material(r.get("material number")), r.get("material description"))

    # ---- Per-material metrics ----------------------------------------------
    if ACTIVE_MATERIALS == "dc":
        active = {m for m, q in dc_side.items() if q > 0}
    elif ACTIVE_MATERIALS == "any":
        active = {m for m in set(dc_side) | set(ext_all) if dc_side.get(m, 0) + ext_all.get(m, 0) > 0}
    else:
        active = {m for m in set(dc_side) | set(ext_all) | set(total_cons)
                  if dc_side.get(m, 0) + ext_all.get(m, 0) > 0 or total_cons.get(m, 0) > 0}

    mats = {}
    for m in active:
        total = dc_side.get(m, 0.0) + ext_all.get(m, 0.0)
        weekly = window_cons.get(m, 0.0) / DEMAND_WEEKS
        if weekly > 0:
            weeks = total / weekly
            # Health score: 100 at the target cover and lower on both sides of
            # it (shortage and excess).  This keeps the availability table
            # meaningful instead of giving every overstocked item 100.
            score = round(100.0 * min(weeks, TARGET_WEEKS) / max(weeks, TARGET_WEEKS), 1)
        else:
            weeks = None            # display as "No demand"
            score = None            # not part of the availability average
        if weeks is None:
            status = "overstocked" if total > 0 else "critical"
        elif weeks < CRITICAL_BELOW:
            status = "critical"
        elif weeks < LOW_BELOW:
            status = "low"
        elif weeks <= OVERSTOCK_ABOVE:
            status = "adequate"
        else:
            status = "overstocked"
        dw, dq = idle_weight.get(m, (0.0, 0.0))
        mats[m] = dict(code=m, desc=desc.get(m, ""), total=total, dc=dc_side.get(m, 0.0),
                       ext=ext_stock.get(m, 0.0), weeks=weeks, score=score, status=status,
                       cons=total_cons.get(m, 0.0), c3=cons_3d.get(m, 0.0),
                       weekly=weekly,
                       dc_weeks=(dc_side.get(m, 0.0) / weekly) if weekly > 0 else None,
                       idle=(dw / dq) if dq else None,
                       uom=ext_qty_unit.get(m, "KG"),
                       fefo=sorted(fefo_by_mat.get(m, []),
                                   key=lambda b: (b["exp"] is None, b["exp"] or "9999",
                                                  b["gr"] or "9999", b["batch"])))

    L = list(mats.values())
    if not L:
        sys.exit("No active material found (no stock in the DC warehouse). Check the workbook and the "
                 f"ACTIVE_MATERIALS / DC_QTY_COLUMN settings. Nothing was written.")
    total_stock = sum(x["total"] for x in L)

    # Transfer universe: the active DC materials plus any material that has
    # 3-day demand or a named safety target.  This lets the transfer plan cover
    # a material that is not in DC yet (e.g. 15669) without changing the
    # dashboard's active-material KPI.
    transfer_universe = {x["code"]: dict(x) for x in L}

    def add_transfer_material(m):
        if m in transfer_universe:
            return
        total = dc_side.get(m, 0.0) + ext_all.get(m, 0.0)
        weekly = window_cons.get(m, 0.0) / DEMAND_WEEKS
        weeks = (total / weekly) if weekly > 0 else None
        if weeks is None:
            status = "overstocked" if total > 0 else "critical"
        elif weeks < CRITICAL_BELOW:
            status = "critical"
        elif weeks < LOW_BELOW:
            status = "low"
        elif weeks <= OVERSTOCK_ABOVE:
            status = "adequate"
        else:
            status = "overstocked"
        transfer_universe[m] = dict(
            code=m, desc=desc.get(m, ""), total=total, dc=dc_side.get(m, 0.0),
            ext=ext_stock.get(m, 0.0), weeks=weeks, score=None, status=status,
            cons=total_cons.get(m, 0.0), c3=cons_3d.get(m, 0.0),
            weekly=weekly,
            dc_weeks=(dc_side.get(m, 0.0) / weekly) if weekly > 0 else None,
            idle=None, uom=ext_qty_unit.get(m, "KG"), fefo=[], price=0.0, value=0.0)

    for m in set(cons_3d) | set(SAFETY_STOCK_PALLETS):
        add_transfer_material(m)
    transfer_universe = list(transfer_universe.values())

    # Portfolio price per kg: total value / total kg from the KG lines of the
    # externals sheet.  Used only for materials that have no price of their own.
    priced_value = sum(kg_val.values())
    priced_qty = sum(kg_qty.values())
    price_pool = priced_value / priced_qty if priced_qty > 0 else 0.0

    # Value per material: its own price (value / quantity in the externals sheet), else the pool price
    priced_kg = 0.0
    for x in L:
        qty, val = price_qty.get(x["code"], 0.0), price_val.get(x["code"], 0.0)
        own = val / qty if qty > 0 and val > 0 else None
        x["price"] = own if own is not None else price_pool
        x["value"] = x["total"] * x["price"]
        if own is not None:
            priced_kg += x["total"]
    total_value = sum(x["value"] for x in L)
    blended = total_value / total_stock if total_stock else 0.0

    by_status = {k: sum(1 for x in L if x["status"] == k) for k in ("adequate", "overstocked", "low", "critical")}
    status_value = {k: dict(kg=sum(x["total"] for x in L if x["status"] == k),
                            value=sum(x["value"] for x in L if x["status"] == k)) for k in by_status}

    critical_all = sorted([x for x in L if x["status"] == "critical"], key=lambda x: (math_inf if x["weeks"] is None else x["weeks"], -x["total"]))
    critical = critical_all[:TOP_CRITICAL]
    overstock_all = [x for x in L if x["status"] == "overstocked"]
    overstock = sorted(overstock_all,
                       key=lambda x: (x["weeks"] is None, -(x["weeks"] or 0.0), -x["total"]))[:TOP_OVERSTOCK]
    avail = sorted([x for x in L if x["score"] is not None],
                   key=lambda x: (-x["score"], -x["total"]))[:TOP_AVAIL]
    demand = sorted([x for x in L if x["cons"] > 0], key=lambda x: -x["cons"])
    for m, q in total_cons.items():                      # most-consumed list includes materials without stock
        if m not in mats and q > 0:
            demand.append(dict(code=m, desc=desc.get(m, ""), cons=q))
    demand = sorted(demand, key=lambda x: -x["cons"])[:TOP_DEMAND]

    transfers = []
    for x in L:
        need = max(0.0, x["c3"] - x["dc"])
        if need > 0 and x["ext"] > 0:                    # must exist outside DC to be moved
            picks, ideal, ideal_qty = fefo_picks(ext_batches, batch_expiry, x["code"], need)
            transfers.append(dict(x, need=need, picks=picks, ideal=ideal, ideal_qty=ideal_qty))
    transfers.sort(key=lambda x: -x["need"])
    transfer_total = sum(x["need"] for x in transfers)

    slow_all = [x for x in L if x["idle"] is not None and x["idle"] > SLOW_MOVER_DAYS and x["dc"] > 0]
    slow = sorted(slow_all, key=lambda x: -x["idle"])[:TOP_SLOW]
    top_stock = sorted(L, key=lambda x: -x["total"])[:TOP_CHART]

    scored = [x["score"] for x in L if x["score"] is not None]
    kpi = dict(
        materials=len(L),
        total_stock=total_stock,
        value=total_value,
        rkg=blended,
        availability=round(sum(scored) / len(scored)) if scored else 0,
        critical=by_status["critical"],
        overstocked=by_status["overstocked"],
        transfer=transfer_total,
    )
    capital = dict(
        overstock=dict(n=len(overstock_all), kg=sum(x["total"] for x in overstock_all), value=sum(x["value"] for x in overstock_all)),
        idle=dict(n=len(slow_all), kg=sum(x["dc"] for x in slow_all), value=sum(x["dc"] * x["price"] for x in slow_all)),
        priced_pct=(priced_kg / total_stock) if total_stock else 0.0)
    non_kg = sorted(x["code"] for x in L if x["uom"] not in ("KG", ""))
    return dict(kpi=kpi, non_kg=non_kg, by_status=by_status, status_value=status_value, capital=capital,
                critical=critical, overstock=overstock, avail=avail, demand=demand,
                transfers=transfers[:TOP_TRANSFER], transfer_total=transfer_total, slow=slow,
                top_stock=top_stock, top_cons=demand[:TOP_CHART], desc=desc, today=today,
                transfer_universe=transfer_universe, ext_batches=ext_batches,
                batch_expiry=batch_expiry)


# =============================================================================
# 3b. LOAD  (warehouse positions: "MP Stock Saturation" workbook)
# =============================================================================
def sat_pct(rows):
    """Saturation % = occupied / total over [code, total, occupied] rows (None when there are no positions)."""
    tot = sum(t for _, t, _ in rows)
    return round((1 - sum(t - o for _, t, o in rows) / tot) * 100, 1) if tot > 0 else None


def drop_excluded_positions(slots, desc=None):
    """Remove excluded materials from the dedicated positions and recompute the dedicated saturation.
    Used for the HTML-kept positions too, so an old dashboard cannot bring an excluded material back."""
    desc = desc or {}
    names = slots.get("names") or {}

    def excluded(code):
        return is_excluded(code, desc.get(code) or names.get(code, ""))

    removed = {r[0] for r in slots.get("ded", []) if excluded(r[0])}
    removed |= {k for k in names if excluded(k)}
    if not removed:
        return slots, 0
    ded = [r for r in slots.get("ded", []) if r[0] not in removed]
    out = dict(slots, ded=ded, names={k: v for k, v in names.items() if k not in removed})
    out["rep"] = dict(slots.get("rep") or {}, ded=sat_pct(ded))
    return out, len(removed)


def load_positions(path, desc):
    """Dedicated positions, shared positions, blocked row, reported saturation and the daily trend.

    Sheet layout (see sample_data/Position Saturation.xlsx):
      A: PROD/END   B: Total   C: Occupied   D: Free   E: Turnover days      G: Date  H: Shared saturation  I: Dedicated saturation
      rows: shared areas (POOL_LABELS) | one row per dedicated material | Blocked |
            'Shared saturation' / 'Dedicated saturation' summary rows (value in column B)
    """
    wb, tmp_dir = open_workbook_copy(path)
    try:
        ws = find_sheet(wb, SAT_SHEET or wb.sheetnames[0])
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)

    if not rows:
        raise RuntimeError("positions sheet is empty")

    head = [plain(c) for c in rows[0]]
    gi = next((i for i, h in enumerate(head) if h == "date"), None)
    hi = next((i for i, h in enumerate(head) if "saturation" in h and "shared" in h), None)
    di = next((i for i, h in enumerate(head) if "saturation" in h and "dedicated" in h), None)

    ded, pool, blocked, rep, dropped = [], [], None, {}, []
    for r in rows[1:]:
        label = plain(r[0] if r else None)
        if not label:
            continue
        total, occ = r[1] if len(r) > 1 else None, r[2] if len(r) > 2 else None
        if "saturation" in label:
            v = to_pct(total)
            if "shared" in label:
                rep["shared"] = v
            elif "dedicated" in label:
                rep["ded"] = v
            continue
        if label == "date" or total is None or occ is None:
            continue
        t, o = int(round(to_num(total))), int(round(to_num(occ)))
        if label == BLOCKED_LABEL:
            blocked = [t, o]
        elif label in POOL_LABELS:
            pool.append([str(r[0]).strip(), t, o])
        elif label in IGNORE_POSITION_LABELS:
            continue
        else:
            code = norm_material(r[0])
            if is_excluded(code, desc.get(code, "")):
                dropped.append(code)
                continue
            ded.append([code, t, o])

    if dropped:        # the sheet's own saturation cells still count them -> recompute from the kept rows
        rep.pop("ded", None)
        print(f"  Positions: {', '.join(dropped)} removed from the position map; dedicated saturation recomputed.")
        print("  ! The saturation history (trend chart) is read from the sheet and still includes it "
              "- edit the sheet history if the trend must be restated.")
    trend = []
    if None not in (gi, hi, di):
        for r in rows[1:]:
            if len(r) <= max(gi, hi, di):
                continue
            d, sv, dv = to_date(r[gi]), to_pct(r[hi]), to_pct(r[di])
            if d and sv is not None and dv is not None:
                trend.append((d, sv, dv))
        trend.sort(key=lambda x: x[0])
        trend = [[fmt_date_short(d), sv, dv] for d, sv, dv in trend[-TREND_POINTS:]]

    # fallback saturation: 1 - free / total  (same formula as the sheet)
    if "ded" not in rep and sat_pct(ded) is not None:
        rep["ded"] = sat_pct(ded)
    if "shared" not in rep and sat_pct(pool) is not None:
        rep["shared"] = sat_pct(pool)
    names = {c: desc[c][:NAME_MAX].strip().title() if c in desc else "" for c, _, _ in ded}
    return dict(ded=ded, pool=pool, blocked=blocked, trend=trend, rep=rep, names={k: v for k, v in names.items() if v})


# =============================================================================
# 4. HTML REWRITE  (anchored on stable text; safe to run again and again)
# =============================================================================
DATA_RE = re.compile(r'(<script id="dash-data" type="application/json">)(.*?)(</script>)', re.S)


def rows_html(rows):
    return "<tbody>" + "".join(
        "<tr>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>" for r in rows) + "</tbody>"


def replace_table(h, title, rows):
    pat = re.compile(r'(<div class="section-title">' + re.escape(title) + r'</div>.*?<table class="dt">.*?</thead>)<tbody>.*?</tbody>',
                     re.S)
    if not pat.search(h):
        raise RuntimeError(f"Table '{title}' not found in HTML")
    return pat.sub(lambda m: m.group(1) + rows_html(rows), h, count=1)


def replace_kpi(h, label, value):
    pat = re.compile(r'(<div class="kpi-val">)[^<]*(</div>\s*<div class="kpi-lbl">' + re.escape(label) + r'</div>)')
    if not pat.search(h):
        raise RuntimeError(f"KPI '{label}' not found in HTML")
    return pat.sub(lambda m: m.group(1) + value + m.group(2), h, count=1)


def build_transfer_plan(R, slots):
    """Build the demand + safety + strategic transfer plan and its daily waves.

    Priority:
      1. 3-day demand shortage (zero-rupture cover).
      2. Named safety stock targets (3-day demand + safety pallets).
      3. Dedicated strategic fill: use every remaining free trip position for
         the highest-consumption dedicated materials, starting with the
         lowest DC cover.
    """
    ded_by_code = {r[0]: r for r in (slots.get("ded") or []) if r}

    def pallet_kg(x):
        d = ded_by_code.get(x["code"])
        occupied = d[2] if d and len(d) > 2 else 0
        kg = (x["dc"] / occupied) if occupied > 0 and x["dc"] > 0 else DEFAULT_KG_PER_POSITION
        return max(float(kg), 1.0)

    def pack_picks(picks):
        return [dict(batch=p["batch"], sloc=p["sloc"], qty=round(p["qty"]),
                     available=round(p["available"]), exp=p["exp"],
                     src_name=TRANSFER_SOURCE_NAMES.get(p["sloc"], p["sloc"]))
                for p in picks]

    def type_label(demand, safety, strategic):
        parts = []
        if demand > 0:
            parts.append("Demand")
        if safety > 0:
            parts.append("Safety")
        if strategic > 0:
            parts.append("Strategic")
        return " + ".join(parts) or "Plan"

    universe = R.get("transfer_universe") or R.get("all") or R["transfers"]
    base = {}
    safety_rows = []

    # ---- 1/2. Demand shortage + safety target ------------------------------
    for x in universe:
        code = x["code"]
        kp = pallet_kg(x)
        demand = max(0.0, x.get("c3", 0.0) - x.get("dc", 0.0))
        target_pallets = SAFETY_STOCK_PALLETS.get(code, 0)
        target_kg = target_pallets * kp
        safety = 0.0
        if target_pallets:
            # Keep the safety buffer after the 3-day consumption.
            total_gap = max(0.0, x.get("c3", 0.0) + target_kg - x.get("dc", 0.0))
            safety = max(0.0, total_gap - demand)
        qty = demand + safety
        ext = max(0.0, x.get("ext", 0.0))
        picked_est = min(qty, ext)
        pallets = max(1, math.ceil(picked_est / kp)) if picked_est > 0 else 0
        short = max(0.0, qty - picked_est)
        if qty > 0 or target_pallets:
            base[code] = dict(x=x, kp=kp, demand=demand, safety=safety, strategic=0.0,
                              qty=qty, ext=ext, picked_est=picked_est,
                              pallets=pallets, short=short)
        if target_pallets:
            safety_rows.append(dict(
                c=code, d=x.get("desc", ""), kp=round(kp, 1),
                target_pallets=target_pallets, target_kg=round(target_kg),
                dc=round(x.get("dc", 0.0)), demand=round(x.get("c3", 0.0)),
                safety_need=round(safety),
                gap_pallets=(math.ceil(safety / kp) if safety > 0 else 0),
                ext=round(ext),
                status=("OK" if safety <= 0 else "Bring" if ext > 0 else "No source")))

    # ---- 3. Strategic fill of the remaining free trip positions -----------
    free_capacity = TRANSFER_FREE_TRIPS * TRANSFER_NORMAL_SPOTS
    base_pallets = sum(b["pallets"] for b in base.values())
    strategic_target = max(0, free_capacity - base_pallets)
    planned = {code: b["pallets"] for code, b in base.items()}
    short_codes = {code for code, b in base.items() if b["short"] > 0}

    candidates = []
    for x in universe:
        code = x["code"]
        d = ded_by_code.get(code)
        if (not d or code in short_codes or x.get("ext", 0.0) <= 0
                or x.get("cons", 0.0) <= 0):
            continue
        total_pos, occupied = d[1], d[2]
        free_after = total_pos - occupied - planned.get(code, 0)
        if free_after <= 0:
            continue
        bw = x.get("dc_weeks")
        tier = 0 if (bw is not None and bw < TARGET_WEEKS) else 1
        candidates.append(dict(x=x, free=free_after, tier=tier,
                               cons=x.get("cons", 0.0), bw=bw))
    candidates.sort(key=lambda c: (c["tier"], -c["cons"],
                                  c["bw"] if c["bw"] is not None else 9999,
                                  c["x"]["code"]))

    remaining = strategic_target
    strategic_pallets_total = 0
    for c in candidates:
        if remaining <= 0:
            break
        x = c["x"]
        code = x["code"]
        kp = pallet_kg(x)
        b = base.get(code)
        ext_left = max(0.0, x.get("ext", 0.0) - (b["picked_est"] if b else 0.0))
        max_by_ext = int(ext_left // kp)
        take = min(c["free"], remaining, max_by_ext)
        if take <= 0:
            continue
        if b is None:
            b = dict(x=x, kp=kp, demand=0.0, safety=0.0, strategic=0.0,
                     qty=0.0, ext=max(0.0, x.get("ext", 0.0)),
                     picked_est=0.0, pallets=0, short=0.0)
            base[code] = b
        add_qty = take * kp
        b["strategic"] += add_qty
        b["qty"] += add_qty
        b["picked_est"] += add_qty
        b["pallets"] += take
        strategic_pallets_total += take
        remaining -= take

    # ---- Final FEFO allocation and audit rows -----------------------------
    transfers, risks = [], []
    for code, b in base.items():
        x = b["x"]
        kp = b["kp"]
        total_qty = b["demand"] + b["safety"] + b["strategic"]
        if total_qty <= 0:
            continue
        picks, ideal, ideal_qty = fefo_picks(
            R["ext_batches"], R["batch_expiry"], code, total_qty)
        picked = sum(p["qty"] for p in picks)
        pallets = b["pallets"] if picked > 0 else 0
        if picked > 0 and total_qty - picked > 0.5:
            pallets = max(1, math.ceil(picked / kp))
        short = max(0.0, total_qty - picked)
        item = dict(
            c=code, d=x.get("desc", ""),
            type_label=type_label(b["demand"], b["safety"], b["strategic"]),
            demand=round(b["demand"]), safety=round(b["safety"]),
            strategic=round(b["strategic"]), need=round(total_qty),
            picked=round(picked), short=round(short),
            kp=round(kp, 1), pallets=pallets, load=round(pallets * kp),
            picks=pack_picks(picks), ideal=ideal,
            ideal_name=TRANSFER_SOURCE_NAMES.get(ideal, ideal),
            ideal_qty=round(ideal_qty),
            c3=round(x.get("c3", 0.0)), dc=round(x.get("dc", 0.0)),
            ext=round(x.get("ext", 0.0)),
            other=round(max(0.0, x.get("total", 0.0) - x.get("dc", 0.0) - x.get("ext", 0.0))),
            cons=round(x.get("cons", 0.0)),
            weeks=x.get("weeks"), dc_weeks=x.get("dc_weeks"))
        if picked > 0:
            transfers.append(item)
        if short > 0.5 or (total_qty > 0 and picked <= 0):
            other = max(0.0, x.get("total", 0.0) - x.get("dc", 0.0) - x.get("ext", 0.0))
            if x.get("ext", 0.0) <= 0:
                reason = "Stock only outside the hubs" if other > 0 else "No hub stock"
            else:
                reason = "External stock short"
            risks.append(dict(item, reason=reason))

    # Low DC cover with no external source at all is a source risk too.
    risk_codes = {r["c"] for r in risks}
    for x in universe:
        code = x["code"]
        if code in risk_codes:
            continue
        bw, weekly = x.get("dc_weeks"), x.get("weekly") or 0.0
        if (bw is None or bw >= 1 or x.get("ext", 0.0) > 0 or weekly <= 0):
            continue
        kp = pallet_kg(x)
        risks.append(dict(
            c=code, d=x.get("desc", ""), type_label="Cover",
            demand=0, safety=0, strategic=0, need=0, picked=0, short=0,
            kp=round(kp, 1), pallets=0, load=0, picks=[], ideal=None,
            ideal_name=None, ideal_qty=0, c3=round(x.get("c3", 0.0)),
            dc=round(x.get("dc", 0.0)), ext=0, other=0,
            cons=round(x.get("cons", 0.0)), weeks=x.get("weeks"),
            dc_weeks=bw, reason="Low DC cover, no external stock"))

    transfers.sort(key=lambda t: (0 if t["demand"] > 0 else 1 if t["safety"] > 0 else 2,
                                  -t["pallets"], -t["need"]))
    risks.sort(key=lambda r: (-r["short"],
                              r["dc_weeks"] if r["dc_weeks"] is not None else 9999,
                              -r["need"]))
    safety_order = {c: i for i, c in enumerate(SAFETY_STOCK_PALLETS)}
    safety_rows.sort(key=lambda r: safety_order.get(r["c"], 99))

    # ---- Ready-to-go waves: 2 free trips x 24 spots, then paid extra 28 ----
    waves = []
    for i in range(TRANSFER_FREE_TRIPS):
        waves.append(dict(key=("morning" if i == 0 else "afternoon" if i == 1 else f"free{i + 1}"),
                          name=("Morning" if i == 0 else "Afternoon" if i == 1 else f"Free trip {i + 1}"),
                          paid=False, capacity=TRANSFER_NORMAL_SPOTS,
                          pallets=0, kg=0.0, items=[]))

    def put(w, t, take):
        srcs = []
        for p in t["picks"]:
            if p["sloc"] not in srcs:
                srcs.append(p["sloc"])
        src_label = " + ".join(f"{TRANSFER_SOURCE_NAMES.get(s, s)} ({s})" for s in srcs) if srcs else t["ideal_name"]
        w["items"].append(dict(c=t["c"], d=t["d"], src=t["ideal"], src_name=t["ideal_name"],
                               srcs=srcs, src_label=src_label, type_label=t["type_label"],
                               pallets=take, kg=round(take * t["kp"]), picks=t["picks"]))
        w["pallets"] += take
        w["kg"] = round(w["kg"] + take * t["kp"])

    extra_trips = 0
    for t in transfers:
        left = t["pallets"]
        for w in waves:
            if left <= 0:
                break
            room = w["capacity"] - w["pallets"]
            if room <= 0:
                continue
            take = min(left, room)
            put(w, t, take)
            left -= take
        while left > 0:
            extra_trips += 1
            w = dict(key=f"extra{extra_trips}",
                     name=("Extra (paid)" if extra_trips == 1 else f"Extra {extra_trips} (paid)"),
                     paid=True, capacity=TRANSFER_EXTRA_SPOTS, pallets=0, kg=0.0, items=[])
            take = min(left, TRANSFER_EXTRA_SPOTS)
            put(w, t, take)
            left -= take
            waves.append(w)

    # An item can cross a wave boundary; show it once per wave.
    for w in waves:
        merged, by_code = [], {}
        for it in w["items"]:
            if it["c"] in by_code:
                m = by_code[it["c"]]
                m["pallets"] += it["pallets"]
                m["kg"] = round(m["kg"] + it["kg"])
            else:
                m = dict(it)
                by_code[it["c"]] = m
                merged.append(m)
        w["items"] = merged

    plan = dict(
        total_pallets=sum(t["pallets"] for t in transfers),
        total_load=round(sum(t["pallets"] * t["kp"] for t in transfers)),
        demand_kg=round(sum(t["demand"] for t in transfers)),
        safety_kg=round(sum(t["safety"] for t in transfers)),
        strategic_kg=round(sum(t["strategic"] for t in transfers)),
        base_pallets=base_pallets, strategic_pallets=strategic_pallets_total,
        free_trips=TRANSFER_FREE_TRIPS, normal_capacity=TRANSFER_NORMAL_SPOTS,
        extra_capacity=TRANSFER_EXTRA_SPOTS,
        free_capacity=free_capacity,
        extra_trips=extra_trips, waves=waves,
        source_names=TRANSFER_SOURCE_NAMES,
        safety=safety_rows, risks=risks[:TOP_RISK], risk_count=len(risks))
    return dict(transfers=transfers, plan=plan)


def build_data(R, slots, today):
    """Everything the page's charts, position map and transfer plan read (one JSON block)."""
    k, desc = R["kpi"], R["desc"]
    r1 = lambda v: round(float(v), 1)

    P = build_transfer_plan(R, slots)
    transfers, plan = P["transfers"], P["plan"]

    return dict(
        asof=today.isoformat(), slow_days=SLOW_MOVER_DAYS, transfer_days=TRANSFER_DAYS,
        default_kpp=DEFAULT_KG_PER_POSITION,
        kpi=dict(materials=k["materials"], total_kg=round(k["total_stock"]), value=round(k["value"]), rkg=round(k["rkg"], 4),
                 availability=k["availability"], critical=k["critical"], overstocked=k["overstocked"], transfer=round(k["transfer"])),
        quality={s: R["by_status"][s] for s in ("critical", "low", "adequate", "overstocked")},
        status_value={s: dict(kg=round(v["kg"]), value=round(v["value"])) for s, v in R["status_value"].items()},
        capital=dict(overstock={a: round(b) if a != "n" else b for a, b in R["capital"]["overstock"].items()},
                     idle={a: round(b) if a != "n" else b for a, b in R["capital"]["idle"].items()},
                     priced_pct=round(R["capital"]["priced_pct"], 3)),
        top_stock=[dict(c=x["code"], kg=r1(x["total"]), v=round(x["value"])) for x in R["top_stock"]],
        top_cons=[dict(c=x["code"], kg=r1(x["cons"])) for x in R["top_cons"]],
        demand_rank={x["code"]: i + 1 for i, x in enumerate(R["demand"])},
        critical=[dict(c=x["code"], d=x["desc"], kg=round(x["total"])) for x in R["critical"][:5]],
        transfers=transfers[:TOP_TRANSFER],
        transfer_plan=plan,
        transfers_total=round(R["transfer_total"]),
        slots=slots)


def update_html(h, R, slots):
    k, d = R["kpi"], R["today"]

    # --- KPI cards
    h = replace_kpi(h, "Active Materials", fmt_int(k["materials"]))
    h = replace_kpi(h, "Total Stock (kg)", fmt_int(k["total_stock"]))
    h = replace_kpi(h, "Stock Value", "$" + fmt_int(k["value"]))
    h = replace_kpi(h, "Availability Score", f"{k['availability']}%")
    h = replace_kpi(h, "Critical Items", fmt_int(k["critical"]))
    h = replace_kpi(h, "Overstocked", fmt_int(k["overstocked"]))
    h = replace_kpi(h, "3d Transfer (kg)", fmt_int(k["transfer"]))
    h = replace_kpi(h, "Average value per kg", f"${k['rkg']:.2f}")
    rep = slots.get("rep") or {}
    if rep.get("ded") is not None:
        h = replace_kpi(h, "Dedicated position saturation", f"{rep['ded']:.1f}%")
    if rep.get("shared") is not None:
        h = replace_kpi(h, "Shared position saturation", f"{rep['shared']:.1f}%")

    # --- Hero floating cards
    h = re.sub(r'(<div class="float glass f1"><b>)[^<]*(</b>)', lambda m: m.group(1) + f"{k['availability']}%" + m.group(2), h)
    h = re.sub(r'(<div class="float glass f2"><b>)[^<]*(</b>)', lambda m: m.group(1) + fmt_int(k["critical"]) + m.group(2), h)

    # --- Ticker (4 repeats, same structure as before)
    items = [f"{fmt_int(k['materials'])} ACTIVE MATERIALS", f"{k['total_stock'] / 1e6:.2f}M KG IN STOCK",
             f"{fmt_int(k['critical'])} CRITICAL ITEMS", f"{fmt_int(k['transfer'])} KG TO TRANSFER"]
    block = "<span>STOCK INTELLIGENCE</span><i>&#9632;</i>" + "".join(f"<span>{t}</span><i>&#9632;</i>" for t in items)
    h, n = re.subn(r'(<div class="ticker"><div>).*?(</div></div>)', lambda m: m.group(1) + block * 4 + m.group(2), h, count=1, flags=re.S)
    if n == 0:
        raise RuntimeError("Ticker not found in HTML")

    # --- Dates (overview eyebrow, transfer eyebrow, footer)
    h = re.sub(r'(&middot; )\d{1,2} \w+ \d{4}', lambda m: m.group(1) + fmt_date_long(d), h)
    h = re.sub(r'(fictional data\s*•\s*)\d{1,2} \w+ \d{4}', lambda m: m.group(1) + fmt_date_long(d), h)

    # --- Tables (the transfer plan is drawn from the data block, not from a table)
    h = replace_table(h, "Critical Shortages",
                      [(x["code"], x["desc"], fmt_int(x["total"]), fmt_weeks(x["weeks"])) for x in R["critical"]])
    h = replace_table(h, "Overstocked Materials",
                      [(x["code"], x["desc"], fmt_int(x["total"]), fmt_weeks(x["weeks"])) for x in R["overstock"]])
    h = replace_table(h, "Highest Availability Materials",
                      [(x["code"], x["desc"], fmt_int(x["total"]), fmt_int(round(x["score"]))) for x in R["avail"]])
    h = replace_table(h, "Most Consumed Materials", [(x["code"], fmt_int(x["cons"])) for x in R["demand"]])
    h = replace_table(h, "Slow Moving Inventory",
                      [(x["code"], x["desc"], fmt_int(x["dc"]), fmt_int(x["idle"])) for x in R["slow"]])

    # --- Data block read by the charts, position map and transfer plan
    blob = json.dumps(build_data(R, slots, d), separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
    h, n = DATA_RE.subn(lambda m: m.group(1) + blob + m.group(3), h, count=1)
    if n == 0:
        raise RuntimeError('Data block <script id="dash-data"> not found in HTML')
    return h


# =============================================================================
# 5. MAIN
# =============================================================================
def check_exclusions(R, slots):
    """Hard stop if an excluded material is still present in anything that reaches the page."""
    desc = R["desc"]
    names = slots.get("names") or {}

    def leak(code, description=""):
        return is_excluded(code, description or desc.get(code, "") or names.get(code, ""))

    leaks = set()
    for key in ("critical", "overstock", "avail", "demand", "transfers", "slow", "top_stock", "top_cons"):
        leaks |= {x["code"] for x in R[key]
                  if leak(x["code"], x.get("desc") or x.get("d", ""))}
    leaks |= {r[0] for r in slots.get("ded", []) if leak(r[0])}
    leaks |= {k for k in names if leak(k)}
    if leaks:
        sys.exit(f"Excluded material(s) still present in the output: {sorted(leaks)}. Nothing was written.")


def check_output(new, R, old):
    """Re-read the data block of the finished HTML and compare it with the calculation."""
    m = DATA_RE.search(new)
    try:
        data = json.loads(m.group(2))
    except (AttributeError, ValueError) as e:
        sys.exit(f"The generated HTML has an unreadable data block ({e}). Nothing was written.")
    if data["kpi"]["materials"] != R["kpi"]["materials"] or not data["top_stock"]:
        sys.exit("The generated data block does not match the calculation. Nothing was written.")
    try:
        before = json.loads(DATA_RE.search(old).group(2))["kpi"]["materials"]
        now = R["kpi"]["materials"]
        if before and abs(now - before) / before > 0.3:
            print(f"  ! Active materials changed from {before} to {now} (more than 30%). Check the workbook before sharing.")
    except (AttributeError, ValueError, KeyError):
        pass


def print_summary(R, slots):
    k, bs, cp = R["kpi"], R["by_status"], R["capital"]
    print(f"\n  Data date ........ {R['today']}")
    print(f"  Active materials . {k['materials']}")
    print(f"  Total stock (kg) . {fmt_int(k['total_stock'])}")
    print(f"  Stock value $... {fmt_int(k['value'])}   (average ${k['rkg']:.2f}/kg)")
    print(f"  Availability ..... {k['availability']}%")
    print(f"  Saturation ....... critical {bs['critical']} | low {bs['low']} | adequate {bs['adequate']} | overstocked {bs['overstocked']}")
    print(f"  Capital in overstock ${fmt_int(cp['overstock']['value'])} ({cp['overstock']['n']} materials), "
          f"idle ${fmt_int(cp['idle']['value'])} ({cp['idle']['n']} materials), "
          f"{cp['priced_pct'] * 100:.0f}% of kg valued at own price")
    print(f"  3d transfer (kg) . {fmt_int(k['transfer'])}  ({len(R['transfers'])} demand lines)")
    P = build_transfer_plan(R, slots)
    tp = P["plan"]
    print(f"  Transfer plan .... {tp['total_pallets']} pallets / {fmt_int(tp['total_load'])} kg "
          f"(demand {fmt_int(tp['demand_kg'])}, safety {fmt_int(tp['safety_kg'])}, "
          f"strategic {fmt_int(tp['strategic_kg'])})")
    print(f"  Trips ............ {tp['free_trips']} free x {tp['normal_capacity']} = "
          f"{tp['free_capacity']} positions; paid extra {tp['extra_trips']} trip(s) x {tp['extra_capacity']}")
    if tp["risk_count"]:
        print(f"  ! Rupture risks .. {tp['risk_count']} material(s) without enough hub stock: "
              f"{', '.join(r['c'] for r in tp['risks'][:6])}")
    print(f"  Rows: critical {len(R['critical'])}, overstock {len(R['overstock'])}, availability {len(R['avail'])}, "
          f"demand {len(R['demand'])}, transfer {len(R['transfers'])}, slow {len(R['slow'])}")
    s = slots
    print(f"  Positions ........ dedicated {len(s['ded'])}, shared areas {len(s['pool'])}, trend readings {len(s['trend'])}, "
          f"reported saturation {s.get('rep')}")
    if R.get("non_kg"):
        print(f"  ! {len(R['non_kg'])} active material(s) are not in KG but are summed into 'kg': {', '.join(R['non_kg'][:8])}")
    noname = sum(1 for x in R["avail"] + R["critical"] if not x["desc"])
    if noname:
        print(f"  ! {noname} listed material(s) have no description in any sheet")
    missing = [x["c"] for x in P["transfers"] if not any(c == x["c"] for c, _, _ in s["ded"])]
    if missing:
        print(f"  ! {len(missing)} transfer material(s) have no dedicated position (all go to shared): {', '.join(missing[:8])}")


def main():
    ap = argparse.ArgumentParser(description="Update the Stock Intelligence HTML dashboard from the ERP workbook")
    ap.add_argument("--xlsx", default=XLSX_PATH, help="Stock Control workbook")
    ap.add_argument("--sat-xlsx", default=SAT_XLSX_PATH, help="Position Saturation workbook (warehouse positions)")
    ap.add_argument("--html", default=HTML_PATH)
    ap.add_argument("--as-of", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    today = dt.date.fromisoformat(a.as_of) if a.as_of else dt.date.today()
    for p, n in ((a.xlsx, "Excel"), (a.html, "HTML")):
        if not os.path.isfile(p):
            sys.exit(f"{n} file not found: {p}\nEdit XLSX_PATH / HTML_PATH at the top of this script or use --xlsx / --html.")

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Reading {a.xlsx}")
    R = load_and_calculate(a.xlsx, today)

    with open(a.html, encoding="utf-8") as f:
        old = f.read()

    # Warehouse positions: from the saturation workbook when available, else keep what is in the HTML
    slots = None
    if os.path.isfile(a.sat_xlsx):
        print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Reading {a.sat_xlsx}")
        try:
            slots = load_positions(a.sat_xlsx, R["desc"])
        except Exception as e:                       # keep the dashboard usable if the layout changed
            print(f"  ! Could not read positions ({e}). Keeping the positions already in the HTML.")
    else:
        print(f"  ! Saturation workbook not found ({a.sat_xlsx}). Keeping the positions already in the HTML.")
    if slots is None:
        m = DATA_RE.search(old)
        if not m:
            sys.exit('Data block <script id="dash-data"> not found in HTML. Use the V3 dashboard file.')
        slots = json.loads(m.group(2))["slots"]

    slots, n_drop = drop_excluded_positions(slots, R["desc"])
    if n_drop:
        print(f"  Positions: {n_drop} excluded material(s) removed from the kept position data.")
    check_exclusions(R, slots)
    print_summary(R, slots)
    new = update_html(old, R, slots)
    check_output(new, R, old)

    if a.dry_run:
        print("\nDry run - nothing written.")
        return
    if BACKUP:
        shutil.copy2(a.html, a.html + ".bak")
    tmp = a.html + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(new)
    os.replace(tmp, a.html)          # atomic: the dashboard is never half-written
    print(f"\nDashboard updated -> {a.html}")


if __name__ == "__main__":
    main()
