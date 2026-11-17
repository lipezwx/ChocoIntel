"""
Generate FICTIONAL sample workbooks for the Stock Intelligence dashboard.

Nothing here comes from a real company: every material, quantity, price, batch,
warehouse position and saturation reading is produced from a fixed random seed.

Output (written to ./sample_data):
  - Stock Control.xlsx      sheets: In Stock - DC | In Stock - Externals | Usage Material
  - Position Saturation.xlsx  warehouse position map + daily saturation history

Usage:
  python generate_sample_data.py
  python update_dashboard.py --as-of 2026-11-17
"""
import datetime as dt
import math
import os
import random

import openpyxl

SEED = 2026
DEMO_DATE = dt.date(2026, 11, 17)      # the "today" of the demo data set
OUT_DIR = "sample_data"
DC_SLOC = "1000"                       # main distribution centre
HUBS = {"3001": "North Hub", "3002": "East Hub"}
OTHER_SLOC = "3009"                    # location that is NOT a transfer source

rnd = random.Random(SEED)

# ----------------------------------------------------------------------------
# 1. Fictional material catalogue
# ----------------------------------------------------------------------------
RAW = [
    ("Refined Sugar Fine", 0.95), ("Refined Sugar Coarse", 0.90), ("Brown Cane Sugar", 1.40),
    ("Whole Milk Powder", 4.60), ("Skim Milk Powder", 3.40), ("Sweet Whey Powder", 1.80),
    ("Cocoa Butter Natural", 9.50), ("Cocoa Butter Deodorized", 9.90), ("Cocoa Powder Dark", 5.20),
    ("Cocoa Powder Red", 5.60), ("Cocoa Liquor Block", 6.10), ("Soy Lecithin Fluid", 3.10),
    ("Sunflower Lecithin", 3.90), ("Vanillin Powder", 24.0), ("Polyglycerol Emulsifier", 7.20),
    ("Hazelnut Paste", 12.5), ("Almond Paste", 11.8), ("Sea Salt Fine", 0.40),
    ("Palm Oil Fractionated", 1.70), ("Shea Stearin", 3.60), ("Glucose Syrup 42DE", 0.85),
    ("Inulin Powder", 3.30), ("Stevia Extract", 38.0), ("Caramel Powder", 4.90),
    ("Malt Extract Dry", 2.70), ("Coconut Milk Powder", 5.80), ("Oat Flour Toasted", 1.60),
    ("Rice Crisps", 2.40), ("Wafer Crumbs", 1.90), ("Strawberry Pieces Dried", 18.0),
    ("Raspberry Pieces Dried", 21.0), ("Orange Peel Candied", 6.40), ("Flavour Vanilla Natural", 30.0),
    ("Flavour Caramel", 14.0), ("Colour Beetroot Red", 16.0), ("Citric Acid Anhydrous", 1.50),
    ("Maltodextrin", 1.10), ("Cocoa Nibs Roasted", 7.30), ("Pistachio Paste", 26.0),
    ("Peanut Paste", 4.20), ("Honey Powder", 5.10), ("Lactose Powder", 1.30),
    ("Butter Oil Anhydrous", 8.40), ("Cream Powder", 6.20), ("Yoghurt Powder", 7.00),
    ("Matcha Powder", 29.0), ("Espresso Powder", 13.0), ("Ginger Pieces Candied", 5.60),
    ("Cinnamon Ground", 6.90), ("Cardamom Ground", 34.0), ("Chilli Flakes", 8.10),
    ("Sesame Seeds Toasted", 3.80), ("Sunflower Seeds", 2.20), ("Pumpkin Seeds", 4.10),
    ("Quinoa Puffs", 5.30), ("Freeze-Dried Mango", 24.0), ("Baobab Powder", 17.0),
]
FILM_TYPES = ["Barrier", "Mono PE", "Laminate", "Shrink"]
FILM_FLAVOURS = ["Dark", "Milk", "White"]
FILM_FORMATS = ["Block 2.1 kg", "Chips 2.5 kg", "Bar 1.0 kg"]
CARTONS = ["RSC 15 x 11 x 7", "RSC 20 x 14 x 9", "RSC 30 x 20 x 12", "Display 12-pack",
           "Display 24-pack", "Tray 6-pack", "Shipper 40 x 30 x 25", "Mailer 25 x 18 x 8"]
MISC = ["Label Roll Thermal 100x150", "Tote Bag Liner", "Stretch Wrap Clear", "Corner Board 50 mm",
        "Pallet Cover Top", "Strapping Band PP", "Silica Gel Sachet", "Tape Carton Clear"]


def build_catalogue():
    items = []
    code = 1100
    for name, price in RAW:
        code += rnd.randint(7, 60)
        items.append(dict(code=str(code), desc=name, price=price, kind="raw"))
    code = 70010100
    for t in FILM_TYPES:
        for f in FILM_FLAVOURS:
            for fmt in FILM_FORMATS:
                if rnd.random() < 0.95:
                    code += rnd.randint(3, 90)
                    items.append(dict(code=str(code), desc=f"Film {t} {f} {fmt}",
                                      price=round(rnd.uniform(2.1, 4.6), 2), kind="film"))
    code = 80010000
    for c in CARTONS:
        code += rnd.randint(11, 120)
        items.append(dict(code=str(code), desc=f"Carton {c}", price=round(rnd.uniform(0.9, 1.9), 2), kind="pack"))
    code = 90010000
    for c in MISC:
        code += rnd.randint(11, 120)
        items.append(dict(code=str(code), desc=c, price=round(rnd.uniform(0.6, 3.2), 2), kind="pack"))
    return items


MATS = build_catalogue()
rnd.shuffle(MATS)

# Fixed demo roles ------------------------------------------------------------
SAFETY_CODES = [m["code"] for m in MATS if m["kind"] == "raw"][:8]      # named safety-stock materials
NO_DC_CODE = SAFETY_CODES[1]                                            # needs transfer, nothing in DC yet
OTHER_ONLY_CODE = [m["code"] for m in MATS if m["kind"] == "raw"][9]    # stock only at the "other" location
NO_SOURCE_CODE = [m["code"] for m in MATS if m["kind"] == "film"][0]    # low cover, no hub stock

# class mix: critical / low / adequate / overstocked / no demand
n = len(MATS)
plan = (["critical"] * 3 + ["low"] * 14 + ["adequate"] * 36 + ["nodemand"] * 5)
plan += ["over"] * (n - len(plan))
rnd.shuffle(plan)

# ----------------------------------------------------------------------------
# 2. Per-material numbers
# ----------------------------------------------------------------------------
for m, cls in zip(MATS, plan):
    code = m["code"]
    if code == NO_SOURCE_CODE:
        cls = "low"
    m["cls"] = cls
    base = rnd.lognormvariate(7.15, 1.05)                   # weekly demand, kg (median ~1.3 t)
    if m["kind"] == "raw":
        base *= 1.3
    m["weekly"] = 0 if cls == "nodemand" else max(120, round(base / 10) * 10)
    weeks = {"critical": rnd.uniform(0.12, 0.85), "low": rnd.uniform(1.2, 3.8),
             "adequate": rnd.uniform(4.4, 11.5), "over": min(160, 13 + rnd.expovariate(1 / 22)),
             "nodemand": 0}[cls]
    total = weeks * m["weekly"] if cls != "nodemand" else rnd.uniform(800, 9000)
    if code in SAFETY_CODES and cls in ("over", "nodemand"):
        m["cls"] = cls = "adequate"
        m["weekly"] = m["weekly"] or 2400
        total = rnd.uniform(5, 9) * m["weekly"]
    m["total"] = max(300.0, total)
    m["idle"] = rnd.randint(75, 240) if rnd.random() < 0.34 else rnd.randint(2, 55)
    m["kp"] = rnd.choice([800, 1000, 1000, 1200, 1250, 1300, 1500])   # kg per pallet

# Highlight a few heavy hitters so the charts have a clear Pareto shape
for m in sorted(MATS, key=lambda x: -x["weekly"])[:4]:
    m["weekly"] = int(m["weekly"] * 1.8)
    m["total"] *= 1.3


def split_stock(m):
    """-> (dc_kg, {sloc: kg})"""
    t = m["total"]
    code = m["code"]
    if code == NO_DC_CODE:
        return 0.0, {"3001": t * 0.6, "3002": t * 0.4}
    if code == OTHER_ONLY_CODE:
        return t * 0.25, {OTHER_SLOC: t * 0.75}
    if code == NO_SOURCE_CODE:
        return t, {}
    f = rnd.uniform(0.28, 0.92) if m["cls"] != "critical" else rnd.uniform(0.5, 1.0)
    dc = t * f
    rest = t - dc
    if rest < 1:
        return dc, {}
    hubs = {}
    a = rnd.random()
    hubs["3001"] = rest * (a if rnd.random() < 0.7 else 1)
    hubs["3002"] = rest - hubs["3001"]
    if rnd.random() < 0.06:
        hubs[OTHER_SLOC] = hubs.pop("3002") * 0.5
    return dc, {k: v for k, v in hubs.items() if v > 20}


def r5(v):
    return round(v / 5) * 5 or 5


# ----------------------------------------------------------------------------
# 3. Workbook 1: Stock Control
# ----------------------------------------------------------------------------
def month_add(d, months):
    y, mo = divmod(d.month - 1 + months, 12)
    return dt.date(d.year + y, mo + 1, min(d.day, 28))


def dt_(d):
    return dt.datetime(d.year, d.month, d.day)


wb = openpyxl.Workbook()

# --- In Stock - DC (positional layout, header in row 1) -----------------------
ws = wb.active
ws.title = "In Stock - DC"
ws.append(["Material", "Material Description", "Batch", "Special Stock", "Storage Location", "Storage Type",
           "Storage Bin", "GR Date", "Stock Category", "Available stock", "Total Stock", "Number of quants",
           "Last movement", "Time", "Base unit", "Expiry date"])
dc_stock, hub_rows = {}, {}
for m in MATS:
    dc, hubs = split_stock(m)
    dc_stock[m["code"]], hub_rows[m["code"]] = dc, hubs
    if dc <= 0:
        continue
    nb = rnd.randint(1, 4)
    cuts = sorted(rnd.random() for _ in range(nb - 1))
    parts = [b - a for a, b in zip([0] + cuts, cuts + [1])]
    for i, p in enumerate(parts):
        q = round(dc * p, 1)
        if q <= 0:
            continue
        last = DEMO_DATE - dt.timedelta(days=max(1, int(m["idle"] * rnd.uniform(0.85, 1.15))))
        gr = last - dt.timedelta(days=rnd.randint(5, 60))
        ws.append([m["code"], m["desc"], f"{rnd.randint(2400000, 2699999)}", None, DC_SLOC, "100",
                   f"A{rnd.randint(1, 40):02d}-{rnd.randint(1, 20):02d}", dt_(gr), None, q, q,
                   max(1, round(q / m["kp"])), dt_(last), "08:00:00", "KG",
                   dt_(month_add(DEMO_DATE, rnd.randint(5, 24)))])
# rows that must be ignored by the dashboard (finished goods, melted, excluded codes)
for code, desc, q in [("FG-BAR-120", "Finished Bar Dark 120 g", 18000), ("FG-EZ-210", "Finished Chips Milk", 9500),
                      ("MLT100", "Chocolate Dark Melted Tank", 42000), ("1999", "Special Sweetener Bulk", 61000),
                      ("90000001", "Pallets Plastic 1200x800", 700)]:
    ws.append([code, desc, "2500001", None, DC_SLOC, "100", "Z01-01", dt_(DEMO_DATE - dt.timedelta(days=9)),
               None, q, q, 3, dt_(DEMO_DATE - dt.timedelta(days=3)), "08:00:00", "KG", None])

# --- In Stock - Externals ----------------------------------------------------
we = wb.create_sheet("In Stock - Externals")
we.append(["Material", "Material Description", "Storage Location", "Batch", "Base Unit of Measure",
           "Unrestricted", "Transit and Transfer", "In Quality Insp.", "Restricted-Use Stock", "Blocked",
           "Returns", "Value Unrestricted", "Val. in Trans./Tfr", "Value in QualInsp.", "Value Restricted",
           "Value BlockedStock"])
for m in MATS:
    price = m["price"]
    # the DC stock also shows up (location 1000) in this export, as in a real stock list
    rows = [(DC_SLOC, round(dc_stock[m["code"]], 1))] if dc_stock[m["code"]] > 0 else []
    for sloc, q in hub_rows[m["code"]].items():
        nb = rnd.randint(1, 3)
        cuts = sorted(rnd.random() for _ in range(nb - 1))
        for a, b in zip([0] + cuts, cuts + [1]):
            rows.append((sloc, round(q * (b - a), 1)))
    for sloc, q in rows:
        if q <= 0:
            continue
        qi = round(q * 0.04, 1) if rnd.random() < 0.12 else 0.0
        qu = round(q - qi, 1)
        we.append([m["code"], m["desc"], sloc, str(rnd.randint(1700000, 1799999)) if sloc != DC_SLOC else "",
                   "KG", qu, 0, qi, 0, 0, 0, round(qu * price, 2), 0, round(qi * price, 2), 0, 0])

# --- Usage Material (demand) -------------------------------------------------
wu = wb.create_sheet("Usage Material")
wu.append(["Material", "Material Description", "Requirements Date", "Difference Quantity", "Item Deleted",
           "Phantom Item"])


def workdays(start, n):
    d, out = start, []
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


for m in MATS:
    w = m["weekly"]
    if w <= 0:
        continue
    days = workdays(DEMO_DATE, 20)                      # 4 weeks of working days
    weights = [rnd.uniform(0.4, 1.6) for _ in days]
    k = (4 * w) / sum(weights)
    for d, wt in zip(days, weights):
        wu.append([m["code"], m["desc"], dt_(d), round(wt * k, 1), None, None])
    for wk in range(4, 12):                             # weeks 5..12, one line per week
        d = DEMO_DATE + dt.timedelta(days=7 * wk + rnd.randint(0, 4))
        wu.append([m["code"], m["desc"], dt_(d), round(w * rnd.uniform(0.75, 1.25), 1), None, None])
# a few rows that must be ignored (deleted / phantom)
wu.append([MATS[0]["code"], MATS[0]["desc"], dt_(DEMO_DATE), 99999, "X", None])
wu.append([MATS[1]["code"], MATS[1]["desc"], dt_(DEMO_DATE), 99999, None, "X"])

# Demand shock in the next 3 days for the transfer plan
extra = {}
for m in MATS:
    if m["code"] in SAFETY_CODES[:5] or m["code"] in (OTHER_ONLY_CODE, NO_SOURCE_CODE):
        qty = round(max(900, m["weekly"] * rnd.uniform(0.9, 1.6)), -1)
        d = DEMO_DATE + dt.timedelta(days=rnd.randint(0, 2))
        wu.append([m["code"], m["desc"], dt_(d), qty, None, None])

# ----------------------------------------------------------------------------
# 4. Workbook 2: Position Saturation
# ----------------------------------------------------------------------------
POOL = [("Drive-in Bulk", 118), ("Drive-in Cases", 96), ("Roller Lanes", 64), ("Racks", 150)]
# dedicated positions are for mid-size flows (large volumes live in the shared areas)
ded_candidates = [m for m in MATS if 0 < dc_stock[m["code"]] / m["kp"] <= 14 and m["weekly"] > 0
                  and m["code"] not in (OTHER_ONLY_CODE, NO_SOURCE_CODE)]
ded_candidates.sort(key=lambda x: -x["weekly"])
chosen = ded_candidates[:24]
for c in SAFETY_CODES:                                  # all safety materials get dedicated positions
    mm = next(x for x in MATS if x["code"] == c)
    if mm not in chosen:
        chosen.append(mm)

ded_rows = []
for m in chosen:
    occ = max(1, math.ceil(dc_stock[m["code"]] / m["kp"]))
    mood = rnd.random()
    if mood < 0.12:                                     # overflow into shared positions
        tot = max(1, occ - rnd.randint(1, 3))
    elif mood < 0.22:                                   # exactly full
        tot = occ
    else:                                               # room to spare
        tot = occ + max(1, round(occ * rnd.uniform(0.25, 1.1)))
    ded_rows.append((m["code"], tot, occ))
if NO_DC_CODE not in [r[0] for r in ded_rows]:
    ded_rows.append((NO_DC_CODE, 6, 0))
ded_rows = [(c, t, 0 if c == NO_DC_CODE else o) for c, t, o in ded_rows]

pool_rows = [(n_, t, int(t * rnd.uniform(0.45, 0.8))) for n_, t in POOL]
blocked = (24, 11)
ded_sat = 1 - sum(t - o for _, t, o in ded_rows) / sum(t for _, t, _ in ded_rows)
pool_sat = 1 - sum(t - o for _, t, o in pool_rows) / sum(t for _, t, _ in pool_rows)

wb2 = openpyxl.Workbook()
w2 = wb2.active
w2.title = "Positions"
w2.append(["PROD/END", "Total", "Occupied", "Free", "Turnover (days)", None, "Date", "Shared saturation",
           "Dedicated saturation"])
# trend: 40 daily readings ending on the demo date
trend = []
d_val, s_val = rnd.uniform(66, 76), rnd.uniform(58, 68)
for i in range(40):
    d_val = min(93, max(52, d_val + rnd.uniform(-3.2, 3.4)))
    s_val = min(88, max(45, s_val + rnd.uniform(-2.6, 2.7)))
    trend.append((DEMO_DATE - dt.timedelta(days=39 - i), s_val / 100, d_val / 100))
trend[-1] = (DEMO_DATE, pool_sat, ded_sat)           # latest reading = headline KPI
trend_iter = iter(trend)

body = [(n_, t, o) for n_, t, o in pool_rows] + list(ded_rows) + [("Blocked",) + blocked]
for lab, t, o in body:
    r = [lab, t, o, t - o, rnd.randint(4, 40), None]
    nxt = next(trend_iter, None)
    if nxt:
        r += [dt_(nxt[0]), nxt[1], nxt[2]]
    w2.append(r)
for lab, val in [("Shared saturation", pool_sat), ("Dedicated saturation", ded_sat)]:
    r = [lab, val, None, None, None, None]
    nxt = next(trend_iter, None)
    if nxt:
        r += [dt_(nxt[0]), nxt[1], nxt[2]]
    w2.append(r)
for nxt in trend_iter:                                  # remaining history rows
    w2.append([None] * 6 + [dt_(nxt[0]), nxt[1], nxt[2]])

os.makedirs(OUT_DIR, exist_ok=True)
wb.save(os.path.join(OUT_DIR, "Stock Control.xlsx"))
wb2.save(os.path.join(OUT_DIR, "Position Saturation.xlsx"))
print(f"{len(MATS)} fictional materials written to {OUT_DIR}/")
print("safety-stock codes :", SAFETY_CODES)
print("excluded demo codes: FG-BAR-120, FG-EZ-210, MLT100, 1999, 90000001")
