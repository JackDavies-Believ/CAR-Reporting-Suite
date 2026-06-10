#!/usr/bin/env python3
"""
CAR Reporting Suite — build script
===================================
Drop your four Salesforce XLS exports into the data/ folder, then run:

    python build.py

This rebuilds index.html with the latest data embedded.

Expected files in data/:
    account_ops.xls         — Account-level ops report (Account Name, PM, AM, Contract Stage, etc.)
    site_data.xls           — Site-level report (Site Name, Status, Technical Stage, Bays, PM, AM, etc.)
    forecast_movement.xls   — Charger Install Sign Off (F) date change history
    status_movement.xls     — Site status change history
"""

import json
import re
import sys
import os
from html.parser import HTMLParser
from datetime import datetime, timedelta
from collections import defaultdict

DATA_DIR    = os.path.join(os.path.dirname(__file__), "data")
TEMPLATE    = os.path.join(os.path.dirname(__file__), "template.html")
OUTPUT      = os.path.join(os.path.dirname(__file__), "index.html")

EXPECTED_FILES = {
    "account_ops":        "account_ops.xls",
    "site_data":          "site_data.xls",
    "forecast_movement":  "forecast_movement.xls",
    "status_movement":    "status_movement.xls",
}

# ── HTML table parser ──────────────────────────────────────────────────────────

class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []; self.cur = []; self.cell = ""; self.in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":   self.cur = []
        elif tag in ("td","th"): self.in_cell = True; self.cell = ""

    def handle_endtag(self, tag):
        if tag == "tr":
            if self.cur: self.rows.append(self.cur[:])
        elif tag in ("td","th"):
            self.cur.append(self.cell.strip()); self.in_cell = False

    def handle_data(self, data):
        if self.in_cell: self.cell += data


def parse_xls(path):
    with open(path, "r", encoding="iso-8859-1") as f:
        content = f.read()
    p = TableParser()
    p.feed(content)
    if not p.rows:
        return [], []
    headers = p.rows[0]
    return headers, p.rows[1:]


# ── Date helpers ───────────────────────────────────────────────────────────────

def parse_date(s):
    s = s.strip()
    if not s: return None
    for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
        try:
            d = datetime.strptime(s, fmt).date()
            if d.year < 2020: return None
            return d
        except: pass
    return None

def parse_dt(s):
    s = s.strip()
    if not s: return None
    for fmt in ["%d/%m/%Y, %H:%M", "%d/%m/%Y"]:
        try: return datetime.strptime(s, fmt)
        except: pass
    return None

def week_of(dt):
    mon = dt - timedelta(days=dt.weekday())
    return mon.strftime("%Y-%m-%d")


# ── Parse account ops report ───────────────────────────────────────────────────

def parse_account_ops(path):
    headers, rows = parse_xls(path)
    h = headers

    def col(name, fallback=None):
        try: return h.index(name)
        except ValueError: return fallback

    acc_c   = col("Account Name",           col("Account"))
    pm_c    = col("Programme Manager",       col("Account: Programme Manager: Full Name"))
    am_c    = col("Account Manager*",        col("Account: Account Manager*: Full Name"))
    stage_c = col("Account Contract Stage",  col("Site Contract Stage"))
    target_c= col("Operational Target",      col("Operational Target"))
    cplive_c= col("Sockets CP Live",         col("Sockets CP Live"))
    ssbal_c = col("Site Selections Balance",  col("Site Selections Balance"))

    acc_to_am = {}
    account_data = []

    for r in rows:
        def v(c): return r[c].strip() if c is not None and c < len(r) else ""
        acc = v(acc_c)
        if not acc: continue
        am = v(am_c)
        if am: acc_to_am[acc] = am

        def num(c):
            try: return int(float(v(c)))
            except: return 0

        account_data.append({
            "account":       acc,
            "pm":            v(pm_c),
            "am":            am,
            "contractStage": v(stage_c),
            "opsTarget":     num(target_c),
            "cpLive":        num(cplive_c),
            "ssBalance":     num(ssbal_c),
            "socketsInFlow": 0,   # calculated below if site data available
        })

    pms      = sorted(set(r["pm"]      for r in account_data if r["pm"]))
    ams      = sorted(set(r["am"]      for r in account_data if r["am"]))
    accounts = sorted(set(r["account"] for r in account_data))

    return account_data, acc_to_am, pms, ams, accounts


# ── Parse site data report ─────────────────────────────────────────────────────

def parse_site_data(path, acc_to_am):
    headers, rows = parse_xls(path)
    h = headers

    def col(name, fallback=None):
        try: return h.index(name)
        except ValueError: return fallback

    acc_c    = col("Account: Account Name", col("Account"))
    site_c   = col("Site Name",             col("Site History: Site History ID"))
    status_c = col("Site Status",           col("Status"))
    ts_c     = col("Technical Stage",       col("Technical Stage"))
    stage_c  = col("Site Contract Stage",   col("Account Contract Stage"))
    resp_c   = col("Site Responsibility",   None)
    bays_c   = col("Total No. Active Bays", col("Active Bays"))
    pm_c     = col("Account: Programme Manager: Full Name", col("Programme Manager"))
    am_c     = col("Account: Account Manager*: Full Name",  None)
    fi_c     = col("Charger(s) Install Sign Off (F)",       None)
    ai_c     = col("Charger(s) Install Sign Off (A)",       None)

    site_data = []
    for r in rows:
        def v(c): return r[c].strip() if c is not None and c < len(r) else ""
        acc = v(acc_c)
        if not acc: continue
        try: bays = int(float(v(bays_c)))
        except: bays = 0

        fi_date = v(fi_c)
        is_cp   = v(status_c) == "CP Live"

        # Forecast month from Charger Install Sign Off (F)
        f_date = parse_date(fi_date)
        forecast_month = f_date.month if f_date else None
        forecast_year  = f_date.year  if f_date else None
        forecast_date  = fi_date if fi_date else None

        site_data.append({
            "account":       acc,
            "site":          v(site_c),
            "status":        v(status_c),
            "technicalStage":v(ts_c),
            "contractStage": v(stage_c),
            "responsibility":v(resp_c) if resp_c is not None else "Unknown",
            "bays":          bays,
            "pm":            v(pm_c),
            "am":            v(am_c) if am_c is not None else acc_to_am.get(acc, ""),
            "forecastDate":  forecast_date,
            "forecastMonth": forecast_month,
            "forecastYear":  forecast_year,
            "isCPLive":      is_cp,
        })

    return site_data


# ── Parse status movement report ──────────────────────────────────────────────

def parse_status_movement(path):
    headers, rows = parse_xls(path)
    h = headers

    def col(name, fallback=None):
        try: return h.index(name)
        except ValueError: return fallback

    site_c   = col("Site Name",                          col("Site History: Site History ID"))
    acc_c    = col("Account",                            col("Account: Account Name"))
    old_c    = col("Old Value",                          None)
    new_c    = col("New Value",                          None)
    date_c   = col("Modify Date",                        None)
    bays_c   = col("Total No. Active Bays",              col("Active Bays"))
    field_c  = col("Field",                              None)

    STATUSES = [
        "1. Site Selection","2. Site Selection Client Approval","3. Surveys and HLD",
        "4. Awaiting POC","5. HLD Review","6. HLD Client Approval","7. Contract Negotiation",
        "8. ATP Preparation","9. DD Preparation","10. DD Review","11. DD Client Approval",
        "12. Connection Legals","13. Public Consultation","14. ATB Preparation",
        "15. Equipment Procurement","16. Legals","17. Permitting","18. Build in Progress",
        "19. Civils complete - awaiting power","20. Meter pending","21. Commissioning Pending",
        "CP Live","On Hold","Terminated"
    ]

    def v(r, c): return r[c].strip() if c is not None and c < len(r) else ""

    records = []; created = []; accounts = set()
    for r in rows:
        site  = v(r, site_c); acc = v(r, acc_c)
        if not site or not acc: continue
        old   = v(r, old_c)  if old_c  is not None else ""
        new   = v(r, new_c)  if new_c  is not None else ""
        field = v(r, field_c) if field_c is not None else "Site Status"
        raw_d = v(r, date_c) if date_c is not None else ""
        try: bays = int(float(v(r, bays_c)))
        except: bays = 0
        dt = parse_dt(raw_d)
        if not dt: continue
        accounts.add(acc)
        iso  = dt.strftime("%Y-%m-%d")
        wk   = week_of(dt)
        if field == "Created":
            created.append({"s": site, "a": acc, "d": iso, "w": wk, "b": bays})
        elif field == "Site Status" and old != new:
            records.append({"s": site, "a": acc, "o": old, "n": new, "d": iso, "w": wk, "b": bays})

    return {"records": records, "created": created,
            "statuses": STATUSES, "accounts": sorted(accounts)}


# ── Parse forecast movement report ────────────────────────────────────────────

def parse_forecast_movement(path):
    headers, rows = parse_xls(path)
    h = headers

    def col(name, fallback=None):
        try: return h.index(name)
        except ValueError: return fallback

    pid_c    = col("Project ID",   None)
    acc_c    = col("Account",      col("Account: Account Name"))
    batch_c  = col("Batch No.",    None)
    site_c   = col("Site Name",    col("Site History: Site History ID"))
    old_c    = col("Old Value",    None)
    new_c    = col("New Value",    None)
    date_c   = col("Modify Date",  None)
    bays_c   = col("Active Bays",  col("Total No. Active Bays"))
    status_c = col("Site Status",  None)

    def v(r, c): return r[c].strip() if c is not None and c < len(r) else ""

    weeks_data = {}; accounts = set()
    for r in rows:
        acc  = v(r, acc_c); site = v(r, site_c)
        if not acc or not site: continue
        old_v   = v(r, old_c); new_v = v(r, new_c)
        status  = v(r, status_c)
        raw_d   = v(r, date_c)
        try: bays = int(float(v(r, bays_c)))
        except: bays = 0
        dt = parse_dt(raw_d)
        if not dt: continue
        wk = week_of(dt); accounts.add(acc)

        old_d = parse_date(old_v); new_d = parse_date(new_v)
        rec = {"pid": v(r,pid_c), "acc": acc, "batch": v(r,batch_c),
               "site": site, "old": old_v, "new": new_v,
               "date": dt.strftime("%Y-%m-%d"), "bays": bays, "status": status}

        if status == "Terminated":      rec["type"] = "term"
        elif not old_v and new_v and new_d: rec["type"] = "new"
        elif old_v and not new_v:       rec["type"] = "blank"
        elif old_d and new_d:
            rec["type"] = "push" if new_d > old_d else "pull"
        else: continue

        if wk not in weeks_data:
            weeks_data[wk] = {
                "pushed_b":0,"pulled_b":0,"blanked_b":0,"new_b":0,"term_b":0,
                "pushed_s":0,"pulled_s":0,"blanked_s":0,"new_s":0,"term_s":0,
                "records":[]
            }
        wd = weeks_data[wk]; t = rec["type"]
        short = {"push":"pushed","pull":"pulled","blank":"blanked","new":"new","term":"term"}[t]
        wd[short+"_b"] += bays; wd[short+"_s"] += 1
        wd["records"].append(rec)

    weeks = sorted(weeks_data.keys(), reverse=True)
    return {"weeks": weeks, "accounts": sorted(accounts), "data": weeks_data}


# ── Main build ────────────────────────────────────────────────────────────────

def main():
    print("CAR Reporting Suite — build script")
    print("=" * 40)

    # Check data files
    missing = []
    paths = {}
    for key, filename in EXPECTED_FILES.items():
        p = os.path.join(DATA_DIR, filename)
        if not os.path.exists(p):
            missing.append(filename)
        else:
            paths[key] = p
            print(f"  ✓ {filename}")

    if missing:
        print(f"\n  ✗ Missing files in data/:")
        for m in missing: print(f"      {m}")
        print("\nPlace your Salesforce XLS exports in the data/ folder with the names above.")
        print("You can rename your exports or update the EXPECTED_FILES dict at the top of this script.")
        sys.exit(1)

    print("\nParsing reports...")

    # Account ops
    account_data, acc_to_am, pms, ams, accounts = parse_account_ops(paths["account_ops"])
    print(f"  Account ops:        {len(account_data)} accounts, {len(pms)} PMs, {len(ams)} AMs")

    # Site data
    site_data = parse_site_data(paths["site_data"], acc_to_am)
    print(f"  Site data:          {len(site_data)} sites")

    # Back-fill socketsInFlow from site data
    flow_by_acc = defaultdict(int)
    for s in site_data:
        if not s["isCPLive"] and s["forecastYear"] and s["forecastYear"] >= datetime.now().year:
            flow_by_acc[s["account"]] += s["bays"]
    for rec in account_data:
        rec["socketsInFlow"] = flow_by_acc.get(rec["account"], rec["socketsInFlow"])

    # Forecast data (sites with install date this year)
    forecast_data = [s for s in site_data if s["forecastYear"] and s["forecastYear"] >= datetime.now().year]
    print(f"  Forecast data:      {len(forecast_data)} sites with install date this year+")

    # Status movement
    movement_db = parse_status_movement(paths["status_movement"])
    print(f"  Status movement:    {len(movement_db['records'])} records, {len(movement_db['accounts'])} accounts")

    # Forecast movement
    fm_data = parse_forecast_movement(paths["forecast_movement"])
    print(f"  Forecast movement:  {sum(len(w['records']) for w in fm_data['data'].values())} records, {len(fm_data['weeks'])} weeks")

    # Build RAW object
    RAW = {
        "accountData":  account_data,
        "siteData":     site_data,
        "forecastData": forecast_data,
        "pms":          pms,
        "ams":          ams,
        "accounts":     accounts,
    }

    # Read template
    with open(TEMPLATE, "r", encoding="utf-8") as f:
        html = f.read()

    # Inject RAW
    raw_pattern = re.compile(r'const RAW = \{.*?\};', re.DOTALL)
    if not raw_pattern.search(html):
        print("\n✗ ERROR: Could not find 'const RAW = {...};' in template.html")
        sys.exit(1)
    raw_json = json.dumps(RAW, ensure_ascii=False, separators=(",",":"))
    html = raw_pattern.sub(f'const RAW = {raw_json};', html, count=1)

    # Inject movement DB
    db_pattern = re.compile(r'const DB = \{.*?\};', re.DOTALL)
    if db_pattern.search(html):
        db_json = json.dumps(movement_db, ensure_ascii=False, separators=(",",":"))
        html = db_pattern.sub(f'const DB = {db_json};', html, count=1)

    # Inject FM_DATA
    fm_pattern = re.compile(r'const FM_DATA = \{.*?\};', re.DOTALL)
    if fm_pattern.search(html):
        fm_json = json.dumps(fm_data, ensure_ascii=False, separators=(",",":"))
        html = fm_pattern.sub(f'const FM_DATA = {fm_json};', html, count=1)

    # Update last-updated timestamp
    ts = datetime.now().strftime("%d %b %Y %H:%M")
    html = html.replace('id="last-updated"></div>', f'id="last-updated">{ts}</div>')

    # Write output
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT) // 1024
    print(f"\n✓ Built index.html ({size_kb} KB)")
    print(f"  Last updated: {ts}")
    print("\nDone! Open index.html in your browser or push to GitHub.")


if __name__ == "__main__":
    main()
