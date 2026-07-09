"""
snapshot_cloud.py v9
Angel One Smart API → Excel (Nifty 50 snapshot) + Levels Excel + Email + Drive + Sheets
OB Signal read from GitHub Variable OB_SIGNAL (set BULLISH / BEARISH / NONE before run)
"""

import os, json, pyotp, gspread, traceback, smtplib, time, requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from SmartApi import SmartConnect
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

IST = ZoneInfo("Asia/Kolkata")
INSTRUMENT_MASTER_URL = "https://margincalculator.angelone.in/OpenAPI_File/files/OpenAPIScripMaster.json"

# ════════════════════════════════════════════════════════════
# LEVEL MULTIPLIERS (from your Excel — fixed % from prev close)
# ════════════════════════════════════════════════════════════
LEVEL_MULTIPLIERS = {
    "RAR":   ( 0.50881481481481472, "UP",   "TOP Level (Primary Bullish Target)"),
    "HS":    ( 0.32575444444444435, "UP",   "2nd Top (Secondary Bullish Target)"),
    "CS":    ( 0.31405555555555557, "UP",   "Upside Level 3"),
    "ALMSC": ( 0.25322222222222213, "UP",   "Upside Level 4"),
    "IFNT":  ( 0.23495833333333329, "UP",   "Upside Level 5"),
    "AL2n1": ( 0.21695611111111102, "UP",   "Upside Level 6"),
    "2n1":   ( 0.18991666666666659, "UP",   "Upside Level 7"),
    "MPLR":  ( 0.12661111111111106, "UP",   "Upside Level 8"),
    "TDAY":  (-0.11987271111111109, "DOWN", "2nd Bottom (Secondary Bearish Target)"),
    "NnN":   (-0.14973456790123463, "DOWN", "BOTTOM Level (Primary Bearish Target)"),
}

OB_TARGETS = {
    "BULLISH": {"primary": "RAR",  "secondary": "HS",   "label": "🟢 BULLISH OB", "color": "1B5E20", "bg": "E8F5E9"},
    "BEARISH": {"primary": "NnN",  "secondary": "TDAY", "label": "🔴 BEARISH OB", "color": "B71C1C", "bg": "FFEBEE"},
    "NONE":    {"primary": None,   "secondary": None,   "label": "⚪ No OB Signal","color": "757575", "bg": "F5F5F5"},
}

# ════════════════════════════════════════════════════════════
# CREDENTIALS
# ════════════════════════════════════════════════════════════
def env(key, default=""):
    return os.environ.get(key, default).strip()

def get_creds():
    return {
        "api_key":      env("ANGEL_API_KEY"),
        "client_id":    env("ANGEL_CLIENT_ID"),
        "password":     env("ANGEL_PASSWORD"),
        "totp_secret":  env("ANGEL_TOTP_SECRET"),
        "gmail_sender": env("GMAIL_SENDER"),
        "gmail_pass":   env("GMAIL_APP_PASSWORD"),
        "recipient":    env("RECIPIENT_EMAIL"),
        "drive_folder": env("GOOGLE_DRIVE_FOLDER_ID"),
        "sheet_id":     env("GOOGLE_SHEET_ID"),
        "sa_json":      env("SERVICE_ACCOUNT_JSON"),
        "ob_signal":    env("OB_SIGNAL", "NONE").upper(),
    }

def get_sa_creds(sa_json_str, scopes):
    return Credentials.from_service_account_info(json.loads(sa_json_str), scopes=scopes)

# ════════════════════════════════════════════════════════════
# EXCEL STYLES
# ════════════════════════════════════════════════════════════
C = {
    "title_bg":"0D47A1","hdr_mid":"1976D2","hdr_dark":"283593",
    "pos_bg":"E8F5E9","pos_alt":"F1F8E9","pos_txt":"1B5E20",
    "neg_bg":"FFEBEE","neg_alt":"FCE4EC","neg_txt":"B71C1C",
    "idx_bg":"E8EAF6","idx_alt":"FFFFFF",
    "grn_hdr":"2E7D32","red_hdr":"C62828","white":"FFFFFF",
    "yellow":"FFF176","yellow_lt":"FFF9C4",
}
def fill(h): return PatternFill("solid", start_color=h, fgColor=h)
def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Arial")
def aln(h="left", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
def brd():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)
def sc(cell, bg=None, fg="000000", bold=False, size=10, ha="left"):
    if bg: cell.fill = fill(bg)
    cell.font = font(bold, fg, size)
    cell.alignment = aln(ha)
    cell.border = brd()

# ════════════════════════════════════════════════════════════
# ANGEL ONE
# ════════════════════════════════════════════════════════════
def angel_login(creds):
    print("🔑 Logging in to Angel One...")
    obj  = SmartConnect(api_key=creds["api_key"])
    totp = pyotp.TOTP(creds["totp_secret"]).now()
    data = obj.generateSession(creds["client_id"], creds["password"], totp)
    if data["status"] is False:
        raise Exception("Login failed: " + str(data.get("message","")))
    print("✅ Login OK")
    return obj

def load_token_map():
    print("\n📥 Downloading instrument master...")
    r = requests.get(INSTRUMENT_MASTER_URL, timeout=30)
    r.raise_for_status()
    data = r.json()
    token_map = {}
    for item in data:
        if item.get("exch_seg") == "NSE" and item.get("symbol","").endswith("-EQ"):
            base = item["symbol"].replace("-EQ","")
            token_map[base] = {"token": item["token"], "tradingsymbol": item["symbol"]}
    print(f"  ✅ {len(token_map)} NSE equities loaded")
    return token_map

INDEX_TOKENS = {
    "NIFTY 50":       ("NSE", "Nifty 50",          "99926000"),
    "SENSEX":         ("BSE", "SENSEX",             "99919000"),
    "BANK NIFTY":     ("NSE", "Nifty Bank",         "99926009"),
    "NIFTY IT":       ("NSE", "Nifty IT",           "99926011"),
    "NIFTY SMALLCAP 50": ("NSE", "Nifty Smallcap 50", "99926061"),
}

def fetch_one_index(obj, name, exch, sym, token):
    try:
        r = obj.ltpData(exch, sym, token)
        if r.get("status") and r.get("data"):
            ltp   = float(r["data"].get("ltp",   0) or 0)
            close = float(r["data"].get("close", ltp) or ltp)
            chng  = round(ltp - close, 2)
            pct   = round((chng / close) * 100, 2) if close else 0.0
            return name, {"ltp": ltp, "chng": chng, "pct": pct, "close": close}
        return name, {"ltp": None, "chng": None, "pct": None, "close": None}
    except Exception as e:
        print(f"  ❌ {name}: {e}")
        return name, {"ltp": None, "chng": None, "pct": None, "close": None}

def fetch_indices(obj):
    print("\n📊 Fetching indices in parallel...")
    result = {}
    with ThreadPoolExecutor(max_workers=5) as ex:
        futures = {ex.submit(fetch_one_index, obj, name, exch, sym, token): name
                   for name, (exch, sym, token) in INDEX_TOKENS.items()}
        for future in as_completed(futures):
            name, data = future.result()
            result[name] = data
            pct = data.get("pct")
            ltp = data.get("ltp")
            if ltp:
                print(f"  ✅ {name}: {ltp}  {pct:+.2f}%")
            else:
                print(f"  ⚠️  {name}: no data")
    return result

NIFTY50_LIST = {
    "RELIANCE":9.25,"HDFCBANK":12.80,"ICICIBANK":8.85,"INFY":5.85,"TCS":3.90,
    "BHARTIARTL":4.20,"ITC":3.55,"LT":3.45,"KOTAKBANK":2.95,"AXISBANK":2.85,
    "SBIN":2.75,"HINDUNILVR":2.30,"BAJFINANCE":2.05,"M&M":1.95,"MARUTI":1.55,
    "SUNPHARMA":1.50,"ULTRACEMCO":1.35,"HCLTECH":1.30,"TITAN":1.25,"BAJAJFINSV":1.20,
    "ASIANPAINT":1.15,"ADANIENT":1.10,"NTPC":1.10,"ONGC":1.05,"POWERGRID":1.05,
    "TATASTEEL":1.00,"TECHM":0.95,"WIPRO":0.90,"ADANIPORTS":0.90,"COALINDIA":0.85,
    "JSWSTEEL":0.80,"BAJAJ-AUTO":0.80,"GRASIM":0.70,"DRREDDY":0.65,"CIPLA":0.65,
    "EICHERMOT":0.65,"INDUSINDBK":0.60,"APOLLOHOSP":0.55,"TATACONSUM":0.55,
    "BEL":0.50,"SHRIRAMFIN":0.45,"HINDALCO":0.75,"HDFCLIFE":0.90,"INDIGO":0.70,
    "JIOFIN":0.85,"MAXHEALTH":0.55,"SBILIFE":0.85,"TMPV":0.60,"TRENT":0.95,
    "ETERNAL":0.95,
}

def fetch_one_stock(obj, sym, wt, info, nifty_ltp):
    """Fetch a single stock — called in parallel."""
    if not info:
        wpts = round(nifty_ltp*(wt/100),2) if nifty_ltp else None
        return {"symbol":sym,"ltp":None,"chng":None,"pChng":None,
                "weight":wt,"weight_pts":wpts,"contrib":None}
    for attempt in range(3):
        try:
            r = obj.ltpData("NSE", info["tradingsymbol"], info["token"])
            if r.get("status") and r.get("data"):
                ltp   = float(r["data"].get("ltp",   0) or 0)
                close = float(r["data"].get("close", ltp) or ltp)
                chng  = round(ltp - close, 2)
                pct   = round((chng / close) * 100, 2) if close else 0.0
                contrib   = round(nifty_ltp * (pct/100) * (wt/100), 2) if nifty_ltp else None
                weight_pts= round(nifty_ltp * (wt/100), 2) if nifty_ltp else None
                return {"symbol":sym,"ltp":ltp,"chng":chng,"pChng":pct,
                        "weight":wt,"weight_pts":weight_pts,"contrib":contrib}
            else:
                msg = r.get("message","")
                if "rate" in msg.lower() and attempt < 2:
                    time.sleep(2); continue
                wpts = round(nifty_ltp*(wt/100),2) if nifty_ltp else None
                return {"symbol":sym,"ltp":None,"chng":None,"pChng":None,
                        "weight":wt,"weight_pts":wpts,"contrib":None}
        except Exception as e:
            if "rate" in str(e).lower() and attempt < 2:
                time.sleep(3); continue
            wpts = round(nifty_ltp*(wt/100),2) if nifty_ltp else None
            return {"symbol":sym,"ltp":None,"chng":None,"pChng":None,
                    "weight":wt,"weight_pts":wpts,"contrib":None}
    wpts = round(nifty_ltp*(wt/100),2) if nifty_ltp else None
    return {"symbol":sym,"ltp":None,"chng":None,"pChng":None,
            "weight":wt,"weight_pts":wpts,"contrib":None}

def fetch_nifty50(obj, nifty_ltp, token_map):
    print("\n📋 Fetching Nifty 50 stocks in parallel (10 at a time)...")
    symbols = list(NIFTY50_LIST.items())
    results = {}

    # Fetch in batches of 10 parallel — avoids rate limit
    BATCH = 10
    for i in range(0, len(symbols), BATCH):
        batch = symbols[i:i+BATCH]
        with ThreadPoolExecutor(max_workers=BATCH) as ex:
            futures = {ex.submit(fetch_one_stock, obj, sym, wt,
                                 token_map.get(sym), nifty_ltp): sym
                       for sym, wt in batch}
            for future in as_completed(futures):
                s = future.result()
                results[s["symbol"]] = s
                if s["ltp"]:
                    print(f"  ✅ {s['symbol']}: {s['ltp']}  {s['pChng']:+.2f}%")
        # Small pause between batches to avoid rate limit
        if i + BATCH < len(symbols):
            time.sleep(1)

    # Return in original NIFTY50_LIST order
    stocks = [results[sym] for sym, _ in symbols if sym in results]
    filled = sum(1 for s in stocks if s["ltp"] is not None)
    print(f"\n  Result: {filled}/{len(stocks)} stocks with data")
    return stocks

# ════════════════════════════════════════════════════════════
# LEVEL CALCULATOR
# ════════════════════════════════════════════════════════════
def calculate_levels(prev_close, nifty_ltp):
    levels = {}
    for name, (mult_pct, direction, desc) in LEVEL_MULTIPLIERS.items():
        level_val  = prev_close * (1 + mult_pct / 100)
        pts_change = round(level_val - prev_close, 2)
        dist_ltp   = round(level_val - nifty_ltp, 2)
        levels[name] = {
            "value":     round(level_val, 2),
            "pts":       pts_change,
            "pct":       round(mult_pct, 6),
            "direction": direction,
            "desc":      desc,
            "dist_ltp":  dist_ltp,
        }
    return levels

# ════════════════════════════════════════════════════════════
# BUILD SNAPSHOT EXCEL (File 1)
# ════════════════════════════════════════════════════════════
def build_snapshot_excel(label, ist_dt, indices, stocks):
    wb = Workbook(); ws = wb.active
    ws.title = "Snapshot_" + label
    disp = label[:2]+":"+label[2:] if len(label)==4 else label
    time_str = ist_dt.strftime("%d-%b-%Y %H:%M:%S")

    ws.merge_cells("A1:P1")
    c = ws["A1"]; c.value = f"NSE Market Snapshot  —  {disp} IST"
    sc(c, bg=C["title_bg"], fg=C["white"], bold=True, size=14, ha="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:P2")
    ws["A2"].value = f"Captured At (IST):  {time_str}"
    ws["A2"].fill = fill("E3F2FD"); ws["A2"].font = font(size=10, color="333333")
    ws["A2"].alignment = aln("center")

    r = 4
    ws.merge_cells(f"A{r}:D{r}")
    sc(ws.cell(r,1,"📊  INDEX SUMMARY"), bg=C["hdr_dark"], fg=C["white"], bold=True, size=11, ha="center")
    r += 1
    for h, col in zip(["INDEX","LTP (₹)","CHANGE (₹)","% CHANGE"], range(1,5)):
        sc(ws.cell(r,col,h), bg=C["hdr_mid"], fg=C["white"], bold=True, ha="center")

    for i, name in enumerate(["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]):
        r += 1; d = indices.get(name,{}); pct = d.get("pct")
        bg = C["idx_bg"] if i%2==0 else C["idx_alt"]
        pos = pct is not None and pct >= 0
        ws.cell(r,1,name).font = font(bold=True)
        ws.cell(r,1).fill=fill(bg); ws.cell(r,1).border=brd(); ws.cell(r,1).alignment=aln()
        for col, val, fmt in [(2,d.get("ltp"),"#,##0.00"),(3,d.get("chng"),"+#,##0.00;-#,##0.00"),
                               (4,pct and pct/100,"+0.00%;-0.00%")]:
            cell = ws.cell(r,col)
            if pct is not None:
                cell.value=val; cell.number_format=fmt
                cell.font=font(bold=True, color=C["pos_txt"] if pos else C["neg_txt"])
                cell.fill=fill(C["pos_bg"] if pos else C["neg_bg"])
            else:
                cell.value="N/A"; cell.font=font(color="9E9E9E"); cell.fill=fill(bg)
            cell.border=brd(); cell.alignment=aln("center")

    sr = r + 3
    ws.merge_cells(f"A{sr}:G{sr}")
    sc(ws.cell(sr,1,f"📋  ALL NIFTY 50 STOCKS  [{len(stocks)} stocks]"),
       bg=C["hdr_dark"],fg=C["white"],bold=True,size=11,ha="center")
    sr += 1
    for h, col in zip(["#","SYMBOL","LTP (₹)","CHANGE (₹)","% CHANGE","WEIGHT (pts)","CONTRIB (pts)"],range(1,8)):
        sc(ws.cell(sr,col,h),bg="3949AB",fg=C["white"],bold=True,ha="center")

    for i, s in enumerate(stocks):
        sr += 1
        pos = (s["pChng"] or 0) >= 0
        bg  = (C["pos_bg"] if pos else C["neg_bg"]) if i%2==0 else \
              (C["pos_alt"] if pos else C["neg_alt"])
        tc  = C["pos_txt"] if pos else C["neg_txt"]
        ws.cell(sr,1,i+1).fill=fill(bg);ws.cell(sr,1).border=brd();ws.cell(sr,1).alignment=aln("center")
        ws.cell(sr,2,s["symbol"]).font=font(bold=True)
        ws.cell(sr,2).fill=fill(bg);ws.cell(sr,2).border=brd();ws.cell(sr,2).alignment=aln()
        for col,val,fmt in [(3,s["ltp"],"#,##0.00"),(4,s["chng"],"+#,##0.00;-#,##0.00"),
                             (5,s["pChng"] and s["pChng"]/100,"+0.00%;-0.00%"),
                             (6,s.get("weight_pts"),"#,##0.00"),(7,s.get("contrib"),"+0.00;-0.00")]:
            cell=ws.cell(sr,col);cell.value=val
            if val is not None and fmt: cell.number_format=fmt
            cell.font=font(bold=(col in(4,5,7)),color=tc if col in(4,5,7) else "000000")
            cell.fill=fill(bg);cell.border=brd();cell.alignment=aln("center")

    valid=[s for s in stocks if s["pChng"] is not None]
    contrib_valid=[s for s in stocks if s.get("contrib") is not None]
    top7p=sorted(contrib_valid,key=lambda x:x["contrib"],reverse=True)[:7]
    top7n=sorted(contrib_valid,key=lambda x:x["contrib"])[:7]
    trow=r+4
    for top7,col,title,hbg,tc,b1,b2 in [
        (top7p,9,"🟢  TOP 7 POSITIVE",C["grn_hdr"],C["pos_txt"],C["pos_bg"],C["pos_alt"]),
        (top7n,15,"🔴  TOP 7 NEGATIVE",C["red_hdr"],C["neg_txt"],C["neg_bg"],C["neg_alt"]),
    ]:
        ws.merge_cells(start_row=trow,start_column=col,end_row=trow,end_column=col+4)
        sc(ws.cell(trow,col,title),bg=hbg,fg=C["white"],bold=True,size=11,ha="center")
        for h,dc in zip(["SYMBOL","LTP (₹)","CHG (₹)","% CHG","CONTRIB(pts)"],range(col,col+5)):
            sc(ws.cell(trow+1,dc,h),bg=hbg,fg=C["white"],bold=True,ha="center")
        for i,s in enumerate(top7):
            tr=trow+2+i;bg=b1 if i%2==0 else b2
            for dc,val,fmt in [(col,s["symbol"],None),(col+1,s["ltp"],"#,##0.00"),
                                (col+2,s["chng"],"+#,##0.00;-#,##0.00"),
                                (col+3,s["pChng"] and s["pChng"]/100,"+0.00%;-0.00%"),
                                (col+4,s.get("contrib"),"+0.00;-0.00")]:
                cell=ws.cell(tr,dc,val)
                if val is not None and fmt: cell.number_format=fmt
                cell.font=font(bold=True,color=tc);cell.fill=fill(bg)
                cell.border=brd();cell.alignment=aln("center" if dc>col else "left")

    for col,w in {1:4,2:14,3:12,4:12,5:11,6:13,7:13,9:14,10:12,11:11,12:10,13:13,
                  15:14,16:12,17:11,18:10,19:13}.items():
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.freeze_panes="A3"
    return wb

# ════════════════════════════════════════════════════════════
# BUILD LEVELS EXCEL (File 2)
# ════════════════════════════════════════════════════════════
def build_levels_excel(label, ist_dt, nifty_ltp, prev_close, ob_signal, indices, levels):
    wb = Workbook(); ws = wb.active
    ws.title = "Nifty Levels"
    disp = label[:2]+":"+label[2:] if len(label)==4 else label
    time_str = ist_dt.strftime("%d-%b-%Y %H:%M:%S")
    ob = OB_TARGETS.get(ob_signal, OB_TARGETS["NONE"])
    primary_name   = ob["primary"]
    secondary_name = ob["secondary"]

    # Title
    ws.merge_cells("A1:J1")
    c = ws["A1"]; c.value = f"NIFTY 50 LEVELS  —  {disp} IST"
    sc(c, bg="0D47A1", fg="FFFFFF", bold=True, size=14, ha="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    ws["A2"].value = (f"Captured: {time_str}  |  "
                      f"Prev Close (Base): {prev_close:,.2f}  |  "
                      f"Current LTP: {nifty_ltp:,.2f}  |  "
                      f"OB Signal: {ob['label']}")
    ws["A2"].fill=fill("E3F2FD"); ws["A2"].font=font(size=10,color="333333")
    ws["A2"].alignment=aln("center")

    # OB Signal Box
    ws.merge_cells("A4:J4")
    c = ws["A4"]
    if primary_name:
        pv = levels[primary_name]["value"]
        sv = levels[secondary_name]["value"]
        c.value = (f"{ob['label']}  →  "
                   f"⭐ PRIMARY TARGET: {primary_name} = {pv:,.2f}  "
                   f"({levels[primary_name]['pts']:+.2f} pts from prev close)     "
                   f"✅ SECONDARY TARGET: {secondary_name} = {sv:,.2f}  "
                   f"({levels[secondary_name]['pts']:+.2f} pts from prev close)")
    else:
        c.value = f"{ob['label']}  →  No OB signal. Showing all levels for reference."
    c.fill=fill(ob["bg"]); c.font=font(bold=True,color=ob["color"],size=12)
    c.alignment=aln("center"); c.border=brd()
    ws.row_dimensions[4].height = 30

    # Index summary
    r = 6
    ws.merge_cells(f"A{r}:J{r}")
    sc(ws.cell(r,1,"📊  INDEX SNAPSHOT"), bg="283593", fg="FFFFFF", bold=True, size=11, ha="center")
    r += 1
    for h, col in zip(["INDEX","LTP","CHANGE","% CHANGE"], range(1,5)):
        sc(ws.cell(r,col,h), bg="1976D2", fg="FFFFFF", bold=True, ha="center")
    for i, name in enumerate(["NIFTY 50","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50","SENSEX"]):
        r += 1; d=indices.get(name,{}); pct=d.get("pct")
        bg="E8EAF6" if i%2==0 else "FFFFFF"; pos=pct is not None and pct>=0
        ws.cell(r,1,name).font=font(bold=True)
        ws.cell(r,1).fill=fill(bg);ws.cell(r,1).border=brd();ws.cell(r,1).alignment=aln()
        for col,val,fmt in [(2,d.get("ltp"),"#,##0.00"),(3,d.get("chng"),"+#,##0.00;-#,##0.00"),
                             (4,pct and pct/100,"+0.00%;-0.00%")]:
            cell=ws.cell(r,col)
            if pct is not None:
                cell.value=val;cell.number_format=fmt
                cell.font=font(bold=True,color="1B5E20" if pos else "B71C1C")
                cell.fill=fill("E8F5E9" if pos else "FFEBEE")
            else:
                cell.value="N/A";cell.font=font(color="9E9E9E");cell.fill=fill(bg)
            cell.border=brd();cell.alignment=aln("center")

    # Levels table
    r += 2
    ws.merge_cells(f"A{r}:J{r}")
    sc(ws.cell(r,1,f"📐  NIFTY LEVELS  (All calculated from Prev Close: {prev_close:,.2f})"),
       bg="283593", fg="FFFFFF", bold=True, size=11, ha="center")
    r += 1
    hdrs = ["LEVEL","DESCRIPTION","DIRECTION","LEVEL VALUE","FROM PREV CLOSE (pts)",
            "FROM PREV CLOSE (%)","FROM CURRENT LTP (pts)","OB TARGET","SIGNAL"]
    for col,h in enumerate(hdrs,1):
        sc(ws.cell(r,col,h), bg="3949AB", fg="FFFFFF", bold=True, ha="center")

    # Upside levels (descending) then downside
    up = sorted([(n,v) for n,v in levels.items() if v["direction"]=="UP"],
                key=lambda x:x[1]["value"], reverse=True)
    dn = sorted([(n,v) for n,v in levels.items() if v["direction"]=="DOWN"],
                key=lambda x:x[1]["value"], reverse=True)

    for i, (name, lv) in enumerate(up + dn):
        r += 1
        is_primary   = (name == primary_name)
        is_secondary = (name == secondary_name)
        is_up = lv["direction"] == "UP"

        if is_primary:   bg = "FFF176"
        elif is_secondary: bg = "FFF9C4"
        elif is_up:      bg = "E8F5E9" if i%2==0 else "F1F8E9"
        else:            bg = "FFEBEE" if i%2==0 else "FCE4EC"

        tc = "1B5E20" if is_up else "B71C1C"
        ob_lbl = "⭐ PRIMARY TARGET" if is_primary else ("✅ SECONDARY TARGET" if is_secondary else "")
        sig_lbl = ob["label"] if (is_primary or is_secondary) else ""

        row_vals = [
            (name,                  True),
            (lv["desc"],            False),
            ("▲ UPSIDE" if is_up else "▼ DOWNSIDE", True),
            (lv["value"],           True),
            (lv["pts"],             True),
            (lv["pct"]/100,         True),
            (lv["dist_ltp"],        True),
            (ob_lbl,                True),
            (sig_lbl,               False),
        ]
        fmts = [None,None,None,"#,##0.00","+#,##0.00;-#,##0.00",
                "+0.0000%;-0.0000%","+#,##0.00;-#,##0.00",None,None]

        for col,(val,bold) in enumerate(row_vals,1):
            cell=ws.cell(r,col,val)
            cell_bg = "FF6F00" if (is_primary and col==1) else bg
            cell.fill=fill(cell_bg)
            cell.font=font(bold=bold or is_primary,
                           color=("FFFFFF" if (is_primary and col==1) else
                                  ("FF6F00" if is_primary else
                                   ("E65100" if is_secondary else tc))))
            cell.border=brd()
            cell.alignment=aln("left" if col in(2,8,9) else "center")
            if fmts[col-1] and val is not None:
                cell.number_format=fmts[col-1]
        ws.row_dimensions[r].height=18

    # Visual level guide
    r += 2
    ws.merge_cells(f"A{r}:J{r}")
    sc(ws.cell(r,1,"📈  VISUAL LEVEL GUIDE  (⭐ = Primary Target  ✅ = Secondary Target  🔵 = Current LTP)"),
       bg="283593", fg="FFFFFF", bold=True, size=11, ha="center")
    r += 1

    all_sorted = sorted([(n,v) for n,v in levels.items()],
                        key=lambda x:x[1]["value"], reverse=True)
    for name, lv in all_sorted:
        is_primary   = (name == primary_name)
        is_secondary = (name == secondary_name)
        is_up = lv["direction"] == "UP"

        # Insert current LTP marker between levels
        if lv["value"] <= nifty_ltp and (all_sorted.index((name,lv))==0 or
           all_sorted[all_sorted.index((name,lv))-1][1]["value"] > nifty_ltp):
            ws.merge_cells(f"A{r}:J{r}")
            c=ws.cell(r,1,f"  🔵  CURRENT LTP: {nifty_ltp:,.2f}")
            c.fill=fill("E3F2FD");c.font=font(bold=True,color="0D47A1",size=11)
            c.alignment=aln("left");c.border=brd()
            ws.row_dimensions[r].height=22
            r += 1

        ws.merge_cells(f"A{r}:J{r}")
        star = "⭐  " if is_primary else ("✅  " if is_secondary else "      ")
        arr  = "▲" if is_up else "▼"
        dist = abs(lv["dist_ltp"])
        c=ws.cell(r,1,
            f"  {star}{arr}  {name:<7}  {lv['value']:>10,.2f}  "
            f"({lv['pts']:>+7.2f} pts from prev close)    "
            f"[{dist:>7.2f} pts from LTP]    {lv['desc']}")
        bg = "FFF176" if is_primary else ("FFF9C4" if is_secondary else
             ("E8F5E9" if is_up else "FFEBEE"))
        tc = "1B5E20" if is_up else "B71C1C"
        c.fill=fill(bg)
        c.font=font(bold=is_primary or is_secondary,
                    color="FF6F00" if is_primary else ("E65100" if is_secondary else tc),
                    size=10)
        c.alignment=aln("left");c.border=brd()
        ws.row_dimensions[r].height=18
        r += 1

    # Column widths
    for col,w in {1:9,2:30,3:14,4:14,5:18,6:18,7:20,8:22,9:20}.items():
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.freeze_panes="A5"
    return wb

# ════════════════════════════════════════════════════════════
# EMAIL
# ════════════════════════════════════════════════════════════
def send_email(creds, snapshot_path, levels_path, label, ist_dt,
               indices, stocks, levels, ob_signal, nifty_ltp, prev_close):
    disp    = label[:2]+":"+label[2:] if len(label)==4 else label
    subject = f"📊 NSE Snapshot + Levels {disp} IST — {ist_dt.strftime('%d %b %Y')}"
    ob      = OB_TARGETS.get(ob_signal, OB_TARGETS["NONE"])
    primary_name   = ob["primary"]
    secondary_name = ob["secondary"]

    # Index HTML
    idx_html = ""
    for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
        d=indices.get(name,{}); pct=d.get("pct"); ltp=d.get("ltp")
        if pct is not None:
            color="#1b5e20" if pct>=0 else "#b71c1c"
            bg="#e8f5e9" if pct>=0 else "#ffebee"
            arrow="▲" if pct>=0 else "▼"
            idx_html += (f'<tr style="background:{bg}"><td style="padding:7px 14px;font-weight:bold">{name}</td>'
                f'<td style="padding:7px;color:{color};font-weight:bold;text-align:center">{arrow} {pct:+.2f}%</td>'
                f'<td style="padding:7px;text-align:center">₹{ltp:,.2f}</td></tr>')
        else:
            idx_html += f'<tr><td style="padding:7px 14px;font-weight:bold">{name}</td><td style="padding:7px;color:#9e9e9e;text-align:center">N/A</td><td>—</td></tr>'

    # OB signal box
    if primary_name:
        pv = levels[primary_name]["value"]
        sv = levels[secondary_name]["value"]
        ob_html = f"""
        <div style="background:{ob['bg']};border:2px solid #{ob['color']};border-radius:10px;
                    padding:16px;margin:16px 0;text-align:center">
          <div style="font-size:20px;font-weight:bold;color:#{ob['color']};margin-bottom:10px">
            {ob['label']}
          </div>
          <table width="100%" style="border-collapse:collapse">
            <tr>
              <td width="50%" style="text-align:center;padding:8px;
                  background:white;border-radius:8px;border:1px solid #{ob['color']}">
                <div style="font-size:11px;color:#666;margin-bottom:4px">⭐ PRIMARY TARGET</div>
                <div style="font-size:22px;font-weight:bold;color:#{ob['color']}">{primary_name}</div>
                <div style="font-size:18px;font-weight:bold">{pv:,.2f}</div>
                <div style="font-size:12px;color:#666">{levels[primary_name]['pts']:+.2f} pts from prev close</div>
              </td>
              <td width="4%"></td>
              <td width="46%" style="text-align:center;padding:8px;
                  background:white;border-radius:8px;border:1px solid #{ob['color']}">
                <div style="font-size:11px;color:#666;margin-bottom:4px">✅ SECONDARY TARGET</div>
                <div style="font-size:22px;font-weight:bold;color:#{ob['color']}">{secondary_name}</div>
                <div style="font-size:18px;font-weight:bold">{sv:,.2f}</div>
                <div style="font-size:12px;color:#666">{levels[secondary_name]['pts']:+.2f} pts from prev close</div>
              </td>
            </tr>
          </table>
          <div style="margin-top:10px;font-size:12px;color:#666">
            Prev Close: {prev_close:,.2f}  |  Current LTP: {nifty_ltp:,.2f}
          </div>
        </div>"""
    else:
        ob_html = '<div style="background:#f5f5f5;border-radius:8px;padding:12px;text-align:center;color:#757575;margin:16px 0">⚪ No OB Signal set — Set OB_SIGNAL variable in GitHub before next run</div>'

    # All levels table
    up_rows=""; dn_rows=""
    for name, lv in sorted(levels.items(), key=lambda x: x[1]["value"], reverse=True):
        is_p = (name==primary_name); is_s = (name==secondary_name)
        is_up = lv["direction"]=="UP"
        star = "⭐ " if is_p else ("✅ " if is_s else "")
        rbg = "#FFF176" if is_p else ("#FFF9C4" if is_s else ("#e8f5e9" if is_up else "#ffebee"))
        fw  = "bold" if (is_p or is_s) else "normal"
        tc  = "#1b5e20" if is_up else "#b71c1c"
        row = (f'<tr style="background:{rbg}">'
               f'<td style="padding:6px 10px;font-weight:{fw};color:{tc}">{star}{name}</td>'
               f'<td style="padding:6px 10px;font-weight:bold;text-align:right">{lv["value"]:,.2f}</td>'
               f'<td style="padding:6px 10px;color:{tc};text-align:right">{lv["pts"]:+.2f} pts</td>'
               f'<td style="padding:6px 10px;color:#666;text-align:right">{lv["dist_ltp"]:+.2f} from LTP</td></tr>')
        if is_up: up_rows += row
        else:     dn_rows += row

    levels_html = f"""
    <h3 style="color:#0d47a1;margin:16px 0 8px">📐 Nifty Levels (Base: Prev Close {prev_close:,.2f})</h3>
    <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0;margin-bottom:8px">
      <tr style="background:#1b5e20;color:white"><th style="padding:8px;text-align:left">▲ UPSIDE</th>
      <th style="padding:8px;text-align:right">Level</th><th style="padding:8px;text-align:right">From Close</th>
      <th style="padding:8px;text-align:right">From LTP</th></tr>{up_rows}
    </table>
    <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
      <tr style="background:#b71c1c;color:white"><th style="padding:8px;text-align:left">▼ DOWNSIDE</th>
      <th style="padding:8px;text-align:right">Level</th><th style="padding:8px;text-align:right">From Close</th>
      <th style="padding:8px;text-align:right">From LTP</th></tr>{dn_rows}
    </table>"""

    # Top 3 movers
    valid=[s for s in stocks if s.get("contrib") is not None]
    top3p=sorted(valid,key=lambda x:x["contrib"],reverse=True)[:3]
    top3n=sorted(valid,key=lambda x:x["contrib"])[:3]
    def srows(lst,color,bg):
        return "".join(f'<tr style="background:{bg}"><td style="padding:6px 12px;font-weight:bold">{s["symbol"]}</td>'
                       f'<td style="padding:6px;color:{color};font-weight:bold;text-align:center">{s["pChng"]:+.2f}%</td>'
                       f'<td style="padding:6px;color:{color};text-align:center">{s["contrib"]:+.2f} pts</td></tr>'
                       for s in lst)

    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:640px;margin:auto">
<div style="background:#0d47a1;color:white;padding:18px 22px;border-radius:10px 10px 0 0">
  <h2 style="margin:0;font-size:20px">📊 NSE Snapshot + Levels — {disp} IST</h2>
  <p style="margin:5px 0 0;opacity:.85;font-size:13px">{ist_dt.strftime('%d %b %Y  %H:%M:%S IST')}</p>
</div>
<div style="border:1px solid #ddd;border-top:none;padding:18px;border-radius:0 0 10px 10px">
  {ob_html}
  {levels_html}
  <h3 style="color:#0d47a1;margin:16px 0 8px">Index Performance</h3>
  <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
    <tr style="background:#1976d2;color:white">
      <th style="padding:8px 14px;text-align:left">Index</th>
      <th style="padding:8px">% Change</th><th style="padding:8px">LTP</th>
    </tr>{idx_html}
  </table>
  <table width="100%" style="margin-top:16px;border-collapse:collapse"><tr valign="top">
    <td width="50%" style="padding-right:8px">
      <h3 style="color:#2e7d32;margin:0 0 8px">🏆 Top 3 Boosters</h3>
      <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
        {srows(top3p,"#1b5e20","#e8f5e9") if top3p else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
      </table>
    </td>
    <td width="50%" style="padding-left:8px">
      <h3 style="color:#c62828;margin:0 0 8px">📉 Top 3 Draggers</h3>
      <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
        {srows(top3n,"#b71c1c","#ffebee") if top3n else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
      </table>
    </td>
  </tr></table>
  <div style="margin-top:16px;padding:12px 14px;background:#f5f5f5;border-radius:8px;font-size:13px">
    📎 2 Excel files attached: Snapshot (50 stocks) + Levels (OB targets)<br>
    ☁️ Both saved to Google Drive → NSE Snapshots folder
  </div>
</div></body></html>"""

    msg = MIMEMultipart("alternative")
    msg["From"]=creds["gmail_sender"]; msg["To"]=creds["recipient"]; msg["Subject"]=subject
    msg.attach(MIMEText(html,"html"))

    for path in [snapshot_path, levels_path]:
        with open(path,"rb") as f:
            part=MIMEBase("application","vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header("Content-Disposition","attachment",filename=os.path.basename(path))
            msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com",465) as s:
        s.login(creds["gmail_sender"],creds["gmail_pass"])
        s.sendmail(creds["gmail_sender"],creds["recipient"],msg.as_string())
    print(f"✅ Email sent with 2 attachments")

# ════════════════════════════════════════════════════════════
# GOOGLE DRIVE
# ════════════════════════════════════════════════════════════
def upload_drive(creds, xlsx_path):
    folder_id = creds["drive_folder"]
    if not folder_id or len(folder_id) < 10:
        print("❌ Drive folder ID missing"); return
    try:
        svc = build("drive","v3",credentials=get_sa_creds(
            creds["sa_json"],["https://www.googleapis.com/auth/drive"]))
        name = os.path.basename(xlsx_path)
        q    = f"name='{name}' and '{folder_id}' in parents and trashed=false"
        for f in svc.files().list(q=q,fields="files(id)",supportsAllDrives=True,
                includeItemsFromAllDrives=True).execute().get("files",[]):
            svc.files().delete(fileId=f["id"],supportsAllDrives=True).execute()
        meta  = {"name":name,"parents":[folder_id]}
        media = MediaFileUpload(xlsx_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False)
        res = svc.files().create(body=meta,media_body=media,
            fields="id,webViewLink",supportsAllDrives=True).execute()
        print(f"✅ Drive: {name}  →  {res.get('webViewLink','')}")
    except Exception as e:
        print(f"❌ Drive upload failed: {e}")

# ════════════════════════════════════════════════════════════
# GOOGLE SHEETS
# ════════════════════════════════════════════════════════════
def update_sheets(creds, label, ist_dt, indices, stocks, levels, ob_signal, prev_close):
    try:
        gc  = gspread.authorize(get_sa_creds(
            creds["sa_json"],["https://www.googleapis.com/auth/spreadsheets"]))
        sh  = gc.open_by_key(creds["sheet_id"])
        disp = label[:2]+":"+label[2:] if len(label)==4 else label
        ob   = OB_TARGETS.get(ob_signal, OB_TARGETS["NONE"])

        # Tab 1: Snapshot
        tab1 = "Snapshot_" + label
        try: sh.del_worksheet(sh.worksheet(tab1))
        except: pass
        ws1 = sh.add_worksheet(title=tab1, rows=120, cols=20)
        rows1 = [
            [f"NSE Snapshot — {disp} IST"]+[""]*19,
            [f"Captured: {ist_dt.strftime('%d-%b-%Y %H:%M:%S IST')}"]+[""]*19,
            [""]*20,
            ["INDEX","LTP","CHG","% CHG",""]+["TOP 7 +ve","","","","",""]+["TOP 7 -ve","","","",""]
        ]
        for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
            d=indices.get(name,{}); p=d.get("pct")
            rows1.append([name,d.get("ltp","N/A"),d.get("chng","N/A"),
                          f"{p:+.2f}%" if p is not None else "N/A",""]+[""]*15)
        rows1.append([""]*20)
        rows1.append(["#","SYMBOL","LTP","CHG","% CHG","WT(pts)","CONTRIB","",
                      "SYMBOL","LTP","CHG","% CHG","CONTRIB","",
                      "SYMBOL","LTP","CHG","% CHG","CONTRIB"])
        valid=[s for s in stocks if s.get("contrib") is not None]
        top7p=sorted(valid,key=lambda x:x["contrib"],reverse=True)[:7]
        top7n=sorted(valid,key=lambda x:x["contrib"])[:7]
        for i, s in enumerate(stocks):
            p=s["pChng"]
            row=[i+1,s["symbol"],s.get("ltp",""),
                 f'{s["chng"]:+.2f}' if s["chng"] is not None else "",
                 f'{p:+.2f}%' if p is not None else "",
                 f'{s["weight_pts"]:.2f}' if s.get("weight_pts") is not None else "",
                 f'{s["contrib"]:+.2f}' if s.get("contrib") is not None else "",""]
            row+=([top7p[i]["symbol"],top7p[i]["ltp"],
                   f'{top7p[i]["chng"]:+.2f}',f'{top7p[i]["pChng"]:+.2f}%',
                   f'{top7p[i]["contrib"]:+.2f}',""] if i<len(top7p) else [""]*6)
            row+=([top7n[i]["symbol"],top7n[i]["ltp"],
                   f'{top7n[i]["chng"]:+.2f}',f'{top7n[i]["pChng"]:+.2f}%',
                   f'{top7n[i]["contrib"]:+.2f}'] if i<len(top7n) else [""]*5)
            rows1.append(row)
        ws1.update("A1",rows1)
        print(f"✅ Sheets: {tab1}")

        # Tab 2: Levels
        tab2 = "Levels_" + label
        try: sh.del_worksheet(sh.worksheet(tab2))
        except: pass
        ws2 = sh.add_worksheet(title=tab2, rows=30, cols=10)
        rows2 = [
            [f"NIFTY LEVELS — {disp} IST"]+[""]*9,
            [f"Prev Close: {prev_close:,.2f}  |  OB Signal: {ob['label']}"]+[""]*9,
            [""]*10,
        ]
        if ob["primary"]:
            rows2.append([f"⭐ PRIMARY: {ob['primary']} = {levels[ob['primary']]['value']:,.2f}",
                          f"✅ SECONDARY: {ob['secondary']} = {levels[ob['secondary']]['value']:,.2f}"]
                         +[""]*8)
        rows2.append([""]*10)
        rows2.append(["LEVEL","DESCRIPTION","DIRECTION","VALUE","PTS FROM CLOSE",
                      "% FROM CLOSE","PTS FROM LTP","OB TARGET","",""])
        for name, lv in sorted(levels.items(), key=lambda x: x[1]["value"], reverse=True):
            ob_lbl = "⭐ PRIMARY" if name==ob["primary"] else ("✅ SECONDARY" if name==ob["secondary"] else "")
            rows2.append([name, lv["desc"], lv["direction"], lv["value"],
                          f'{lv["pts"]:+.2f}', f'{lv["pct"]/100:+.4f}%',
                          f'{lv["dist_ltp"]:+.2f}', ob_lbl, "", ""])
        ws2.update("A1",rows2)
        print(f"✅ Sheets: {tab2}")

    except Exception as e:
        print(f"⚠️  Sheets update failed: {e}")
        traceback.print_exc()

# ════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════
def main():
    ist_dt    = datetime.now(IST)
    label     = env("MANUAL_LABEL") or ist_dt.strftime("%H%M")
    disp      = label[:2]+":"+label[2:] if len(label)==4 else label
    ob_signal = env("OB_SIGNAL","NONE").upper()

    print(f"\n{'='*60}")
    print(f"  NSE Snapshot + Levels — {disp} IST | {ist_dt.strftime('%d-%b-%Y')}")
    print(f"  OB Signal: {ob_signal}")
    print(f"{'='*60}")

    creds      = get_creds()
    obj        = angel_login(creds)
    token_map  = load_token_map()
    indices    = fetch_indices(obj)

    nifty_ltp   = indices.get("NIFTY 50",{}).get("ltp")   or 0
    prev_close  = indices.get("NIFTY 50",{}).get("close") or 0
    print(f"\n  Nifty LTP: {nifty_ltp}  |  Prev Close: {prev_close}")

    stocks = fetch_nifty50(obj, nifty_ltp, token_map)

    # Calculate levels
    print("\n📐 Calculating levels...")
    levels = calculate_levels(prev_close, nifty_ltp)
    ob = OB_TARGETS.get(ob_signal, OB_TARGETS["NONE"])
    print(f"  OB Signal: {ob['label']}")
    if ob["primary"]:
        print(f"  ⭐ Primary Target:   {ob['primary']} = {levels[ob['primary']]['value']:,.2f}")
        print(f"  ✅ Secondary Target: {ob['secondary']} = {levels[ob['secondary']]['value']:,.2f}")

    os.makedirs("output", exist_ok=True)
    date_str = ist_dt.strftime("%Y-%m-%d")

    # Build Excel 1 — Snapshot
    print("\n📁 Building Snapshot Excel...")
    wb1 = build_snapshot_excel(label, ist_dt, indices, stocks)
    snap_path = f"output/NSE_Snapshot_{date_str}_{label}.xlsx"
    wb1.save(snap_path)
    print(f"  ✅ {snap_path}")

    # Build Excel 2 — Levels
    print("\n📐 Building Levels Excel...")
    wb2 = build_levels_excel(label, ist_dt, nifty_ltp, prev_close, ob_signal, indices, levels)
    lvl_path = f"output/NSE_Levels_{date_str}_{label}.xlsx"
    wb2.save(lvl_path)
    print(f"  ✅ {lvl_path}")

    # Send email with BOTH files attached
    print("\n📧 Sending email...")
    send_email(creds, snap_path, lvl_path, label, ist_dt,
               indices, stocks, levels, ob_signal, nifty_ltp, prev_close)

    # Upload BOTH to Drive
    print("\n☁️  Uploading to Drive...")
    upload_drive(creds, snap_path)
    upload_drive(creds, lvl_path)

    # Update Sheets
    print("\n📊 Updating Google Sheets...")
    update_sheets(creds, label, ist_dt, indices, stocks, levels, ob_signal, prev_close)

    print(f"\n✅ ALL DONE — {disp} IST\n")

def calculate_levels(prev_close, nifty_ltp):
    levels = {}
    for name, (mult_pct, direction, desc) in LEVEL_MULTIPLIERS.items():
        level_val  = prev_close * (1 + mult_pct / 100)
        pts_change = round(level_val - prev_close, 2)
        dist_ltp   = round(level_val - nifty_ltp, 2)
        levels[name] = {
            "value":     round(level_val, 2),
            "pts":       pts_change,
            "pct":       round(mult_pct, 6),
            "direction": direction,
            "desc":      desc,
            "dist_ltp":  dist_ltp,
        }
    return levels

if __name__ == "__main__":
    main()
