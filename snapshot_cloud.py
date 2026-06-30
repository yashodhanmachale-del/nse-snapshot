"""
snapshot_cloud.py v7 — Added Nifty point contribution per stock
Contribution(pts) = Nifty Level × (Stock %Chg/100) × Stock Weight%
"""

import os, json, pyotp, gspread, traceback, smtplib, time
from SmartApi import SmartConnect
from datetime import datetime
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

def env(key):
    val = os.environ.get(key, "")
    if not val:
        raise ValueError(f"Missing secret: {key}")
    return val.strip()

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
    }

def get_sa_creds(sa_json_str, scopes):
    return Credentials.from_service_account_info(json.loads(sa_json_str), scopes=scopes)

# ── Styles ───────────────────────────────────────────────────
C = {
    "title_bg":"0D47A1","hdr_mid":"1976D2","hdr_dark":"283593",
    "pos_bg":"E8F5E9","pos_alt":"F1F8E9","pos_txt":"1B5E20",
    "neg_bg":"FFEBEE","neg_alt":"FCE4EC","neg_txt":"B71C1C",
    "idx_bg":"E8EAF6","idx_alt":"FFFFFF",
    "grn_hdr":"2E7D32","red_hdr":"C62828","white":"FFFFFF",
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

# ── Angel One Login ──────────────────────────────────────────
def angel_login(creds):
    print("🔑 Logging in to Angel One...")
    obj  = SmartConnect(api_key=creds["api_key"])
    totp = pyotp.TOTP(creds["totp_secret"]).now()
    data = obj.generateSession(creds["client_id"], creds["password"], totp)
    if data["status"] is False:
        raise Exception("Login failed: " + str(data.get("message","")))
    print("✅ Login OK")
    return obj

INDEX_TOKENS = {
    "NIFTY 50":          ("NSE", "Nifty 50",          "99926000"),
    "SENSEX":            ("BSE", "SENSEX",             "99919000"),
    "BANK NIFTY":        ("NSE", "Nifty Bank",         "99926009"),
    "NIFTY IT":          ("NSE", "Nifty IT",           "99926011"),
    "NIFTY SMALLCAP 50": ("NSE", "Nifty Smallcap 50", "99926074"),
}

def fetch_indices(obj):
    print("\n📊 Fetching indices...")
    result = {}
    for name, (exch, sym, token) in INDEX_TOKENS.items():
        try:
            r = obj.ltpData(exch, sym, token)
            if r.get("status") and r.get("data"):
                ltp   = float(r["data"].get("ltp",   0) or 0)
                close = float(r["data"].get("close", ltp) or ltp)
                chng  = round(ltp - close, 2)
                pct   = round((chng / close) * 100, 2) if close else 0.0
                result[name] = {"ltp": ltp, "chng": chng, "pct": pct}
                print(f"  ✅ {name}: {ltp}  {pct:+.2f}%")
            else:
                result[name] = {"ltp": None, "chng": None, "pct": None}
        except Exception as e:
            print(f"  ❌ {name}: {e}")
            result[name] = {"ltp": None, "chng": None, "pct": None}
        time.sleep(1)
    return result

# ── Stock tokens + Official Nifty 50 weights (NSE index methodology, % of free-float mcap) ─
# Weights approximate as of latest Nifty rebalance — update periodically from nseindia.com
NIFTY50_STOCKS = {
    "RELIANCE":   {"tok": "2885",  "ts": "RELIANCE-EQ",   "wt": 9.25},
    "HDFCBANK":   {"tok": "1333",  "ts": "HDFCBANK-EQ",   "wt": 12.80},
    "ICICIBANK":  {"tok": "4963",  "ts": "ICICIBANK-EQ",  "wt": 8.85},
    "INFY":       {"tok": "1594",  "ts": "INFY-EQ",       "wt": 5.85},
    "TCS":        {"tok": "11536", "ts": "TCS-EQ",        "wt": 3.90},
    "BHARTIARTL": {"tok": "10604", "ts": "BHARTIARTL-EQ", "wt": 4.20},
    "ITC":        {"tok": "1660",  "ts": "ITC-EQ",        "wt": 3.55},
    "LT":         {"tok": "11483", "ts": "LT-EQ",         "wt": 3.45},
    "KOTAKBANK":  {"tok": "1922",  "ts": "KOTAKBANK-EQ",  "wt": 2.95},
    "AXISBANK":   {"tok": "5900",  "ts": "AXISBANK-EQ",   "wt": 2.85},
    "SBIN":       {"tok": "3045",  "ts": "SBIN-EQ",       "wt": 2.75},
    "HINDUNILVR": {"tok": "1394",  "ts": "HINDUNILVR-EQ", "wt": 2.30},
    "BAJFINANCE": {"tok": "317",   "ts": "BAJFINANCE-EQ", "wt": 2.05},
    "M&M":        {"tok": "2031",  "ts": "M&M-EQ",        "wt": 1.95},
    "MARUTI":     {"tok": "10999", "ts": "MARUTI-EQ",     "wt": 1.55},
    "SUNPHARMA":  {"tok": "3351",  "ts": "SUNPHARMA-EQ",  "wt": 1.50},
    "TATAMOTORS": {"tok": "3456",  "ts": "TATAMOTORS-EQ", "wt": 1.40},
    "ULTRACEMCO": {"tok": "11532", "ts": "ULTRACEMCO-EQ", "wt": 1.35},
    "HCLTECH":    {"tok": "7229",  "ts": "HCLTECH-EQ",    "wt": 1.30},
    "TITAN":      {"tok": "3506",  "ts": "TITAN-EQ",      "wt": 1.25},
    "BAJAJFINSV": {"tok": "16675", "ts": "BAJAJFINSV-EQ", "wt": 1.20},
    "ASIANPAINT": {"tok": "236",   "ts": "ASIANPAINT-EQ", "wt": 1.15},
    "ADANIENT":   {"tok": "25",    "ts": "ADANIENT-EQ",   "wt": 1.10},
    "NTPC":       {"tok": "11630", "ts": "NTPC-EQ",       "wt": 1.10},
    "ONGC":       {"tok": "2475",  "ts": "ONGC-EQ",       "wt": 1.05},
    "POWERGRID":  {"tok": "14977", "ts": "POWERGRID-EQ",  "wt": 1.05},
    "TATASTEEL":  {"tok": "3499",  "ts": "TATASTEEL-EQ",  "wt": 1.00},
    "TECHM":      {"tok": "13538", "ts": "TECHM-EQ",      "wt": 0.95},
    "WIPRO":      {"tok": "3787",  "ts": "WIPRO-EQ",      "wt": 0.90},
    "ADANIPORTS": {"tok": "15083", "ts": "ADANIPORTS-EQ", "wt": 0.90},
    "COALINDIA":  {"tok": "20374", "ts": "COALINDIA-EQ",  "wt": 0.85},
    "JSWSTEEL":   {"tok": "11723", "ts": "JSWSTEEL-EQ",   "wt": 0.80},
    "BAJAJ-AUTO": {"tok": "16669", "ts": "BAJAJ-AUTO-EQ", "wt": 0.80},
    "BPCL":       {"tok": "526",   "ts": "BPCL-EQ",       "wt": 0.75},
    "HINDALCO":   {"tok": "1363",  "ts": "HINDALCO-EQ",   "wt": 0.75},
    "GRASIM":     {"tok": "1232",  "ts": "GRASIM-EQ",     "wt": 0.70},
    "DRREDDY":    {"tok": "881",   "ts": "DRREDDY-EQ",    "wt": 0.65},
    "CIPLA":      {"tok": "694",   "ts": "CIPLA-EQ",      "wt": 0.65},
    "EICHERMOT":  {"tok": "910",   "ts": "EICHERMOT-EQ",  "wt": 0.65},
    "BRITANNIA":  {"tok": "547",   "ts": "BRITANNIA-EQ",  "wt": 0.60},
    "INDUSINDBK": {"tok": "5258",  "ts": "INDUSINDBK-EQ", "wt": 0.60},
    "APOLLOHOSP": {"tok": "157",   "ts": "APOLLOHOSP-EQ", "wt": 0.55},
    "DIVISLAB":   {"tok": "10940", "ts": "DIVISLAB-EQ",   "wt": 0.55},
    "TATACONSUM": {"tok": "3432",  "ts": "TATACONSUM-EQ", "wt": 0.55},
    "HEROMOTOCO": {"tok": "1348",  "ts": "HEROMOTOCO-EQ", "wt": 0.50},
    "BEL":        {"tok": "383",   "ts": "BEL-EQ",        "wt": 0.50},
    "SHRIRAMFIN": {"tok": "4306",  "ts": "SHRIRAMFIN-EQ", "wt": 0.45},
    "LTIM":       {"tok": "17818", "ts": "LTM-EQ",        "wt": 0.45},
    "ZOMATO":     {"tok": "5097",  "ts": "ZOMATO-EQ",     "wt": 0.40},
}

def fetch_nifty50(obj, nifty_ltp):
    print("\n📋 Fetching Nifty 50 stocks...")
    stocks = []
    for sym, info in NIFTY50_STOCKS.items():
        for attempt in range(3):
            try:
                r = obj.ltpData("NSE", info["ts"], info["tok"])
                if r.get("status") and r.get("data"):
                    ltp   = float(r["data"].get("ltp",   0) or 0)
                    close = float(r["data"].get("close", ltp) or ltp)
                    chng  = round(ltp - close, 2)
                    pct   = round((chng / close) * 100, 2) if close else 0.0
                    wt    = info["wt"]
                    # Nifty point contribution = Nifty Level × (%Chg/100) × Weight%
                    contrib = round(nifty_ltp * (pct/100) * (wt/100), 2) if nifty_ltp else None
                    stocks.append({
                        "symbol": sym, "ltp": ltp, "chng": chng, "pChng": pct,
                        "weight": wt, "contrib": contrib
                    })
                    print(f"  ✅ {sym}: {ltp}  {pct:+.2f}%  wt={wt}%  contrib={contrib:+.2f}pts" if contrib is not None else f"  ✅ {sym}: {ltp}  {pct:+.2f}%")
                else:
                    msg = r.get("message","")
                    if "rate" in msg.lower() and attempt < 2:
                        time.sleep(3); continue
                    stocks.append({"symbol": sym, "ltp": None, "chng": None, "pChng": None,
                                   "weight": info["wt"], "contrib": None})
                    print(f"  ⚠️  {sym}: {msg}")
                break
            except Exception as e:
                if "rate" in str(e).lower() and attempt < 2:
                    time.sleep(5); continue
                stocks.append({"symbol": sym, "ltp": None, "chng": None, "pChng": None,
                               "weight": info["wt"], "contrib": None})
                print(f"  ❌ {sym}: {e}")
                break
        time.sleep(0.8)

    filled = sum(1 for s in stocks if s["ltp"] is not None)
    print(f"\n  Result: {filled}/{len(stocks)} stocks with data")
    return stocks

# ── Build Excel ──────────────────────────────────────────────
def build_excel(label, ist_dt, indices, stocks):
    wb = Workbook(); ws = wb.active
    ws.title = "Snapshot_" + label
    disp     = label[:2]+":"+label[2:] if len(label)==4 else label
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
        for col, val, fmt in [
            (2, d.get("ltp"),    "#,##0.00"),
            (3, d.get("chng"),   "+#,##0.00;-#,##0.00"),
            (4, pct and pct/100, "+0.00%;-0.00%"),
        ]:
            cell = ws.cell(r,col)
            if pct is not None:
                cell.value=val; cell.number_format=fmt
                cell.font=font(bold=True, color=C["pos_txt"] if pos else C["neg_txt"])
                cell.fill=fill(C["pos_bg"] if pos else C["neg_bg"])
            else:
                cell.value="N/A"; cell.font=font(color="9E9E9E"); cell.fill=fill(bg)
            cell.border=brd(); cell.alignment=aln("center")

    # All 50 stocks — now with WEIGHT and CONTRIBUTION (pts) columns
    sr = r + 3
    ws.merge_cells(f"A{sr}:G{sr}")
    sc(ws.cell(sr,1,f"📋  ALL NIFTY 50 STOCKS  [{len(stocks)} stocks]"),
       bg=C["hdr_dark"], fg=C["white"], bold=True, size=11, ha="center")
    sr += 1
    headers = ["#","SYMBOL","LTP (₹)","CHANGE (₹)","% CHANGE","WEIGHT %","CONTRIB (pts)"]
    for h, col in zip(headers, range(1,8)):
        sc(ws.cell(sr,col,h), bg="3949AB", fg=C["white"], bold=True, ha="center")

    for i, s in enumerate(stocks):
        sr += 1
        pos = (s["pChng"] or 0) >= 0
        bg  = (C["pos_bg"] if pos else C["neg_bg"]) if i%2==0 else \
              (C["pos_alt"] if pos else C["neg_alt"])
        tc  = C["pos_txt"] if pos else C["neg_txt"]
        ws.cell(sr,1,i+1).fill=fill(bg); ws.cell(sr,1).border=brd()
        ws.cell(sr,1).alignment=aln("center")
        ws.cell(sr,2,s["symbol"]).font=font(bold=True)
        ws.cell(sr,2).fill=fill(bg); ws.cell(sr,2).border=brd(); ws.cell(sr,2).alignment=aln()
        for col, val, fmt in [
            (3, s["ltp"],                      "#,##0.00"),
            (4, s["chng"],                     "+#,##0.00;-#,##0.00"),
            (5, s["pChng"] and s["pChng"]/100, "+0.00%;-0.00%"),
            (6, s.get("weight") and s["weight"]/100, "0.00%"),
            (7, s.get("contrib"),              "+0.00;-0.00"),
        ]:
            cell=ws.cell(sr,col); cell.value=val
            if val is not None and fmt: cell.number_format=fmt
            color = tc if col in (4,5,7) else "000000"
            cell.font=font(bold=(col in (4,5,7)), color=color)
            cell.fill=fill(bg); cell.border=brd(); cell.alignment=aln("center")

    # Top 7 positive/negative by % change — also show contribution
    valid = [s for s in stocks if s["pChng"] is not None]
    top7p = sorted(valid, key=lambda x: x["pChng"], reverse=True)[:7]
    top7n = sorted(valid, key=lambda x: x["pChng"])[:7]
    trow  = r + 4

    for top7, col, title, hbg, tc, b1, b2 in [
        (top7p,9, "🟢  TOP 7 POSITIVE",C["grn_hdr"],C["pos_txt"],C["pos_bg"],C["pos_alt"]),
        (top7n,15,"🔴  TOP 7 NEGATIVE",C["red_hdr"],C["neg_txt"],C["neg_bg"],C["neg_alt"]),
    ]:
        ws.merge_cells(start_row=trow, start_column=col, end_row=trow, end_column=col+4)
        sc(ws.cell(trow,col,title), bg=hbg, fg=C["white"], bold=True, size=11, ha="center")
        for h, dc in zip(["SYMBOL","LTP (₹)","CHG (₹)","% CHG","CONTRIB(pts)"], range(col,col+5)):
            sc(ws.cell(trow+1,dc,h), bg=hbg, fg=C["white"], bold=True, ha="center")
        for i, s in enumerate(top7):
            tr=trow+2+i; bg=b1 if i%2==0 else b2
            for dc, val, fmt in [
                (col,   s["symbol"],                   None),
                (col+1, s["ltp"],                      "#,##0.00"),
                (col+2, s["chng"],                     "+#,##0.00;-#,##0.00"),
                (col+3, s["pChng"] and s["pChng"]/100, "+0.00%;-0.00%"),
                (col+4, s.get("contrib"),               "+0.00;-0.00"),
            ]:
                cell=ws.cell(tr,dc,val)
                if val is not None and fmt: cell.number_format=fmt
                cell.font=font(bold=True, color=tc)
                cell.fill=fill(bg); cell.border=brd()
                cell.alignment=aln("center" if dc>col else "left")

    widths = {1:4,2:14,3:12,4:12,5:11,6:10,7:13,
              9:14,10:12,11:11,12:10,13:13,
              15:14,16:12,17:11,18:10,19:13}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A3"
    return wb

# ── Send Email ───────────────────────────────────────────────
def send_email(creds, xlsx_path, label, ist_dt, indices, stocks):
    disp    = label[:2]+":"+label[2:] if len(label)==4 else label
    subject = f"📊 NSE Snapshot {disp} IST — {ist_dt.strftime('%d %b %Y')}"
    valid   = [s for s in stocks if s["pChng"] is not None]
    # Top movers by Nifty contribution (more meaningful than just %change for index impact)
    contrib_valid = [s for s in stocks if s.get("contrib") is not None]
    top3p = sorted(contrib_valid, key=lambda x: x["contrib"], reverse=True)[:3]
    top3n = sorted(contrib_valid, key=lambda x: x["contrib"])[:3]

    idx_html = ""
    for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
        d = indices.get(name,{}); pct = d.get("pct"); ltp = d.get("ltp")
        if pct is not None:
            color = "#1b5e20" if pct>=0 else "#b71c1c"
            bg    = "#e8f5e9" if pct>=0 else "#ffebee"
            arrow = "▲" if pct>=0 else "▼"
            idx_html += (f'<tr style="background:{bg}"><td style="padding:7px 14px;font-weight:bold">{name}</td>'
                f'<td style="padding:7px;color:{color};font-weight:bold;text-align:center">{arrow} {pct:+.2f}%</td>'
                f'<td style="padding:7px;text-align:center">₹{ltp:,.2f}</td></tr>')
        else:
            idx_html += (f'<tr><td style="padding:7px 14px;font-weight:bold">{name}</td>'
                f'<td style="padding:7px;color:#9e9e9e;text-align:center">N/A</td>'
                f'<td style="padding:7px;text-align:center">—</td></tr>')

    def stock_rows(lst, color, bg):
        return "".join(
            f'<tr style="background:{bg}"><td style="padding:6px 12px;font-weight:bold">{s["symbol"]}</td>'
            f'<td style="padding:6px;color:{color};font-weight:bold;text-align:center">{s["pChng"]:+.2f}%</td>'
            f'<td style="padding:6px;color:{color};text-align:center">{s["contrib"]:+.2f} pts</td></tr>'
            for s in lst)

    note = ('<div style="background:#fff3e0;border-radius:6px;padding:10px 14px;margin-top:12px;'
            'font-size:13px;color:#e65100">⚠️ Market closed or data unavailable. '
            'Live data Mon–Fri 9:15 AM – 3:30 PM IST.</div>' if not valid else "")

    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:620px;margin:auto">
<div style="background:#0d47a1;color:white;padding:18px 22px;border-radius:10px 10px 0 0">
  <h2 style="margin:0;font-size:20px">📊 NSE Snapshot — {disp} IST</h2>
  <p style="margin:5px 0 0;opacity:.85;font-size:13px">{ist_dt.strftime('%d %b %Y  %H:%M:%S IST')}</p>
</div>
<div style="border:1px solid #ddd;border-top:none;padding:18px;border-radius:0 0 10px 10px">
  <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
    <tr style="background:#1976d2;color:white">
      <th style="padding:8px 14px;text-align:left">Index</th>
      <th style="padding:8px">% Change</th><th style="padding:8px">LTP</th>
    </tr>{idx_html}
  </table>
  <table width="100%" style="margin-top:16px;border-collapse:collapse"><tr valign="top">
    <td width="50%" style="padding-right:8px">
      <h3 style="color:#2e7d32;margin:0 0 8px">🏆 Top 3 Nifty Boosters</h3>
      <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
        {stock_rows(top3p,"#1b5e20","#e8f5e9") if top3p else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
      </table>
    </td>
    <td width="50%" style="padding-left:8px">
      <h3 style="color:#c62828;margin:0 0 8px">📉 Top 3 Nifty Draggers</h3>
      <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
        {stock_rows(top3n,"#b71c1c","#ffebee") if top3n else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
      </table>
    </td>
  </tr></table>
  {note}
  <div style="margin-top:16px;padding:12px 14px;background:#f5f5f5;border-radius:8px;font-size:13px">
    📎 Full Excel with all 50 stocks — includes weight% and Nifty point contribution<br>
    ☁️ Also saved to Google Drive → NSE Snapshots folder
  </div>
</div></body></html>"""

    msg = MIMEMultipart("alternative")
    msg["From"] = creds["gmail_sender"]
    msg["To"]   = creds["recipient"]
    msg["Subject"] = subject
    msg.attach(MIMEText(html, "html"))
    with open(xlsx_path, "rb") as f:
        part = MIMEBase("application","vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition","attachment",filename=os.path.basename(xlsx_path))
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(creds["gmail_sender"], creds["gmail_pass"])
        s.sendmail(creds["gmail_sender"], creds["recipient"], msg.as_string())
    print(f"✅ Email sent to {creds['recipient']}")

# ── Google Drive ─────────────────────────────────────────────
def upload_drive(creds, xlsx_path):
    folder_id = creds["drive_folder"]
    if not folder_id or len(folder_id) < 10:
        print("❌ GOOGLE_DRIVE_FOLDER_ID looks wrong — skipping Drive upload")
        return
    svc = build("drive","v3", credentials=get_sa_creds(
        creds["sa_json"], ["https://www.googleapis.com/auth/drive"]))
    try:
        svc.files().get(fileId=folder_id, fields="id,name").execute()
    except Exception as e:
        print(f"❌ Folder not found: {e}")
        return
    name = os.path.basename(xlsx_path)
    q    = f"name='{name}' and '{folder_id}' in parents and trashed=false"
    for f in svc.files().list(q=q, fields="files(id)").execute().get("files",[]):
        svc.files().delete(fileId=f["id"]).execute()
    meta  = {"name": name, "parents": [folder_id]}
    media = MediaFileUpload(xlsx_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    res = svc.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    print(f"✅ Drive upload: {name}")
    print(f"   {res.get('webViewLink','')}")

# ── Google Sheets ─────────────────────────────────────────────
def update_sheets(creds, label, ist_dt, indices, stocks):
    try:
        gc  = gspread.authorize(get_sa_creds(
            creds["sa_json"], ["https://www.googleapis.com/auth/spreadsheets"]))
        sh  = gc.open_by_key(creds["sheet_id"])
        tab = "Snapshot_" + label
        try: sh.del_worksheet(sh.worksheet(tab))
        except: pass
        ws  = sh.add_worksheet(title=tab, rows=120, cols=20)
        disp = label[:2]+":"+label[2:] if len(label)==4 else label
        rows = [
            [f"NSE Market Snapshot — {disp} IST"]+[""]*19,
            [f"Captured: {ist_dt.strftime('%d-%b-%Y %H:%M:%S IST')}"]+[""]*19,
            [""]*20,
            ["INDEX","LTP","CHANGE","% CHANGE",""]+["TOP 7 BOOSTERS","","","","",""]+["TOP 7 DRAGGERS","","","",""]
        ]
        for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
            d=indices.get(name,{}); p=d.get("pct")
            rows.append([name, d.get("ltp","N/A"), d.get("chng","N/A"),
                         f"{p:+.2f}%" if p is not None else "N/A",""]+[""]*15)
        rows.append([""]*20)
        rows.append(["#","SYMBOL","LTP","CHG","% CHG","WT%","CONTRIB","",
                     "SYMBOL","LTP","CHG","% CHG","CONTRIB","",
                     "SYMBOL","LTP","CHG","% CHG","CONTRIB"])

        contrib_valid = [s for s in stocks if s.get("contrib") is not None]
        top7p = sorted(contrib_valid, key=lambda x: x["contrib"], reverse=True)[:7]
        top7n = sorted(contrib_valid, key=lambda x: x["contrib"])[:7]

        for i, s in enumerate(stocks):
            p = s["pChng"]
            row = [i+1, s["symbol"], s.get("ltp",""),
                   f'{s["chng"]:+.2f}' if s["chng"] is not None else "",
                   f'{p:+.2f}%' if p is not None else "",
                   f'{s["weight"]:.2f}%' if s.get("weight") else "",
                   f'{s["contrib"]:+.2f}' if s.get("contrib") is not None else "", ""]
            row += ([top7p[i]["symbol"], top7p[i]["ltp"],
                     f'{top7p[i]["chng"]:+.2f}', f'{top7p[i]["pChng"]:+.2f}%',
                     f'{top7p[i]["contrib"]:+.2f}', ""]
                    if i < len(top7p) else [""]*6)
            row += ([top7n[i]["symbol"], top7n[i]["ltp"],
                     f'{top7n[i]["chng"]:+.2f}', f'{top7n[i]["pChng"]:+.2f}%',
                     f'{top7n[i]["contrib"]:+.2f}']
                    if i < len(top7n) else [""]*5)
            rows.append(row)
        ws.update("A1", rows)
        print(f"✅ Google Sheets updated: {tab}")
    except Exception as e:
        print(f"⚠️  Sheets update failed: {e}")

# ── MAIN ─────────────────────────────────────────────────────
def main():
    ist_dt = datetime.now(IST)
    label  = os.environ.get("MANUAL_LABEL","").strip() or ist_dt.strftime("%H%M")
    disp   = label[:2]+":"+label[2:] if len(label)==4 else label

    print(f"\n{'='*55}")
    print(f"  NSE Snapshot — {disp} IST | {ist_dt.strftime('%d-%b-%Y')}")
    print(f"{'='*55}")

    creds   = get_creds()
    obj     = angel_login(creds)
    indices = fetch_indices(obj)

    nifty_ltp = indices.get("NIFTY 50", {}).get("ltp") or 0
    print(f"\n  Nifty 50 LTP for contribution calc: {nifty_ltp}")

    stocks  = fetch_nifty50(obj, nifty_ltp)

    print("\n📁 Building Excel...")
    wb = build_excel(label, ist_dt, indices, stocks)
    os.makedirs("output", exist_ok=True)
    xlsx = f"output/NSE_{ist_dt.strftime('%Y-%m-%d')}_{label}.xlsx"
    wb.save(xlsx)
    print(f"  Saved: {xlsx}")

    print("\n📧 Sending email...")
    send_email(creds, xlsx, label, ist_dt, indices, stocks)

    print("\n☁️  Uploading to Drive...")
    upload_drive(creds, xlsx)

    print("\n📊 Updating Sheets...")
    update_sheets(creds, label, ist_dt, indices, stocks)

    print(f"\n✅ ALL DONE — {disp} IST\n")

if __name__ == "__main__":
    main()
