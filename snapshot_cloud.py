"""
snapshot_cloud.py  v4 — Fixed Angel One API calls
Uses ltpData() which works correctly both during and outside market hours
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

# ── Credentials from GitHub Secrets ─────────────────────────
def env(key):
    val = os.environ.get(key, "")
    if not val:
        raise ValueError(f"Missing secret: {key}")
    return val

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
    return Credentials.from_service_account_info(
        json.loads(sa_json_str), scopes=scopes)

# ── Excel styles ─────────────────────────────────────────────
C = {
    "title_bg": "0D47A1", "hdr_mid":  "1976D2", "hdr_dark": "283593",
    "pos_bg":   "E8F5E9", "pos_alt":  "F1F8E9", "pos_txt":  "1B5E20",
    "neg_bg":   "FFEBEE", "neg_alt":  "FCE4EC", "neg_txt":  "B71C1C",
    "idx_bg":   "E8EAF6", "idx_alt":  "FFFFFF",
    "grn_hdr":  "2E7D32", "red_hdr":  "C62828", "white":    "FFFFFF",
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

# ── Angel One login ──────────────────────────────────────────
def angel_login(creds):
    obj  = SmartConnect(api_key=creds["api_key"])
    totp = pyotp.TOTP(creds["totp_secret"]).now()
    data = obj.generateSession(creds["client_id"], creds["password"], totp)
    if data["status"] is False:
        raise Exception("Angel One login failed: " + str(data.get("message","")))
    print("✅ Angel One login successful")
    return obj

# ── Index tokens (NSE/BSE exchange + token) ──────────────────
INDEX_TOKENS = {
    "NIFTY 50":          ("NSE", "Nifty 50",           "99926000"),
    "SENSEX":            ("BSE", "SENSEX",              "99919000"),
    "BANK NIFTY":        ("NSE", "Nifty Bank",          "99926009"),
    "NIFTY IT":          ("NSE", "Nifty IT",            "99926011"),
    "NIFTY SMALLCAP 50": ("NSE", "Nifty Smallcap 50",  "99926074"),
}

# ── Nifty 50 stock tokens ────────────────────────────────────
# Format: symbol → (exchange_token, trading_symbol)
# Using hardcoded tokens to avoid searchScrip() which is slow & unreliable
NIFTY50_TOKENS = {
    "RELIANCE":   "2885",  "TCS":        "11536", "HDFCBANK":   "1333",
    "INFY":       "1594",  "ICICIBANK":  "4963",  "HINDUNILVR": "1394",
    "ITC":        "1660",  "KOTAKBANK":  "1922",  "LT":         "11483",
    "SBIN":       "3045",  "AXISBANK":   "5900",  "BAJFINANCE": "317",
    "BHARTIARTL": "10604", "M&M":        "2031",  "MARUTI":     "10999",
    "NESTLEIND":  "17963", "NTPC":       "11630", "ONGC":       "2475",
    "POWERGRID":  "14977", "SUNPHARMA":  "3351",  "TATAMOTORS": "3456",
    "TATASTEEL":  "3499",  "TECHM":      "13538", "TITAN":      "3506",
    "ULTRACEMCO": "11532", "WIPRO":      "3787",  "ADANIENT":   "25",
    "ADANIPORTS": "15083", "APOLLOHOSP": "157",   "ASIANPAINT": "236",
    "BAJAJFINSV": "16675", "BAJAJ-AUTO": "16669", "BEL":        "383",
    "BPCL":       "526",   "BRITANNIA":  "547",   "CIPLA":      "694",
    "COALINDIA":  "20374", "DIVISLAB":   "10940", "DRREDDY":    "881",
    "EICHERMOT":  "910",   "GRASIM":     "1232",  "HCLTECH":    "7229",
    "HEROMOTOCO": "1348",  "HINDALCO":   "1363",  "INDUSINDBK": "5258",
    "JSWSTEEL":   "11723", "LTIM":       "17818", "SHRIRAMFIN": "4306",
    "TATACONSUM": "3432",  "ZOMATO":     "5097",
}

def fetch_indices(obj):
    result = {}
    for name, (exch, trading_sym, token) in INDEX_TOKENS.items():
        try:
            r   = obj.ltpData(exch, trading_sym, token)
            if r["status"] and r.get("data"):
                ltp   = float(r["data"].get("ltp", 0))
                close = float(r["data"].get("close", ltp) or ltp)
                chng  = round(ltp - close, 2)
                pct   = round((chng / close) * 100, 2) if close else 0.0
                result[name] = {"ltp": ltp, "chng": chng, "pct": pct}
                print(f"  ✅ {name}: LTP={ltp}  {pct:+.2f}%")
            else:
                print(f"  ⚠️  {name}: no data returned — {r.get('message','')}")
                result[name] = {"ltp": None, "chng": None, "pct": None}
        except Exception as e:
            print(f"  ⚠️  {name}: {e}")
            result[name] = {"ltp": None, "chng": None, "pct": None}
        time.sleep(0.3)
    return result

def fetch_nifty50(obj):
    stocks = []
    for sym, token in NIFTY50_TOKENS.items():
        try:
            r = obj.ltpData("NSE", sym, token)
            if r["status"] and r.get("data"):
                ltp   = float(r["data"].get("ltp",   0))
                close = float(r["data"].get("close", ltp) or ltp)
                chng  = round(ltp - close, 2)
                pct   = round((chng / close) * 100, 2) if close else 0.0
                stocks.append({"symbol": sym, "ltp": ltp, "chng": chng, "pChng": pct})
                print(f"  ✅ {sym}: {ltp}  {pct:+.2f}%")
            else:
                print(f"  ⚠️  {sym}: no data — {r.get('message','')}")
                stocks.append({"symbol": sym, "ltp": None, "chng": None, "pChng": None})
        except Exception as e:
            print(f"  ⚠️  {sym}: {e}")
            stocks.append({"symbol": sym, "ltp": None, "chng": None, "pChng": None})
        time.sleep(0.2)
    return stocks

# ── Build Excel workbook ─────────────────────────────────────
def build_excel(label, ist_dt, indices, stocks):
    wb = Workbook(); ws = wb.active
    ws.title = "Snapshot_" + label
    disp     = label[:2]+":"+label[2:] if len(label)==4 else label
    time_str = ist_dt.strftime("%d-%b-%Y %H:%M:%S")

    # Title row
    ws.merge_cells("A1:O1")
    c = ws["A1"]; c.value = f"NSE Market Snapshot  —  {disp} IST"
    sc(c, bg=C["title_bg"], fg=C["white"], bold=True, size=14, ha="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:O2")
    ws["A2"].value = f"Captured At (IST):  {time_str}"
    ws["A2"].fill = fill("E3F2FD")
    ws["A2"].font = font(size=10, color="333333")
    ws["A2"].alignment = aln("center")

    # Index summary
    r = 4
    ws.merge_cells(f"A{r}:D{r}")
    sc(ws.cell(r,1,"📊  INDEX SUMMARY"), bg=C["hdr_dark"], fg=C["white"], bold=True, size=11, ha="center")
    r += 1
    for h, col in zip(["INDEX","LTP (₹)","CHANGE (₹)","% CHANGE"], range(1,5)):
        sc(ws.cell(r,col,h), bg=C["hdr_mid"], fg=C["white"], bold=True, ha="center")

    for i, name in enumerate(["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]):
        r += 1; d = indices.get(name,{}); pct = d.get("pct")
        bg  = C["idx_bg"] if i%2==0 else C["idx_alt"]
        pos = pct is not None and pct >= 0
        ws.cell(r,1,name).font = font(bold=True)
        ws.cell(r,1).fill=fill(bg); ws.cell(r,1).border=brd(); ws.cell(r,1).alignment=aln()
        for col, val, fmt in [
            (2, d.get("ltp"),       "#,##0.00"),
            (3, d.get("chng"),      "+#,##0.00;-#,##0.00"),
            (4, pct and pct/100,    "+0.00%;-0.00%"),
        ]:
            cell = ws.cell(r,col)
            if pct is not None:
                cell.value=val; cell.number_format=fmt
                cell.font=font(bold=True, color=C["pos_txt"] if pos else C["neg_txt"])
                cell.fill=fill(C["pos_bg"] if pos else C["neg_bg"])
            else:
                cell.value="N/A"; cell.font=font(color="9E9E9E"); cell.fill=fill(bg)
            cell.border=brd(); cell.alignment=aln("center")

    # All 50 stocks
    sr = r + 3
    ws.merge_cells(f"A{sr}:E{sr}")
    sc(ws.cell(sr,1,f"📋  ALL NIFTY 50 STOCKS  [{len(stocks)} stocks]"),
       bg=C["hdr_dark"], fg=C["white"], bold=True, size=11, ha="center")
    sr += 1
    for h, col in zip(["#","SYMBOL","LTP (₹)","CHANGE (₹)","% CHANGE"], range(1,6)):
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
        ]:
            cell=ws.cell(sr,col); cell.value=val
            if val is not None and fmt: cell.number_format=fmt
            cell.font=font(bold=(col>2), color=tc if col>2 else "000000")
            cell.fill=fill(bg); cell.border=brd(); cell.alignment=aln("center")

    # Top 7 positive & negative
    valid = [s for s in stocks if s["pChng"] is not None]
    top7p = sorted(valid, key=lambda x: x["pChng"], reverse=True)[:7]
    top7n = sorted(valid, key=lambda x: x["pChng"])[:7]
    trow  = r + 4

    for top7, col, title, hbg, tc, b1, b2 in [
        (top7p, 7,  "🟢  TOP 7 POSITIVE", C["grn_hdr"], C["pos_txt"], C["pos_bg"], C["pos_alt"]),
        (top7n, 12, "🔴  TOP 7 NEGATIVE", C["red_hdr"], C["neg_txt"], C["neg_bg"], C["neg_alt"]),
    ]:
        ws.merge_cells(start_row=trow, start_column=col, end_row=trow, end_column=col+3)
        sc(ws.cell(trow,col,title), bg=hbg, fg=C["white"], bold=True, size=11, ha="center")
        for h, dc in zip(["SYMBOL","LTP (₹)","CHG (₹)","% CHG"], range(col,col+4)):
            sc(ws.cell(trow+1,dc,h), bg=hbg, fg=C["white"], bold=True, ha="center")
        for i, s in enumerate(top7):
            tr=trow+2+i; bg=b1 if i%2==0 else b2
            for dc, val, fmt in [
                (col,   s["symbol"],                   None),
                (col+1, s["ltp"],                      "#,##0.00"),
                (col+2, s["chng"],                     "+#,##0.00;-#,##0.00"),
                (col+3, s["pChng"] and s["pChng"]/100, "+0.00%;-0.00%"),
            ]:
                cell=ws.cell(tr,dc,val)
                if val is not None and fmt: cell.number_format=fmt
                cell.font=font(bold=True, color=tc)
                cell.fill=fill(bg); cell.border=brd()
                cell.alignment=aln("center" if dc>col else "left")

    for col, w in {1:4,2:15,3:13,4:13,5:11,7:15,8:13,9:13,10:11,12:15,13:13,14:13,15:11}.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A3"
    return wb

# ── Send Email ───────────────────────────────────────────────
def send_email(creds, xlsx_path, label, ist_dt, indices, stocks):
    disp     = label[:2]+":"+label[2:] if len(label)==4 else label
    subject  = f"📊 NSE Snapshot {disp} IST — {ist_dt.strftime('%d %b %Y')}"

    valid = [s for s in stocks if s["pChng"] is not None]
    top3p = sorted(valid, key=lambda x: x["pChng"], reverse=True)[:3]
    top3n = sorted(valid, key=lambda x: x["pChng"])[:3]

    # Index rows HTML
    idx_html = ""
    for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
        d = indices.get(name,{}); pct = d.get("pct"); ltp = d.get("ltp")
        if pct is not None:
            color = "#1b5e20" if pct>=0 else "#b71c1c"
            bg    = "#e8f5e9" if pct>=0 else "#ffebee"
            arrow = "▲" if pct>=0 else "▼"
            idx_html += (f'<tr style="background:{bg}">'
                f'<td style="padding:7px 14px;font-weight:bold">{name}</td>'
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
            f'<td style="padding:6px;color:{color};text-align:center">₹{s["chng"]:+.2f}</td></tr>'
            for s in lst)

    data_note = ""
    if not valid:
        data_note = '<div style="background:#fff3e0;border-radius:6px;padding:10px 14px;margin-top:12px;font-size:13px;color:#e65100">⚠️ No live data — market may be closed. Excel shows previous closing values.</div>'

    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:620px;margin:auto">
<div style="background:#0d47a1;color:white;padding:18px 22px;border-radius:10px 10px 0 0">
  <h2 style="margin:0;font-size:20px">📊 NSE Snapshot — {disp} IST</h2>
  <p style="margin:5px 0 0;opacity:.85;font-size:13px">{ist_dt.strftime('%d %b %Y  %H:%M:%S IST')}</p>
</div>
<div style="border:1px solid #ddd;border-top:none;padding:18px;border-radius:0 0 10px 10px">
  <h3 style="color:#0d47a1;margin:0 0 10px">Index Performance</h3>
  <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
    <tr style="background:#1976d2;color:white">
      <th style="padding:8px 14px;text-align:left">Index</th>
      <th style="padding:8px">% Change</th>
      <th style="padding:8px">LTP</th>
    </tr>
    {idx_html}
  </table>
  <table width="100%" style="margin-top:16px;border-collapse:collapse">
    <tr valign="top">
      <td width="50%" style="padding-right:8px">
        <h3 style="color:#2e7d32;margin:0 0 8px">🏆 Top 3 Gainers</h3>
        <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
          {stock_rows(top3p,"#1b5e20","#e8f5e9") if top3p else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
        </table>
      </td>
      <td width="50%" style="padding-left:8px">
        <h3 style="color:#c62828;margin:0 0 8px">📉 Top 3 Losers</h3>
        <table width="100%" cellspacing="0" style="border-collapse:collapse;border:1px solid #e0e0e0">
          {stock_rows(top3n,"#b71c1c","#ffebee") if top3n else '<tr><td style="padding:8px;color:#9e9e9e">No data</td></tr>'}
        </table>
      </td>
    </tr>
  </table>
  {data_note}
  <div style="margin-top:16px;padding:12px 14px;background:#f5f5f5;border-radius:8px;font-size:13px">
    📎 Full Excel with all 50 stocks attached<br>
    ☁️ Also saved to your Google Drive → NSE Snapshots folder
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

# ── Google Drive upload ──────────────────────────────────────
def upload_drive(creds, xlsx_path):
    svc = build("drive","v3", credentials=get_sa_creds(
        creds["sa_json"], ["https://www.googleapis.com/auth/drive"]))
    name = os.path.basename(xlsx_path)
    # Delete existing file with same name
    q = f"name='{name}' and '{creds['drive_folder']}' in parents and trashed=false"
    for f in svc.files().list(q=q, fields="files(id)").execute().get("files",[]):
        svc.files().delete(fileId=f["id"]).execute()
    meta  = {"name": name, "parents": [creds["drive_folder"]]}
    media = MediaFileUpload(xlsx_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    res = svc.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    print(f"✅ Drive: {name}")
    print(f"   {res.get('webViewLink','')}")

# ── Google Sheets update ─────────────────────────────────────
def update_sheets(creds, label, ist_dt, indices, stocks):
    gc  = gspread.authorize(get_sa_creds(
        creds["sa_json"], ["https://www.googleapis.com/auth/spreadsheets"]))
    sh  = gc.open_by_key(creds["sheet_id"])
    tab = "Snapshot_" + label
    try: sh.del_worksheet(sh.worksheet(tab))
    except: pass
    ws  = sh.add_worksheet(title=tab, rows=120, cols=16)

    disp = label[:2]+":"+label[2:] if len(label)==4 else label
    rows = [
        [f"NSE Market Snapshot — {disp} IST"] + [""]*14,
        [f"Captured: {ist_dt.strftime('%d-%b-%Y %H:%M:%S IST')}"] + [""]*14,
        [""]*15,
        ["INDEX","LTP","CHANGE","% CHANGE",""] +
        ["TOP 7 POSITIVE","","","",""] +
        ["TOP 7 NEGATIVE","","",""]
    ]
    for name in ["NIFTY 50","SENSEX","BANK NIFTY","NIFTY IT","NIFTY SMALLCAP 50"]:
        d = indices.get(name,{}); p = d.get("pct")
        rows.append([name, d.get("ltp","N/A"), d.get("chng","N/A"),
                     f"{p:+.2f}%" if p is not None else "N/A", ""] + [""]*10)

    rows.append([""]*15)
    rows.append(["#","SYMBOL","LTP","CHG","% CHG","",
                 "SYMBOL","LTP","CHG","% CHG","",
                 "SYMBOL","LTP","CHG","% CHG"])

    valid = [s for s in stocks if s["pChng"] is not None]
    top7p = sorted(valid, key=lambda x: x["pChng"], reverse=True)[:7]
    top7n = sorted(valid, key=lambda x: x["pChng"])[:7]

    for i, s in enumerate(stocks):
        p = s["pChng"]
        row = [i+1, s["symbol"], s.get("ltp",""),
               f'{s["chng"]:+.2f}' if s["chng"] is not None else "",
               f'{p:+.2f}%' if p is not None else "", ""]
        row += ([top7p[i]["symbol"], top7p[i]["ltp"],
                 f'{top7p[i]["chng"]:+.2f}', f'{top7p[i]["pChng"]:+.2f}%', ""]
                if i < len(top7p) else [""]*5)
        row += ([top7n[i]["symbol"], top7n[i]["ltp"],
                 f'{top7n[i]["chng"]:+.2f}', f'{top7n[i]["pChng"]:+.2f}%']
                if i < len(top7n) else [""]*4)
        rows.append(row)

    ws.update("A1", rows)
    print(f"✅ Google Sheets updated: {tab}")

# ── MAIN ─────────────────────────────────────────────────────
def main():
    ist_dt = datetime.now(IST)
    label  = os.environ.get("MANUAL_LABEL","").strip() or ist_dt.strftime("%H%M")
    disp   = label[:2]+":"+label[2:] if len(label)==4 else label

    print(f"\n{'='*55}")
    print(f"  NSE Snapshot — {disp} IST | {ist_dt.strftime('%d-%b-%Y')}")
    print(f"  Running on GitHub Actions — no PC needed")
    print(f"{'='*55}\n")

    creds = get_creds()

    print("🔑 Logging in to Angel One...")
    obj = angel_login(creds)

    print("\n📊 Fetching indices...")
    indices = fetch_indices(obj)

    print("\n📋 Fetching all 50 Nifty stocks...")
    stocks = fetch_nifty50(obj)

    filled = sum(1 for s in stocks if s["ltp"] is not None)
    print(f"\n  Data received for {filled}/{len(stocks)} stocks")

    print("\n📁 Building Excel file...")
    wb = build_excel(label, ist_dt, indices, stocks)
    os.makedirs("output", exist_ok=True)
    xlsx = f"output/NSE_{ist_dt.strftime('%Y-%m-%d')}_{label}.xlsx"
    wb.save(xlsx)
    print(f"  Saved: {xlsx}")

    print("\n📧 Sending email...")
    send_email(creds, xlsx, label, ist_dt, indices, stocks)

    print("\n☁️  Uploading to Google Drive...")
    upload_drive(creds, xlsx)

    print("\n📊 Updating Google Sheets...")
    update_sheets(creds, label, ist_dt, indices, stocks)

    print(f"\n✅ ALL DONE — {disp} IST\n")

if __name__ == "__main__":
    main()
