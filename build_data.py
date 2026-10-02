import json, os, sys, urllib.request, urllib.parse, zipfile, io, re
from datetime import datetime, date, timezone
from openpyxl import load_workbook

SHARE_URL = os.environ["ONEDRIVE_EXCEL_URL"]
OUT = os.environ.get("OUTPUT_JSON", "data.json")

def download_xlsx(url):
    """Download the workbook using a real Chromium browser first.

    The OneDrive/SharePoint anonymous link opens correctly in a normal browser,
    but GitHub's plain HTTP request is redirected to login.microsoftonline.com.
    A headless browser follows the same anonymous sharing flow and can capture
    the actual XLSX download. HTTP fallbacks are retained for other link types.
    """
    import urllib.request, urllib.parse, zipfile, io, re

    def is_xlsx(blob):
        if len(blob) < 1000 or blob[:2] != b"PK":
            return False
        try:
            with zipfile.ZipFile(io.BytesIO(blob)) as z:
                names = set(z.namelist())
                return "[Content_Types].xml" in names and "xl/workbook.xml" in names
        except Exception:
            return False

    browser_error = None
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(7000)

            # Anonymous OneDrive/Excel pages normally expose a Download button.
            # Try several accessible-text variants because Microsoft changes the
            # viewer UI labels between deployments.
            download = None
            selectors = [
                'a:has-text("Download")',
                'button:has-text("Download")',
                '[role="button"]:has-text("Download")',
                'text=Download',
            ]
            for sel in selectors:
                try:
                    loc = page.locator(sel).first
                    if loc.count() and loc.is_visible(timeout=1500):
                        with page.expect_download(timeout=30000) as dl_info:
                            loc.click(timeout=5000)
                        download = dl_info.value
                        break
                except Exception:
                    pass

            # Excel/Office viewers may put download under File.
            if download is None:
                for sel in ['button:has-text("File")', '[role="button"]:has-text("File")', 'text=File']:
                    try:
                        loc = page.locator(sel).first
                        if loc.count() and loc.is_visible(timeout=1500):
                            loc.click(timeout=5000)
                            page.wait_for_timeout(1000)
                            break
                    except Exception:
                        pass
                for sel in [
                    'text=Download a Copy',
                    'text=Download',
                    'button:has-text("Download a Copy")',
                    '[role="menuitem"]:has-text("Download")',
                ]:
                    try:
                        loc = page.locator(sel).first
                        if loc.count() and loc.is_visible(timeout=2000):
                            with page.expect_download(timeout=30000) as dl_info:
                                loc.click(timeout=5000)
                            download = dl_info.value
                            break
                    except Exception:
                        pass

            if download is not None:
                path = download.path()
                if path:
                    blob = open(path, "rb").read()
                    if is_xlsx(blob):
                        browser.close()
                        return blob
                    browser_error = f"browser download was not XLSX ({len(blob)} bytes)"
                else:
                    browser_error = "browser download had no local path"
            else:
                final_url = page.url
                visible = page.locator("body").inner_text(timeout=10000)[:1200]
                browser_error = f"browser could not trigger download; final={final_url}; page={visible!r}"
            browser.close()
    except Exception as e:
        browser_error = f"browser method failed: {type(e).__name__}: {e}"

    # Retain direct-download fallbacks for links that do not require browser UI.
    parsed = urllib.parse.urlsplit(url)
    base = urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, "", ""))
    e = urllib.parse.parse_qs(parsed.query).get("e", [""])[0]
    candidates = [
        ("share-path download", base + "?download=1"),
        ("share-path download with original e", base + "?download=1" + ("&e=" + e if e else "")),
        ("original link", url),
        ("original link + download", url + ("&" if "?" in url else "?") + "download=1"),
    ]
    tried = [f"browser: {browser_error}"]

    for label, u in candidates:
        try:
            req = urllib.request.Request(u, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/153 Safari/537.36",
                "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,text/html;q=0.9,*/*;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=60) as r:
                blob = r.read()
                ctype = (r.headers.get("Content-Type") or "").lower()
                final_url = r.geturl()
            if is_xlsx(blob):
                return blob
            tried.append(f"{label}: non-XLSX response ({ctype}, {len(blob)} bytes, final={final_url})")
        except Exception as e2:
            tried.append(f"{label}: {e2}")

    raise RuntimeError("Unable to retrieve the shared Excel file as XLSX.\n" + "\n".join(tried[-6:]))

def norm(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "").strip()).upper()

def excel_date(v):
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int,float)):
        # Excel serial date, 1899-12-30 system.
        from datetime import timedelta
        return (datetime(1899,12,30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        a,b,y = int(m.group(1)),int(m.group(2)),m.group(3)
        if a > 12:
            day,month = a,b
        else:
            month,day = a,b
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{y}-{month:02d}-{day:02d}"
    return s[:10]

def num(v):
    if isinstance(v,(int,float)) and not isinstance(v,bool):
        return float(v)
    try:
        return float(str(v if v is not None else "").replace(",",""))
    except:
        return 0

def rows_from_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    hr = None
    for i,row in enumerate(rows):
        vals = [norm(x) for x in row]
        if "DATE" in vals or "DATE REPORTED" in vals:
            hr = i
            break
    if hr is None:
        return []
    headers = [norm(x) if norm(x) else f"__EMPTY_{i}" for i,x in enumerate(rows[hr])]
    out=[]
    for row in rows[hr+1:]:
        obj={}
        has=False
        for i,h in enumerate(headers):
            v=row[i] if i < len(row) else ""
            obj[h]=v
            if v not in ("",None):
                has=True
        if has:
            out.append(obj)
    return out

def val(row,names):
    names = names if isinstance(names,list) else [names]
    for n in names:
        if norm(n) in row:
            return row[norm(n)]
    return ""

def clean_sheet_name(s):
    return str(s)

wb = load_workbook(io.BytesIO(download_xlsx(SHARE_URL)), data_only=True, read_only=True)

excluded={"LAGGING","LEADING","BE SAFE CARD LOG","TRAINING"}
manpower=[]
for ws in wb.worksheets:
    if norm(ws.title) in excluded:
        continue
    raw=rows_from_sheet(ws)
    if not raw:
        continue
    first=raw[0]
    if "DATE" not in first or not ("MANHOUR" in first or "MANDAY" in first):
        continue
    rows=[]
    for r in raw:
        d=excel_date(val(r,"DATE"))
        if not d:
            continue
        rows.append({
            "date":d,
            "week":str(val(r,"WEEK") or ""),
            "company":str(val(r,"COMPANY") or ""),
            "category":str(val(r,"CATEGORY") or ""),
            "manday":num(val(r,"MANDAY")),
            "manhour":num(val(r,"MANHOUR"))
        })
    if rows:
        manpower.append({"sheet":clean_sheet_name(ws.title),"rows":rows})

bsc_raw=rows_from_sheet(wb["Be Safe Card Log"]) if "Be Safe Card Log" in wb.sheetnames else []
bsc=[]
for i,r in enumerate(bsc_raw,1):
    bsc.append({
        "no":num(val(r,"NO")) or i,
        "bscNo":str(val(r,["BSC NO","BSC NO."]) or ""),
        "type":str(val(r,["BSC TYPE","TYPE"]) or ""),
        "hazardRaw":str(val(r,"HAZARD TYPE RAW") or ""),
        "hazard":str(val(r,"HAZARD TYPE") or ""),
        "dateReported":excel_date(val(r,"DATE REPORTED")),
        "week":str(val(r,"WEEK") or ""),
        "dateObserved":excel_date(val(r,"DATE OBSERVED")),
        "observer":str(val(r,["OBSERVED BY","OBSERVER"]) or ""),
        "company":str(val(r,"COMPANY") or ""),
        "position":str(val(r,"POSITION") or ""),
        "description":str(val(r,"WHAT AND WHERE HAZARD/ACT OBSERVED AND DISCUSSED WITH WORKER/SUPERVISOR/MANAGER?") or ""),
        "status":str(val(r,"STATUS") or "")
    })
bsc=[r for r in bsc if r["dateReported"] or r["dateObserved"]]

tr_topics=["SIC","PTW NOVADE","EMERGENCY RESPONSE","FIRE WATCH","CHEMICAL HANDLING","WASTE MANAGEMENT","ELECTRICAL SAFETY","PPE AWARENESS","HAND TOOLS & POWER TOOLS SAFETY"]
tr_raw=rows_from_sheet(wb["Training"]) if "Training" in wb.sheetnames else []
training=[]
for r in tr_raw:
    d=excel_date(val(r,"DATE"))
    if d:
        o={"date":d,"week":str(val(r,"WEEK") or "")}
        for t in tr_topics: o[t]=num(val(r,t))
        training.append(o)

lead_topics=["HSE INDUCTION","TOOLBOX MEETING","HSE COORDINATION MEETING","HSE COMMITTEE MEETING","EMERGENCY DRILL","WORKPLACE INSPECTION","HSE WALKABOUT","WEEKLY HOUSEKEEPING","BESAFE CARD","HSE CAMPAIGN","REWARDS & RECOGNITION","CONSEQUENCE MANAGEMENT"]
lead_raw=rows_from_sheet(wb["Leading"]) if "Leading" in wb.sheetnames else []
leading=[]
for r in lead_raw:
    d=excel_date(val(r,"DATE"))
    if d:
        o={"date":d,"week":str(val(r,"WEEK") or "")}
        for t in lead_topics: o[t]=num(val(r,t))
        leading.append(o)

lag_topics=["FATALITY","LTI","RWC","MTC","FAC","DANGEROUS OCCURANCE","PROPERTY DAMAGE","NEAR MISS","FIRE INCIDENT","REGULATORY NON COMPLIANCE","LoPC"]
lag_raw=rows_from_sheet(wb["Lagging"]) if "Lagging" in wb.sheetnames else []
lagging=[]
for r in lag_raw:
    d=excel_date(val(r,"DATE"))
    if d:
        o={"date":d,"week":str(val(r,"WEEK") or "")}
        for t in lag_topics: o[t]=num(val(r,t))
        lagging.append(o)

payload={
    "generated_at":datetime.now(timezone.utc).isoformat(),
    "source":"OneDrive shared Excel",
    "workbook":"HSE Dashboard Rev 0(3).xlsx",
    "manpower":manpower,
    "bsc":bsc,
    "training":training,
    "trainingTopics":tr_topics,
    "leading":leading,
    "leadingTopics":lead_topics,
    "lagging":lagging,
    "laggingTopics":lag_topics
}
with open(OUT,"w",encoding="utf-8") as f:
    json.dump(payload,f,ensure_ascii=False,separators=(",",":"))
print("Generated",OUT)
print("Manpower rows:",sum(len(x["rows"]) for x in manpower))
print("BSC:",len(bsc),"Training:",len(training),"Leading:",len(leading),"Lagging:",len(lagging))
