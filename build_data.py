import json, os, sys, urllib.request, urllib.parse, zipfile, io, re
from datetime import datetime, date, timezone
from openpyxl import load_workbook

SHARE_URL = os.environ["ONEDRIVE_EXCEL_URL"]
OUT = os.environ.get("OUTPUT_JSON", "data.json")

def download_xlsx(url):
    """Download the shared workbook from a OneDrive/SharePoint sharing URL.

    OneDrive for Business share URLs can return the Office viewer HTML page when
    the original `?e=...` query is retained.  A common direct-download form is
    the same share path with its query replaced by `?download=1`.  We try that
    first, then the original link variants, and finally inspect an HTML response
    for a downloadable URL exposed by the sharing page.
    """
    parsed = urllib.parse.urlsplit(url)
    base = urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, "", ""))

    candidates = []
    # Most important for OneDrive for Business / SharePoint sharing links:
    # remove the e= tracking query and use ?download=1.
    candidates.append(("share-path download", base + "?download=1"))
    candidates.append(("share-path download with original e", base + "?download=1" + ("&e=" + urllib.parse.parse_qs(parsed.query).get("e", [""])[0] if urllib.parse.parse_qs(parsed.query).get("e") else "")))
    candidates.append(("original link", url))
    candidates.append(("original link + download", url + ("&" if "?" in url else "?") + "download=1"))

    tried = []
    html_blobs = []

    def fetch(u):
        req = urllib.request.Request(
            u,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/153 Safari/537.36",
                "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,text/html;q=0.9,*/*;q=0.8",
            },
        )
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read(), (r.headers.get("Content-Type") or "").lower(), r.geturl()

    def is_xlsx(blob):
        # XLSX is a ZIP container and normally starts with PK. Validate that it
        # is actually an OOXML workbook rather than just another ZIP response.
        if len(blob) < 1000 or blob[:2] != b"PK":
            return False
        try:
            with zipfile.ZipFile(io.BytesIO(blob)) as z:
                names = set(z.namelist())
                return "[Content_Types].xml" in names and "xl/workbook.xml" in names
        except Exception:
            return False

    def extract_urls_from_html(text, page_url):
        found = []
        # Decode common JSON/HTML escaping first.
        t = text.replace('\\u0026', '&').replace('\\/', '/').replace('\\"', '"').replace('&amp;', '&')
        patterns = [
            r'https?://[^"\'<>\\s]+',
            r'(?:(?:https?:)?//)[^"\'<>\\s]+',
        ]
        for pat in patterns:
            for m in re.findall(pat, t, flags=re.I):
                u = m
                if u.startswith('//'):
                    u = parsed.scheme + ':' + u
                u = u.rstrip('\\\\,;)]}')
                low = u.lower()
                if any(k in low for k in ("download", "download.aspx", "download=1", "_layouts/15/download", ".xlsx")):
                    found.append(u)
        # Meta refresh and href/src values can contain relative download routes.
        for m in re.findall(r'(?:href|src|content)=["\']([^"\']+)["\']', t, flags=re.I):
            u = urllib.parse.urljoin(page_url, m.replace('\\u0026','&'))
            low = u.lower()
            if any(k in low for k in ("download", ".xlsx")):
                found.append(u)
        # Preserve order and remove duplicates.
        out=[]
        seen=set()
        for u in found:
            if u not in seen:
                seen.add(u); out.append(u)
        return out

    for label, u in candidates:
        try:
            blob, ctype, final_url = fetch(u)
            if is_xlsx(blob):
                return blob
            tried.append(f"{label}: non-XLSX response ({ctype}, {len(blob)} bytes, final={final_url})")
            if "text/html" in ctype or blob.lstrip().lower().startswith((b"<!doctype", b"<html", b"<head")):
                try:
                    html_blobs.append((blob.decode("utf-8", errors="ignore"), final_url))
                except Exception:
                    pass
        except Exception as e:
            tried.append(f"{label}: {e}")

    # Some SharePoint/OneDrive viewer pages expose a short-lived download URL
    # in their HTML/JSON. Try those URLs as a last server-side fallback.
    for html, page_url in html_blobs:
        for u in extract_urls_from_html(html, page_url):
            try:
                blob, ctype, final_url = fetch(u)
                if is_xlsx(blob):
                    return blob
                tried.append(f"embedded download URL: non-XLSX response ({ctype}, {len(blob)} bytes, final={final_url})")
            except Exception as e:
                tried.append(f"embedded download URL: {e}")

    detail = "\n".join(tried[-8:])
    raise RuntimeError("Unable to retrieve the shared Excel file as XLSX. Tried OneDrive direct-download forms and viewer download URLs.\n" + detail)

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
