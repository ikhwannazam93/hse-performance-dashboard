import json, os, sys, urllib.request, urllib.parse, zipfile, io, re
from datetime import datetime, date, timezone
from openpyxl import load_workbook

SHARE_URL = os.environ["ONEDRIVE_EXCEL_URL"]
OUT = os.environ.get("OUTPUT_JSON", "data.json")

def download_xlsx(url):
    candidates = [url, url + ("&" if "?" in url else "?") + "download=1"]
    last = None
    for u in candidates:
        req = urllib.request.Request(u, headers={"User-Agent":"Mozilla/5.0"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                blob = r.read()
                ctype = (r.headers.get("Content-Type") or "").lower()
                # XLSX is a ZIP container and starts with PK.
                if blob[:2] == b"PK" and len(blob) > 1000:
                    return blob
                last = f"{u}: non-XLSX response ({ctype}, {len(blob)} bytes)"
        except Exception as e:
            last = f"{u}: {e}"
    raise RuntimeError("Unable to retrieve the shared Excel file. " + str(last))

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
