import json, os, sys, urllib.request, urllib.parse, zipfile, io, re
from datetime import datetime, date, timezone
from openpyxl import load_workbook

SHARE_URL = os.environ.get("ONEDRIVE_EXCEL_URL", "")
OUT = os.environ.get("OUTPUT_JSON", "data.json")
ONEDRIVE_USER_UPN = os.environ.get("ONEDRIVE_USER_UPN", "hse@databasehub.onmicrosoft.com")
ONEDRIVE_FILE_NAME = os.environ.get("ONEDRIVE_FILE_NAME", "HSE Dashboard Rev 0(3).xlsx")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

def _graph_json(url, token):
    req = urllib.request.Request(url, headers={
        "Authorization": "Bearer " + token,
        "Accept": "application/json",
        "User-Agent": "HSE-Performance-Dashboard/Graph"
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"Microsoft Graph HTTP {e.code}: {body}") from e

def _get_graph_token():
    tenant = os.environ.get("AZURE_TENANT_ID", "").strip()
    client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
    if not tenant or not client_id or not client_secret:
        raise RuntimeError("Missing AZURE_TENANT_ID, AZURE_CLIENT_ID, or AZURE_CLIENT_SECRET GitHub secret.")

    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode("utf-8")
    req = urllib.request.Request(
        TOKEN_URL.format(tenant=urllib.parse.quote(tenant, safe="")),
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            result = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"Microsoft identity token request failed (HTTP {e.code}): {body}") from e
    token = result.get("access_token")
    if not token:
        raise RuntimeError("Microsoft identity token response did not contain access_token.")
    return token

def download_xlsx(_url=None):
    """Download the dashboard workbook from the user's OneDrive using Microsoft Graph app-only authentication.

    The previous anonymous-sharing/browser approach is intentionally removed.
    Graph searches the specified user's OneDrive for the workbook, then downloads
    the DriveItem content using the Files.Read.All application permission.
    """
    token = _get_graph_token()
    upn = ONEDRIVE_USER_UPN.strip()
    filename = ONEDRIVE_FILE_NAME.strip()
    if not upn or not filename:
        raise RuntimeError("ONEDRIVE_USER_UPN and ONEDRIVE_FILE_NAME must be set.")

    # Microsoft Graph supports searching a user's OneDrive by filename with
    # Files.Read.All application permission.
    user_part = urllib.parse.quote(upn, safe="")
    q_part = urllib.parse.quote("'" + filename + "'", safe="")
    search_url = f"{GRAPH_BASE}/users/{user_part}/drive/root/search(q={q_part})"
    result = _graph_json(search_url, token)
    matches = result.get("value", [])

    # Exact filename match first. If Graph's search returns no exact match,
    # retry with the filename stem so minor naming differences such as (3) do
    # not break the dashboard.
    exact = [m for m in matches if str(m.get("name", "")).lower() == filename.lower() and m.get("file")]
    if not exact:
        stem = filename.rsplit(".", 1)[0]
        q2 = urllib.parse.quote("'" + stem + "'", safe="")
        result2 = _graph_json(f"{GRAPH_BASE}/users/{user_part}/drive/root/search(q={q2})", token)
        matches = result2.get("value", [])
        exact = [m for m in matches if str(m.get("name", "")).lower() == filename.lower() and m.get("file")]

    if not exact:
        xlsx = [m for m in matches if m.get("file") and str(m.get("name", "")).lower().endswith((".xlsx", ".xlsm"))]
        if len(xlsx) == 1:
            exact = xlsx

    if not exact:
        names = [str(m.get("name", "")) for m in matches[:20]]
        raise RuntimeError(
            f"Microsoft Graph could not find the workbook '{filename}' in {upn}. "
            f"Search returned: {names}"
        )

    item = exact[0]
    item_id = item.get("id")
    item_name = item.get("name", filename)
    if not item_id:
        raise RuntimeError("Microsoft Graph returned the workbook without a DriveItem ID.")

    item_id_part = urllib.parse.quote(item_id, safe="")
    content_url = f"{GRAPH_BASE}/users/{user_part}/drive/items/{item_id_part}/content"
    req = urllib.request.Request(
        content_url,
        headers={
            "Authorization": "Bearer " + token,
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream",
            "User-Agent": "HSE-Performance-Dashboard/Graph"
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            blob = r.read()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"Microsoft Graph workbook download failed (HTTP {e.code}): {body}") from e

    if len(blob) < 1000 or blob[:2] != b"PK":
        raise RuntimeError(f"Graph returned non-XLSX content for '{item_name}' ({len(blob)} bytes).")
    try:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            names = set(z.namelist())
            if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
                raise RuntimeError(f"Downloaded '{item_name}' is not a valid XLSX workbook.")
    except zipfile.BadZipFile as e:
        raise RuntimeError(f"Downloaded '{item_name}' is not a valid XLSX ZIP file.") from e

    print(f"Downloaded workbook via Microsoft Graph: {item_name} ({len(blob):,} bytes)")
    return blob

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
    "source":"OneDrive via Microsoft Graph",
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
