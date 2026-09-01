#!/usr/bin/env python3
"""
LODHA GCR — DETERMINISTIC ENGINE (no Anthropic API calls, no API key)

This file does NOT research anything and NEVER calls a model. It exists so
the parts that must be deterministic and mechanical — sheet reads/writes,
gate math, the evidence workbook, Drive upload — are exact code, not tokens.

The RESEARCH is done by the routine's own Claude Code agent, using its own
web-search tool, which runs on your Claude Code / claude.ai subscription
quota rather than a separate pay-per-token API key. See ROUTINE_PROMPT.md
for the exact instructions the agent follows and the CLI contract below.

CLI (each subcommand prints one JSON object to stdout):

  read_queue
      -> {"rows":[{"row":7,"reference":"GCR-...","sub":{...}}, ...]}
      Lists leads at QUEUED, plus any stuck at ENRICHING for >30 min.

  claim --row N
      Marks row N ENRICHING and stamps Last Run. Call before researching it.

  evaluate --sub '<json>' --anchor '<json>' --ledger '<json>'
      -> {"gates":[...], "verdict":"Green|Amber|Red"}
      Pure gate math — ticket-scaled, multi-route capacity, exactly the
      logic in Code.gs. The agent supplies findings; this decides the rating.
      The rating is computed here and ONLY here — the agent explains it,
      never sets it.

  workbook --sub '<json>' --anchor '<json>' --ledger '<json>' --notes '<json>'
      -> {"url":"https://docs.google.com/..."}
      Builds the Alibaug-format evidence xlsx and uploads it to Drive.

  writeback --row N --verdict V --confidence C --summary S --reasoning R
             --gaps '["...","..."]' --escalations E --evidence_url U
      Writes the profile back to the row and sets Status=PENDING_APPROVAL.
      (Committee email stays in Apps Script's notifyPending — unchanged.)

  fail --row N --reason "..."
      Sets Status=FAILED and records the reason in Run Log.

ENVIRONMENT
  GCR_SHEET_ID        required
  GCR_DRIVE_FOLDER    optional — Drive folder ID for evidence workbooks
  GOOGLE_SA_JSON       the service-account JSON as one line, OR
  GOOGLE_APPLICATION_CREDENTIALS   a path to the JSON file

  pip install google-api-python-client google-auth openpyxl
  (deliberately no `anthropic` package — this file never imports it)
"""

import argparse, json, os, re, sys
from datetime import datetime, timezone

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SHEET_ID  = os.environ.get("GCR_SHEET_ID", "")
FOLDER_ID = os.environ.get("GCR_DRIVE_FOLDER", "")
LEADS     = "Leads"
STALE_MINUTES = 30

ST = dict(QUEUED="QUEUED", ENRICHING="ENRICHING", PENDING="PENDING_APPROVAL", FAILED="FAILED")

CONFIG_DEFAULTS = dict(identityMin=0.60, baseTicketCr=200, baseTurnoverFloorCr=500,
                       baseNetWorthFloorCr=200, capacityRedBelowRatio=0.5,
                       developerCarveOutCr=500)
BANDS = [("\u20B9200 \u2013 300 Cr", 200), ("\u20B9300 \u2013 500 Cr", 300),
         ("\u20B9500 \u2013 750 Cr", 500), ("Above \u20B9750 Cr", 750),
         ("More than one floor", 750)]
NF, CPSRC = "Not found in public sources", "Channel partner form"
MONS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
NICE = {"g1":"political connections", "g2":"links to a property developer",
        "g3":"practising as a lawyer", "g4":"working as a journalist",
        "5c":"court cases involving him or his fellow directors",
        "5d":"court cases involving his family", "e4":"tax or regulatory action",
        "e5":"unpaid debts or insolvency", "anchorConflict":"a clash with an existing occupier"}

def now_iso(): return datetime.now(timezone.utc).isoformat()
def dmy(dt=None):
    dt = dt or datetime.now()
    return f"{dt.day:02d}/{MONS[dt.month-1]}/{dt.year}"
def col(idx0): return get_column_letter(idx0 + 1)

# ---------------------------------------------------------------- Google IO

def gauth():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    raw = os.environ.get("GOOGLE_SA_JSON", "").strip()
    creds = (service_account.Credentials.from_service_account_info(json.loads(raw), scopes=scopes)
             if raw else
             service_account.Credentials.from_service_account_file(
                 os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=scopes))
    return build("sheets", "v4", credentials=creds), build("drive", "v3", credentials=creds)

def read_all(sheets, tab):
    return sheets.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=tab).execute().get("values", [])

def write_cells(sheets, updates):
    body = {"valueInputOption": "RAW", "data": [{"range": r, "values": [[v]]} for r, v in updates]}
    sheets.spreadsheets().values().batchUpdate(spreadsheetId=SHEET_ID, body=body).execute()

def load_config(sheets):
    cfg = dict(CONFIG_DEFAULTS)
    try:
        for row in read_all(sheets, "Config")[1:]:
            if len(row) >= 2 and row[0].strip() in cfg:
                try: cfg[row[0].strip()] = float(row[1])
                except ValueError: pass
    except Exception:
        pass
    return cfg

def leads_header(sheets):
    data = read_all(sheets, LEADS)
    return data[0], data[1:], {h: i for i, h in enumerate(data[0])}

# ---------------------------------------------------------------- gates (verbatim port of Code.gs)

def cr_from(text):
    if not text: return None
    t = str(text).replace(",", "")
    m = re.search(r"([\d.]+)\s*lakh\s*crore", t, re.I)
    if m: return float(m.group(1)) * 100000
    m = re.search(r"([\d.]+)\s*(crore|cr\b)", t, re.I)
    return float(m.group(1)) if m else None

def evaluate_gates(ledger, anchor, budget, cfg):
    G = []
    def get(fid):
        f = ledger.get(fid)
        return f if f and f.get("status") == "found" else None
    def add(fid, label, result, detail, escalate=False):
        G.append(dict(id=fid, label=label, result=result, detail=detail, escalate=escalate))

    conf = float((anchor or {}).get("confidence") or 0)
    add("identity", "Identity resolved", "pass" if conf >= cfg["identityMin"] else "unresolved",
        f"confidence {conf}" if conf else "identity step returned nothing")

    def required_for(ticket):
        k = ticket / cfg["baseTicketCr"]
        return cfg["baseTurnoverFloorCr"] * k, cfg["baseNetWorthFloorCr"] * k
    def band_for(label):
        return next((b for b in BANDS if b[0] == label), BANDS[0])

    turn = get("6b");  cr  = cr_from(turn["value"]) if turn else None
    worth = get("3a"); wcr = cr_from(worth["value"]) if worth else None
    liq = get("d2");   liq_cr = cr_from(liq["value"]) if liq else None
    if liq_cr is not None and (wcr is None or liq_cr > wcr): wcr = liq_cr
    elif liq and wcr is None: wcr = cfg["baseNetWorthFloorCr"]

    band = band_for(budget or "")
    need_t, need_w = required_for(band[1])
    best = None
    for lbl, tk in BANDS:
        nt, nw = required_for(tk)
        if ((cr is not None and cr >= nt) or (wcr is not None and wcr >= nw)) and (best is None or tk > best[1]):
            best = (lbl, tk)
    routes = []
    if cr is not None and cr >= need_t: routes.append(f"turnover {turn['value']}")
    if worth and wcr is not None and wcr >= need_w: routes.append(f"net worth {worth['value']}")
    if liq and wcr is not None and wcr >= need_w: routes.append(f"liquidity event — {liq['value']}")

    if routes:
        add("capacity", "Capacity against ticket", "pass",
            f"{'; '.join(routes)} — against Rs {need_t:g} Cr turnover or Rs {need_w:g} Cr net worth for a {band[0]} unit")
    elif best:
        add("capacity", "Capacity against ticket", "qualifies lower",
            f"supports a {best[0]} unit, not the {band[0]} unit asked about"
            + (f" — turnover Rs {cr:g} Cr" if cr is not None else "")
            + (f", net worth Rs {wcr:g} Cr" if wcr is not None else ""), True)
    elif cr is None and wcr is None:
        add("capacity", "Capacity against ticket", "not established",
            "no filed turnover, net worth or liquidity event found — confirm in the meeting")
    elif cr is not None and cr < cfg["baseTurnoverFloorCr"] * cfg["capacityRedBelowRatio"]:
        add("capacity", "Capacity against ticket", "far below",
            f"Rs {cr:g} Cr turnover, short of even the smallest unit at Rs {cfg['baseTurnoverFloorCr']:g} Cr, and no other route found")
    else:
        add("capacity", "Capacity against ticket", "short",
            f"Rs {(cr if cr is not None else wcr):g} Cr against Rs {cfg['baseTurnoverFloorCr']:g} Cr for the smallest unit, and filings lag by a year or more")

    for fid, label in [("g1","Politically exposed person"),("g2","Real estate developer interest"),
                       ("g3","Practising lawyer"),("g4","Journalist")]:
        f = get(fid)
        if not f: add(fid, label, "nothing found", "no match in the sources searched"); continue
        if fid == "g2" and cr is not None and cr >= cfg["developerCarveOutCr"]:
            add(fid, label, "exception applies", f"{f['value']} — above the Rs {cfg['developerCarveOutCr']:g} cr carve-out"); continue
        add(fid, label, "for the screening authority", f["value"], True)

    for fid, label in [("5c","Litigation, individual or co-directors"),("e4","Regulatory or tax proceedings"),
                       ("e5","Default, NPA or insolvency"),("5d","Litigation, family members")]:
        f = get(fid)
        if not f: add(fid, label, "nothing found", "no match in the sources searched"); continue
        sev = (f.get("severity") or "mention").lower(); about = (f.get("about") or "subject").lower()
        own = about in ("subject", "company")
        if sev == "finding" and own: add(fid, label, "disqualifying", f"{f['value']} — concluded, against the buyer")
        elif sev == "finding": add(fid, label, "for the committee", f"{f['value']} — concluded, but against a {about}", True)
        elif sev == "allegation": add(fid, label, "for the committee", f"{f['value']} — unproven, {about}", True)
        else: add(fid, label, "noted", f"{f['value']} — press mention only, {about}")

    add("anchorConflict", "Anchor-occupant business conflict", "untestable", "anchor list not supplied")
    return G

def rate(G):
    by = {g["id"]: g for g in G}
    if any(g["result"] == "disqualifying" for g in G): return "Red"
    if by["capacity"]["result"] == "far below": return "Red"
    if any(g["escalate"] for g in G): return "Amber"
    if by["identity"]["result"] != "pass": return "Amber"
    if by["capacity"]["result"] in ("qualifies lower", "short"): return "Amber"
    return "Green"

# ---------------------------------------------------------------- evidence xlsx (Alibaug format)

TITLE_FILL, BAND_FILL, REMARK_FILL = (PatternFill("solid", fgColor=c) for c in ("FABF8F","FCD5B4","FBF8F2"))
INK  = Font(name="Calibri", size=11, bold=True, color="1F4E79")
BOLD = Font(name="Calibri", size=11, bold=True)
BASE = Font(name="Calibri", size=10)
MUTEF = Font(name="Calibri", size=10, color="808080")
META = Font(name="Calibri", size=9, italic=True, color="808080")
WRAP = Alignment(vertical="top", wrap_text=True)

def build_rows(sub, anchor, ledger, notes):
    R = []
    v = lambda x: str(x).strip() if x and str(x).strip() else ""
    def get(fid):
        f = ledger.get(fid)
        return f if f and f.get("status") == "found" else None
    a_src = ((anchor or {}).get("sources") or [""])[0]
    def line(param, value, comment="", source="", guide="", stage="In meeting"):
        val = v(value); found = bool(val)
        R.append(dict(t="line", found=found, c=[param, val if found else NF, comment, (source if found else ""), guide, ("" if found else stage)]))
    def fld(param, fid, guide="", stage="In meeting", comment=""):
        f = get(fid)
        line(param, f["value"] if f else "", (f.get("note") if f else comment) or comment, f.get("source") if f else "", guide, stage)
    def either(param, typed, anchored, guide="", comment=""):
        val = v(typed) or v(anchored or "")
        line(param, val, comment, CPSRC if v(typed) else (a_src if val else ""), guide)
    def band(t): R.append(dict(t="band", c=[t, "", "Comments", "Source of Information", "Guideline for assessing each parameter", "Filling Stage"]))
    def remark(txt, frm): R.append(dict(t="remark", c=["Remarks", v(txt), "", "", f"Open text — {frm}" if frm else "Open text — for the assessing committee", ""]))

    R.append(dict(t="title", c=["LODHA GOLF COURSE ROAD — CLIENT PROFILE"]))
    R.append(dict(t="meta", c=["Reference", sub["reference"]]))
    R.append(dict(t="meta", c=["Prepared", dmy() + " " + datetime.now().strftime("%H:%M")]))
    R.append(dict(t="meta", c=["Identity confidence", str(anchor.get("confidence")) if anchor and anchor.get("confidence") is not None else "not resolved"]))
    R.append(dict(t="meta", c=["Status", "Evidence file — not assessed"]))
    R.append(dict(t="spacer", c=[]))

    band("Client Details")
    line("Client Name", sub["clientName"], "", CPSRC)
    line("Existing Lodha Buyer?", sub.get("existing"), "Needs the CC / HPM record to confirm", CPSRC)
    line("Occupation", sub.get("role"), "", CPSRC)
    either("Company Name", sub.get("company"), anchor.get("company") if anchor else "")
    either("Office Address", sub.get("officeCity"), anchor.get("office_address") if anchor else "")
    either("Designation", sub.get("role"), anchor.get("designation") if anchor else "")
    either("Official company website link:", sub.get("website"), anchor.get("website") if anchor else "")
    remark(anchor.get("note") if anchor else "", "Identity step")

    band("Client Requirement & Source Details")
    line("Budget", sub.get("budget"), "Units start at ₹200 Cr", CPSRC)
    line("Configuration interested in", " · ".join(x for x in [v(sub.get("area")), v(sub.get("purpose"))] if x), "", CPSRC)
    line("Source", "Channel Partner", "", "Intake form")
    line("Source Details ( CP Company name)",
         " · ".join(x for x in [v(sub.get("cpFirm")), v(sub.get("cpSpoc")), f"Reap {sub['reapId']}" if v(sub.get("reapId")) else ""] if x),
         "", CPSRC, "", "Before walkin")
    line("Sourcing SM", " · ".join(x for x in [v(sub.get("smName")), v(sub.get("smEmail"))] if x), "", CPSRC)
    remark(sub.get("notes"), "Channel partner")

    band("1. Primary Residence")
    fld("1a. Residential Building", "1a", "4/5 BHK")
    fld("1b. Location", "1b")
    line("1b. Owned / Rented", "", "", "", "Owned")
    fld("1c. Home market value", "1c", "20 cr+ home value")
    remark(notes.get("lifestyle"), "Residence and affiliations")

    band("2. Secondary Residences")
    fld("2a. Secondary Real Estate Investments/Locations", "2a")
    line("2b. Existing Properties in Gurugram / NCR", "")
    remark("", "")

    band("3. Net worth")
    fld("3a. Estimated Net worth:\n- Company ownership value + Other assets", "3a", "100 Cr+")
    fld("3b. Annual Income", "3b", "5 Cr+")
    fld("d2. Liquidity events", "d2", "Sale, acquisition, listing, funding round or large dividend — value and year")
    remark(notes.get("corporate"), "Company filings")

    band("4. Lifestyle")
    fld("4a. Luxury ecosystem affiliations:\n- Memberships in Golf club, Yacht club, business clubs", "4a")
    line("4b. Sponsor/Organizer/Patron in Art fairs, galleries, events", "")
    line("4c. High-profile events attended", "")
    fld("4d. Involvement in philanthropic activities or trusteeship", "4d")
    remark("", "")

    band("5. Reputation")
    R.append(dict(t="sub", c=["Industry Reputation"]))
    f5a = get("5a")
    line("5a. Seniority in the organization (MD/CEO, MD/CEO-1, MD/CEO-2)",
         f5a["value"] if f5a else (sub.get("role") or (anchor.get("designation") if anchor else "")),
         "", (f5a.get("source") if f5a else (CPSRC if sub.get("role") else a_src)), "MD/CEO minus 2 or higher")
    fld("5b. Positive reputation in the industry", "5b")
    fld("5c. Any litigations or negative news about self or other directors", "5c")
    R.append(dict(t="sub", c=["Social Reputation"]))
    fld("5d. Any litigations or negative news about family members", "5d")
    line("5e. Closed network reference check", "", "Needs the CC / HPM feed — internal lookup", "", "Delayed CAM, conduct in existing society", "Customer Care")
    hits = [get(g) for g in ("g1","g2","g3","g4") if get(g)]
    line("5f. Sensitive Profile check", "; ".join(h["value"] for h in hits),
         "Community screen deliberately excluded — handled offline by the committee",
         hits[0]["source"] if hits else "", "Per policy 05.10.24 v3")
    remark("  —  ".join(x for x in [notes.get("news"), notes.get("litigation"), notes.get("sensitive")] if x), "Press, court and screening lanes")

    band("6. Entity & Occupier — Golf Course Road")
    fld("6a. Entity type", "6a")
    fld("6b. Annual turnover", "6b", "Scaled to the unit asked about")
    fld("6c. PAT trend, three years", "6c")
    fld("6d. Group / related entities", "6d")
    either("6e. Current office, owned or leased", sub.get("currentOffice"), (get("6e") or {}).get("value"))
    fld("6f. Headcount", "6f")
    either("6g. Industry / business line", sub.get("industry"), anchor.get("industry") if anchor else "")
    line("6h. Anchor-occupant business conflict", "", "Not testable — anchor list not supplied", "", "Same business line as any occupant holding 20%+", "Sales lead")
    remark("", "")

    band("7. Standing & Influence — Golf Course Road")
    fld("7a. Board or advisory positions at other companies", "i1")
    fld("7b. Industry body roles", "i2")
    fld("7c. Trusteeship or institutional association", "i3")
    fld("7d. Corporate philanthropy", "i4")
    fld("7e. Public platform", "i5")
    fld("7f. Business lineage", "i6")
    remark(notes.get("standing"), "Standing and influence")

    R.append(dict(t="spacer", c=[]))
    R.append(dict(t="note", c=["Blank cells mean nothing was found in the sources searched. That is not the same as a clear finding."]))
    return R

def write_workbook(rows, sub, out_dir="/tmp"):
    wb = Workbook(); ws = wb.active; ws.title = "Evidence"
    for i, w in enumerate([55, 31.8, 28.8, 27.4, 54.1, 18]):
        ws.column_dimensions[col(i)].width = w
    ws.freeze_panes = "D1"
    for n, r in enumerate(rows, start=1):
        c = (r["c"] + [""] * 6)[:6]
        for j, val in enumerate(c, start=1):
            cell = ws.cell(row=n, column=j, value=val or None); cell.font = BASE; cell.alignment = WRAP
        t = r["t"]
        if t == "title":
            ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=3)
            for j in range(1, 7): ws.cell(row=n, column=j).fill = TITLE_FILL
            ws.cell(row=n, column=1).font = INK
        elif t == "meta":
            ws.cell(row=n, column=1).font = META; ws.cell(row=n, column=2).font = META
        elif t == "band":
            ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=2)
            for j in range(1, 7): ws.cell(row=n, column=j).fill = BAND_FILL
            ws.cell(row=n, column=1).font = BOLD
        elif t == "sub":
            ws.cell(row=n, column=1).font = BOLD
        elif t == "note":
            ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=6)
            ws.cell(row=n, column=1).font = META
        elif t == "remark":
            ws.merge_cells(start_row=n, start_column=2, end_row=n, end_column=4)
            ws.merge_cells(start_row=n, start_column=5, end_row=n, end_column=6)
            for j in range(1, 7): ws.cell(row=n, column=j).fill = REMARK_FILL
            ws.cell(row=n, column=1).font = Font(name="Calibri", size=10, italic=True)
            ws.cell(row=n, column=5).font = META
            ws.row_dimensions[n].height = 33
        elif t == "line" and not r["found"]:
            ws.cell(row=n, column=2).font = MUTEF
    safe = re.sub(r"\s+", "_", sub["clientName"])
    path = os.path.join(out_dir, f"{sub['reference']}_{safe}.xlsx")
    wb.save(path)
    return path

def upload_to_drive(drive, path, ref):
    meta = {"name": os.path.basename(path), "mimeType": "application/vnd.google-apps.spreadsheet"}
    if FOLDER_ID: meta["parents"] = [FOLDER_ID]
    media = MediaFileUpload(path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    f = drive.files().create(body=meta, media_body=media, fields="id,webViewLink", supportsAllDrives=True).execute()
    return f.get("webViewLink", "")

# ---------------------------------------------------------------- CLI

def cmd_read_queue(a):
    sheets, _ = gauth()
    header, data, H = leads_header(sheets)
    out = []
    for i, vals in enumerate(data, start=2):
        g = lambda c: str(vals[H[c]]) if H.get(c) is not None and H[c] < len(vals) else ""
        status = g("Status")
        stale = False
        if status == ST["ENRICHING"]:
            try:
                age = (datetime.now(timezone.utc) - datetime.fromisoformat(g("Last Run").replace("Z","+00:00"))).total_seconds()/60
                stale = age > STALE_MINUTES
            except Exception:
                stale = True
        if status == ST["QUEUED"] or (status == ST["ENRICHING"] and stale):
            sub = dict(clientName=g("Client Name"), company=g("Company"), role=g("Role"),
                       industry=g("Industry"), officeCity=g("Based At"), website=g("Website"),
                       linkedin=g("LinkedIn"), budget=g("Budget"), area=g("Floor Area"),
                       purpose=g("Intended Use"), currentOffice=g("Office Today"),
                       existing=g("Existing Lodha Buyer"), notes=g("Notes"),
                       cpFirm=g("CP Firm"), cpSpoc=g("CP SPOC"), reapId=g("Reap ID"),
                       smName=g("SM Name"), smEmail=g("Owner Email"), reference=g("Reference"))
            out.append(dict(row=i, reference=sub["reference"], sub=sub, was_stale=stale))
    print(json.dumps(dict(rows=out)))

def cmd_claim(a):
    sheets, _ = gauth()
    _, _, H = leads_header(sheets)
    write_cells(sheets, [(f"{LEADS}!{col(H['Status'])}{a.row}", ST["ENRICHING"]),
                        (f"{LEADS}!{col(H['Last Run'])}{a.row}", now_iso())])
    print(json.dumps(dict(ok=True)))

def cmd_evaluate(a):
    sheets, _ = gauth()
    cfg = load_config(sheets)
    sub = json.loads(a.sub); anchor = json.loads(a.anchor) if a.anchor else None
    ledger = json.loads(a.ledger)
    gates = evaluate_gates(ledger, anchor, sub.get("budget",""), cfg)
    print(json.dumps(dict(gates=gates, verdict=rate(gates))))

def cmd_workbook(a):
    sheets, drive = gauth()
    sub = json.loads(a.sub); anchor = json.loads(a.anchor) if a.anchor else {}
    ledger = json.loads(a.ledger); notes = json.loads(a.notes) if a.notes else {}
    path = write_workbook(build_rows(sub, anchor, ledger, notes), sub)
    url = ""
    try:
        url = upload_to_drive(drive, path, sub["reference"])
    except Exception as e:
        print(json.dumps(dict(url="", local_path=path, error=str(e)))); return
    print(json.dumps(dict(url=url)))

def cmd_writeback(a):
    sheets, _ = gauth()
    _, _, H = leads_header(sheets)
    gaps = json.loads(a.gaps) if a.gaps else []
    try:
        vals = read_all(sheets, LEADS)[a.row - 1]
        inter = json.loads(vals[H["Interactions"]]) if H.get("Interactions") is not None and H["Interactions"] < len(vals) and vals[H["Interactions"]] else []
    except Exception:
        inter = []
    inter.append(dict(t=now_iso(), who="System", text="Profile enriched and sent for approval."))
    a1 = lambda c: f"{LEADS}!{col(H[c])}{a.row}"
    write_cells(sheets, [
        (a1("Rating"), a.verdict), (a1("Identity Confidence"), a.confidence),
        (a1("Profile Summary"), a.summary), (a1("Why This Rating"), a.reasoning),
        (a1("Meeting Should Establish"), "\n".join("• " + g for g in gaps)),
        (a1("Escalations"), a.escalations or ""), (a1("Evidence Report"), a.evidence_url or ""),
        (a1("Interactions"), json.dumps(inter)), (a1("Status"), ST["PENDING"]),
        (a1("Last Run"), now_iso()),
    ])
    print(json.dumps(dict(ok=True)))

def cmd_fail(a):
    sheets, _ = gauth()
    _, _, H = leads_header(sheets)
    write_cells(sheets, [(f"{LEADS}!{col(H['Status'])}{a.row}", ST["FAILED"]),
                        (f"{LEADS}!{col(H['Run Log'])}{a.row}", ("FAILED: " + a.reason)[:400])])
    print(json.dumps(dict(ok=True)))

def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("read_queue").set_defaults(fn=cmd_read_queue)
    c = sub.add_parser("claim"); c.add_argument("--row", type=int, required=True); c.set_defaults(fn=cmd_claim)
    c = sub.add_parser("evaluate")
    c.add_argument("--sub", required=True); c.add_argument("--anchor", default="")
    c.add_argument("--ledger", required=True); c.set_defaults(fn=cmd_evaluate)
    c = sub.add_parser("workbook")
    c.add_argument("--sub", required=True); c.add_argument("--anchor", default="")
    c.add_argument("--ledger", required=True); c.add_argument("--notes", default="")
    c.set_defaults(fn=cmd_workbook)
    c = sub.add_parser("writeback")
    c.add_argument("--row", type=int, required=True); c.add_argument("--verdict", required=True)
    c.add_argument("--confidence", default=""); c.add_argument("--summary", required=True)
    c.add_argument("--reasoning", required=True); c.add_argument("--gaps", default="[]")
    c.add_argument("--escalations", default=""); c.add_argument("--evidence_url", default="")
    c.set_defaults(fn=cmd_writeback)
    c = sub.add_parser("fail"); c.add_argument("--row", type=int, required=True)
    c.add_argument("--reason", required=True); c.set_defaults(fn=cmd_fail)
    if not SHEET_ID:
        print("GCR_SHEET_ID is not set", file=sys.stderr); sys.exit(2)
    a = p.parse_args(); a.fn(a)

if __name__ == "__main__":
    main()
