"""
Build script: EMEA DG Curators+Designers Dashboard
Outputs: /tmp/dash_v4.json  (inject into HTML with re.sub)
"""
import json, re, pathlib, warnings
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

BASE = pathlib.Path("/Users/i755892/Desktop/Claude/CURATORS+DESIGNERS/01 - Source Data")

# ── helpers ────────────────────────────────────────────────────────────────────
def safe_str(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()

def norm_wbs(v):
    return re.sub(r'\s+', '', safe_str(v)).upper()

def norm_opp_id(v):
    s = safe_str(v)
    s = re.sub(r'^[Dd]eal\s*[-–]\s*', '', s).strip()
    # keep only digits
    digits = re.sub(r'\D', '', s)
    return digits if len(digits) >= 6 else ""

# ── MU normalisation ───────────────────────────────────────────────────────────
MU_MAP = {
    "CS NORDIC":            "CS Nordics",
    "CS NORDICS":           "CS Nordics",
    "CS SOUTHERN EUROPE":   "CS Spain",       # in DG Plan files = Spain
    "CS SPAIN":             "CS Spain",
    "CS TURKEY":            "CS Southern Europe",
    "CS TÜRKIYE":           "CS Southern Europe",
    "CS TURKEY/ISRAEL/PORTUGAL/GREECE&CYPRUS": "CS Southern Europe",
    "CS TÜRKIYE/ISRAEL/PORTUGAL/GREECE&CYPRUS":"CS Southern Europe",
}
def norm_mu(v):
    k = re.sub(r'\s+', ' ', safe_str(v)).upper().strip()
    return MU_MAP.get(k, safe_str(v).strip() or "Unknown")

# ── asset category map ─────────────────────────────────────────────────────────
ASSET_CAT = {
    # Content & Sales Assets
    "ONE PAGER":                        "Content & Sales Assets",
    "INFOGRAPHIC":                      "Content & Sales Assets",
    "STORYCRAFTING":                    "Content & Sales Assets",
    "ASSET FROM TEMPLATE LIBRARY":      "Content & Sales Assets",
    # Landing Pages
    "FOLLOZE BOARD":                    "Landing Pages",
    "MICROSITE (TIMESITES)":            "Landing Pages",
    "LANDING PAGE / MICROSITE":         "Landing Pages",
    "LANDING PAGE / CONTENT REPOSITORY":"Landing Pages",
    "SHAREPOINT LANDING PAGE":          "Landing Pages",
    "SWAY LANDING PAGE":                "Landing Pages",
    "SURVEY":                           "Landing Pages",
    # Email Assets
    "EMAIL DESIGN":                     "Email Assets",
    "EMAIL – BANNER & TEXT":            "Email Assets",
    "EMAIL - BANNER & TEXT":            "Email Assets",
    "EMAIL – BANNER":                   "Email Assets",
    "EMAIL - BANNER":                   "Email Assets",
    "EMAIL – BANNER, TEXT & VIDEO/VISUAL":"Email Assets",
    "EMAIL - BANNER, TEXT & VIDEO/VISUAL":"Email Assets",
    # Creative & Design
    "BRANDING / LOGO":                  "Creative & Design",
    "LOGO":                             "Creative & Design",
    "LOGO / BRANDING":                  "Creative & Design",
    "FLYER / POSTER":                   "Creative & Design",
    "IMAGE / FLYER / POSTER":           "Creative & Design",
    "IMAGE / FLYER / POSTER / EMAIL":   "Creative & Design",
    "SOCIAL MEDIA VISUAL":              "Creative & Design",
    # Video Production
    "VIDEO – TRANSLATION":              "Video Production",
    "VIDEO - TRANSLATION":              "Video Production",
    "VIDEO – BASIC EDITING":            "Video Production",
    "VIDEO - BASIC EDITING":            "Video Production",
    "VIDEO – EDITING":                  "Video Production",
    "VIDEO - EDITING":                  "Video Production",
    "VIDEO – RECORDING & EDITING":      "Video Production",
    "VIDEO - RECORDING & EDITING":      "Video Production",
    "VIDEO – INTERVIEW & BASIC EDITING":"Video Production",
    "VIDEO - INTERVIEW & BASIC EDITING":"Video Production",
    "VIDEO – EDITING WITH MOTION GRAPHICS":"Video Production",
    "VIDEO - EDITING WITH MOTION GRAPHICS":"Video Production",
    "VIDEO – EDITING WITH HOLOGRAM ANIMATION":"Video Production",
    "VIDEO - EDITING WITH HOLOGRAM ANIMATION":"Video Production",
    "VIDEO – GREEN SCREEN RECORDING & EDITING":"Video Production",
    "VIDEO - GREEN SCREEN RECORDING & EDITING":"Video Production",
    "VIDEO – ANIMATION":                "Video Production",
    "VIDEO - ANIMATION":                "Video Production",
    "VIDEO – ANIMATION (WITHOUT REAL PERSON)":"Video Production",
    "VIDEO - ANIMATION (WITHOUT REAL PERSON)":"Video Production",
    # Presentation
    "PRESENTATION":                     "Presentation",
    "PRESENTATION (PPT, SWAY, STORYBOARD, MAR...)": "Presentation",
    "PRESENTATION (PPT, SWAY, STORYBOARD, SHA...)": "Presentation",
    # BEC
    "EXPERIENCE PACKAGE BEC":           "BEC",
    "APP PROTOTYPE":                    "BEC",
    # Development & Interactive
    "PROTOTYPE":                        "Development & Interactive",
    # Internal
    "TEAM INITIATIVE (INTERNAL)":       "Internal",
    # Others (catch-all including picture editing, please select, digital magazine, etc.)
}
def cat_asset(v):
    k = safe_str(v).upper().strip()
    return ASSET_CAT.get(k, "Others")

# ══════════════════════════════════════════════════════════════════════════════
# 1. LOAD DG PLANS
# ══════════════════════════════════════════════════════════════════════════════
print("Loading DG Plans…")

MU_FILES = {
    "CS UKI":             "DG Plans/UKI - DG Action Plan 2026 CUR+DES.xlsx",
    "CS Nordics":         "DG Plans/Nordics - DG Action Plan 2026 CUR+DES.xlsx",
    "CS Spain":           "DG Plans/Spain - DG Action Plan Q1  2026 CUR+DES.xlsx",
    "CS Spain_SE":        "DG Plans/Southern Europe - DG Action Plan 2026 CUR+DES.xlsx",
    "CS BeNeLux":         "DG Plans/BeNeLux- DG Action Plan 2026 v1 CUR+DES.xlsx",
    "CS France":          "DG Plans/France - DG Action Plan 2026 CUR+DES.xlsx",
    "CS MEA South":       "DG Plans/MEA South - DG Action Plan 2026 v1 CUR+DES.xlsx",
    "CS Italy":           "DG Plans/Italy - DG Action Plan 2026 CUR+DES.xlsx",
}
LOB_FILES = {
    "BTP":    "DG Plans/BTP - DG Action Plan 2026 v1.xlsx",
    "F&S":    "DG Plans/F&S - DG Action Plan 2026 v1.xlsx",
    "HCM&CX": "DG Plans/HCM & CX - DG Action Plan 2026 v1.xlsx",
    "SCM":    "DG Plans/SCM - DG Action Plan 2026 v1.xlsx",
}
# Italy override: comes in as CSV in the folder
ITALY_CSV = BASE / "DG Plans/Italy - DG Action Plan 2026 CUR+DES.xlsx"

def seq_id_from_url(url):
    m = re.search(r'/sequences/(\d+)', safe_str(url))
    return m.group(1) if m else ""

def load_dg_plan(path, mu_label, source):
    """Load one DG plan file. Returns list of campaign dicts."""
    p = BASE / path
    if not p.exists():
        print(f"  MISSING: {path}")
        return []
    try:
        if str(p).endswith('.csv'):
            df = pd.read_csv(p, dtype=str)
        else:
            df = pd.read_excel(p, dtype=str, sheet_name=0)
    except Exception as e:
        print(f"  ERROR loading {path}: {e}")
        return []

    df.columns = [safe_str(c).strip() for c in df.columns]

    # column aliases
    def fc(*names):
        for n in names:
            for c in df.columns:
                if c.lower().replace(' ','').replace('_','') == n.lower().replace(' ','').replace('_',''):
                    return c
                if n.lower() in c.lower():
                    return c
        return None

    col_name     = fc('Campaign Name','Campaign','Name','Campaignname')
    col_wbs      = fc('WBS Code','WBS','wbs','WBSCode')
    col_seq      = fc('Sequence Link','SequenceLink','Outreach Link','OutreachLink','Sequence URL')
    col_status   = fc('Status','Campaign Status')
    col_priority = fc('Priority','Prio')
    col_industry = fc('Industry','Industries','Sector')
    col_executor = fc('Executor','Responsible','Owner','SDE')
    col_quarter  = fc('Quarter','Qtr','Q')
    col_target   = fc('Target Pipeline','Target Pipeline - value kEUR','TargetPipeline','Pipeline Target','Target')
    col_solution = fc('Solution','LOB','Line of Business','Product')

    rows = []
    for _, r in df.iterrows():
        name   = safe_str(r.get(col_name,''))    if col_name   else ''
        wbs    = norm_wbs(r.get(col_wbs,''))     if col_wbs    else ''
        seq_url= safe_str(r.get(col_seq,''))     if col_seq    else ''
        seq_id = seq_id_from_url(seq_url)
        status = safe_str(r.get(col_status,''))  if col_status else ''
        prio   = safe_str(r.get(col_priority,''))if col_priority else ''
        ind    = safe_str(r.get(col_industry,''))if col_industry else ''
        exe    = safe_str(r.get(col_executor,''))if col_executor else ''
        qtr    = safe_str(r.get(col_quarter,'')) if col_quarter else ''
        sol    = safe_str(r.get(col_solution,''))if col_solution else ''

        tgt = 0.0
        if col_target:
            try:
                raw = safe_str(r.get(col_target,''))
                raw = re.sub(r'[€$,\s]','',raw)
                v   = float(raw) if raw else 0.0
                tgt = v/1000 if v > 5000 else v
            except:
                pass

        if not name:
            continue

        mu_clean = mu_label
        if mu_label == "CS Spain_SE":
            mu_clean = "CS Southern Europe"

        rows.append({
            "campaign": name, "wbs": wbs, "seq_id": seq_id, "seq_link": seq_url,
            "mu": mu_clean, "source": source, "solution": sol,
            "status": status, "priority": prio, "industry": ind,
            "executor": exe, "quarter": qtr, "target_pl": tgt,
        })
    return rows

campaigns_raw = []
for mu, fpath in MU_FILES.items():
    rows = load_dg_plan(fpath, mu, "mu")
    print(f"  {mu}: {len(rows)} rows")
    campaigns_raw.extend(rows)

for lob, fpath in LOB_FILES.items():
    rows = load_dg_plan(fpath, lob, "lob")
    print(f"  {lob}: {len(rows)} rows")
    campaigns_raw.extend(rows)

# ══════════════════════════════════════════════════════════════════════════════
# 2. LOAD OUTREACH STATS
# ══════════════════════════════════════════════════════════════════════════════
print("Loading Outreach stats…")
try:
    xf = pd.read_excel(BASE / "Outreach/Outreach Reports CUR+DES.xlsx", dtype=str, sheet_name=None)
    sheet = next((xf[s] for s in xf if 'sequence' in s.lower() or 'stat' in s.lower()), list(xf.values())[0])
except Exception as e:
    print(f"  ERROR: {e}")
    sheet = pd.DataFrame()

sheet.columns = [safe_str(c).strip() for c in sheet.columns]

def fc_sheet(df, *names):
    for n in names:
        for c in df.columns:
            if n.lower() in c.lower():
                return c
    return None

sc_id    = fc_sheet(sheet, 'Sequence ID','ID','SequenceID')
sc_name  = fc_sheet(sheet, 'Sequence Name','Name','Sequence')
sc_owner = fc_sheet(sheet, 'Owner','Created By','SDE')
sc_ptot  = fc_sheet(sheet, 'Prospects Total','Total Prospects','Prospects')
sc_prep  = fc_sheet(sheet, 'Prospects Replied','Replied')
sc_popen = fc_sheet(sheet, 'Prospects Opened','Opened')
sc_popt  = fc_sheet(sheet, 'Prospects Opted','Opted')
sc_edel  = fc_sheet(sheet, 'Emails Deliveries','Deliveries','Delivered','Emails Delivered')
sc_eclk  = fc_sheet(sheet, 'Emails Clicks','Clicks')

seq_stats = {}
for _, r in sheet.iterrows():
    sid = safe_str(r.get(sc_id,'')) if sc_id else ''
    sid = re.sub(r'\D','',sid)
    if not sid:
        continue
    def fv(col):
        if not col: return 0
        try: return int(float(safe_str(r.get(col,0)) or 0))
        except: return 0
    seq_stats[sid] = {
        "name":   safe_str(r.get(sc_name,''))  if sc_name  else '',
        "owner":  safe_str(r.get(sc_owner,'')) if sc_owner else '',
        "p_total":   fv(sc_ptot),
        "p_replied": fv(sc_prep),
        "p_opened":  fv(sc_popen),
        "p_opted":   fv(sc_popt),
        "e_deliver": fv(sc_edel),
        "e_clicks":  fv(sc_eclk),
    }
print(f"  {len(seq_stats)} sequences loaded")

# ══════════════════════════════════════════════════════════════════════════════
# 3. LOAD CURATORS
# ══════════════════════════════════════════════════════════════════════════════
print("Loading Curators…")
try:
    cf = pd.read_excel(BASE / "Curators/CURATORS X MSLIST REQUEST CUR+DES.xlsx", dtype=str, sheet_name=0)
    cf.columns = [safe_str(c).strip() for c in cf.columns]
    wbs_col  = fc_sheet(cf, 'WBS','Campaign ID','wbs')
    mu_col   = fc_sheet(cf, 'Region','MU','Market Unit','mu')
    curator_set = set()
    for _, r in cf.iterrows():
        w = norm_wbs(r.get(wbs_col,'')) if wbs_col else ''
        m = norm_mu(r.get(mu_col,''))   if mu_col  else ''
        if w:
            curator_set.add((w, m))
            curator_set.add((w, ''))  # fallback: WBS only
    print(f"  {len(curator_set)} curator (wbs,mu) pairs")
except Exception as e:
    print(f"  ERROR: {e}")
    curator_set = set()

# ══════════════════════════════════════════════════════════════════════════════
# 4. LOAD PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
print("Loading Pipeline…")
try:
    pf = pd.read_excel(BASE / "Pipeline/Global Cloud Pipeline Data Analyzer CUR+DES.xlsx", dtype=str, sheet_name=0)
    pf.columns = [safe_str(c).strip() for c in pf.columns]
    pc_wbs     = fc_sheet(pf, 'Opp Campaign ID','Campaign ID','WBS')
    pc_oppid   = fc_sheet(pf, 'Opportunity ID','OppID','Opp ID')
    # ACV real value is in Column1 (numeric kEUR); Deal Size (ACV EUR) is a text bucket
    pc_acv     = fc_sheet(pf, 'Column1','ACV kEUR','ACV')
    pc_opps    = fc_sheet(pf, 'Column2','# of Opps','Opps','Num Opps')
    pc_drmcat  = fc_sheet(pf, 'DRM Category','DRM Cat')
    pc_status  = fc_sheet(pf, 'Opp Status','Status')
    pc_phase   = fc_sheet(pf, 'Opp Phase','Phase')
    pc_desc    = fc_sheet(pf, 'Opp Description','Description','Opp Desc')
    pc_account = fc_sheet(pf, 'Account Name','Account')
    pc_region  = fc_sheet(pf, 'Region Lvl 2','Region','MU')

    pipeline = {}   # wbs -> {opps, acv, rows:[...]}
    opp_rows = []   # all individual opp rows for Designers match

    for _, r in pf.iterrows():
        wbs    = norm_wbs(r.get(pc_wbs,''))  if pc_wbs else ''
        opp_id = norm_opp_id(r.get(pc_oppid,'')) if pc_oppid else ''
        try:    acv = float(safe_str(r.get(pc_acv,0)) or 0)
        except: acv = 0.0
        try:    n_opps = int(float(safe_str(r.get(pc_opps,1)) or 1))
        except: n_opps = 1

        row_data = {
            "opp_id":      opp_id,
            "wbs":         wbs,
            "account":     safe_str(r.get(pc_account,''))  if pc_account else '',
            "description": safe_str(r.get(pc_desc,''))     if pc_desc    else '',
            "drm_cat":     safe_str(r.get(pc_drmcat,''))   if pc_drmcat  else '',
            "opp_status":  safe_str(r.get(pc_status,''))   if pc_status  else '',
            "opp_phase":   safe_str(r.get(pc_phase,''))    if pc_phase   else '',
            "region":      safe_str(r.get(pc_region,''))   if pc_region  else '',
            "acv":         acv,
        }
        opp_rows.append(row_data)

        if wbs:
            if wbs not in pipeline:
                pipeline[wbs] = {"opps": 0, "acv": 0.0}
            pipeline[wbs]["opps"] += n_opps
            pipeline[wbs]["acv"]  += acv

    # index by opp_id too for Designers match
    pipeline_by_opp = {}
    for row in opp_rows:
        if row["opp_id"]:
            pipeline_by_opp[row["opp_id"]] = row

    print(f"  {len(pipeline)} WBS codes with pipeline, {len(opp_rows)} opp rows")
except Exception as e:
    print(f"  ERROR: {e}")
    pipeline = {}; pipeline_by_opp = {}; opp_rows = []

# ══════════════════════════════════════════════════════════════════════════════
# 5. LOAD DESIGNERS
# ══════════════════════════════════════════════════════════════════════════════
print("Loading Designers…")
try:
    df_des = pd.read_excel(BASE / "Designers/DESIGNERS CREATIVE TEAM REQUEST REVIEW.xlsx", dtype=str, sheet_name=0)
    df_des.columns = [safe_str(c).strip() for c in df_des.columns]

    dc_id      = fc_sheet(df_des, 'ID')
    dc_name    = fc_sheet(df_des, 'Name of the Project','Project Name','Name')
    dc_owner   = fc_sheet(df_des, 'Project Owner','Owner')
    dc_reqby   = fc_sheet(df_des, 'Requested By','Requestor')
    dc_reqfull = fc_sheet(df_des, 'Requested by (Full name)','Full name')
    dc_due     = fc_sheet(df_des, 'Requested Due Date','Due Date')
    dc_status  = fc_sheet(df_des, 'Status of Creation','Status')
    dc_purpose = fc_sheet(df_des, 'Purpose of asset','Purpose')
    dc_opp     = fc_sheet(df_des, 'Opportunity ID','Opp ID')
    dc_campid  = fc_sheet(df_des, 'Campaign ID','CampaignID')
    dc_asset   = fc_sheet(df_des, 'Type of Asset','Asset Type','Type')
    dc_iac     = fc_sheet(df_des, 'Internal Account Classification','IAC')
    dc_region  = fc_sheet(df_des, 'Region L2','Region')
    dc_industry= fc_sheet(df_des, 'Industry')
    dc_orgpart = fc_sheet(df_des, 'Which part of the organisation','Org')
    dc_nassets = fc_sheet(df_des, 'How many assets','Num Assets','# Assets')
    dc_hours   = fc_sheet(df_des, 'Number of hours','Hours')
    dc_link    = fc_sheet(df_des, 'Link to asset','Link')
    dc_created = fc_sheet(df_des, 'Created','Created Date')
    dc_vat     = fc_sheet(df_des, 'Added to the VAT','VAT')
    dc_topic   = fc_sheet(df_des, 'Consensus - Topic','Topic')

    des_rows = []
    for _, r in df_des.iterrows():
        raw_opp    = safe_str(r.get(dc_opp,''))   if dc_opp    else ''
        raw_campid = safe_str(r.get(dc_campid,''))if dc_campid else ''
        opp_id     = norm_opp_id(raw_opp)
        wbs        = norm_wbs(raw_campid)
        asset_type = safe_str(r.get(dc_asset,'')) if dc_asset else ''

        # skip entirely empty rows
        proj_name = safe_str(r.get(dc_name,''))   if dc_name else ''
        if not proj_name and not opp_id and not wbs:
            continue

        # skip "Please select" asset type (dirty data)
        if asset_type.strip().lower() == 'please select':
            asset_type = ''

        try:    n_assets = int(float(safe_str(r.get(dc_nassets,1)) or 1))
        except: n_assets = 1
        try:    hours = float(safe_str(r.get(dc_hours,0)) or 0)
        except: hours = 0.0

        # match to pipeline opp
        pl_opp = pipeline_by_opp.get(opp_id, {}) if opp_id else {}
        # also check if wbs maps to pipeline
        pl_wbs = pipeline.get(wbs, {}) if wbs else {}

        des_rows.append({
            "id":          safe_str(r.get(dc_id,''))      if dc_id      else '',
            "name":        proj_name,
            "owner":       safe_str(r.get(dc_owner,''))   if dc_owner   else '',
            "req_by":      safe_str(r.get(dc_reqfull,'')) or safe_str(r.get(dc_reqby,'')) if dc_reqfull or dc_reqby else '',
            "due_date":    safe_str(r.get(dc_due,''))     if dc_due     else '',
            "status":      safe_str(r.get(dc_status,''))  if dc_status  else '',
            "purpose":     safe_str(r.get(dc_purpose,'')) if dc_purpose else '',
            "opp_id":      opp_id,
            "wbs":         wbs,
            "asset_type":  asset_type,
            "asset_cat":   cat_asset(asset_type) if asset_type else "Unknown",
            "iac":         safe_str(r.get(dc_iac,''))     if dc_iac     else '',
            "region":      safe_str(r.get(dc_region,''))  if dc_region  else '',
            "industry":    safe_str(r.get(dc_industry,''))if dc_industry else '',
            "org_part":    safe_str(r.get(dc_orgpart,'')) if dc_orgpart else '',
            "n_assets":    n_assets,
            "hours":       hours,
            "link":        safe_str(r.get(dc_link,''))    if dc_link    else '',
            "created":     safe_str(r.get(dc_created,'')) if dc_created else '',
            "vat":         safe_str(r.get(dc_vat,''))     if dc_vat     else '',
            "topic":       safe_str(r.get(dc_topic,''))   if dc_topic   else '',
            # pipeline data from opp match
            "pl_account":  pl_opp.get("account",""),
            "pl_desc":     pl_opp.get("description",""),
            "pl_drm_cat":  pl_opp.get("drm_cat",""),
            "pl_opp_status":pl_opp.get("opp_status",""),
            "pl_opp_phase":pl_opp.get("opp_phase",""),
            "pl_region":   pl_opp.get("region","") or pl_wbs.get("region",""),
            "pl_acv":      pl_opp.get("acv", pl_wbs.get("acv",0.0)),
            "has_opp":     bool(pl_opp),
            "has_wbs":     bool(wbs),
        })

    print(f"  {len(des_rows)} designer request rows")
    print(f"  {sum(1 for r in des_rows if r['has_opp'])} matched to pipeline opp")
    print(f"  {sum(1 for r in des_rows if r['has_wbs'])} have Campaign ID (WBS)")
except Exception as e:
    import traceback; traceback.print_exc()
    print(f"  ERROR: {e}")
    des_rows = []

# ══════════════════════════════════════════════════════════════════════════════
# 6. MERGE DG PLAN + OUTREACH
# ══════════════════════════════════════════════════════════════════════════════
print("Merging campaigns…")
seen_seqid = set()
AC = []

for c in campaigns_raw:
    sid  = c["seq_id"]
    wbs  = c["wbs"]
    mu   = c["mu"]
    src  = c["source"]

    if not sid:
        continue
    if sid in seen_seqid:
        continue
    seen_seqid.add(sid)

    stats = seq_stats.get(sid, {})
    if not stats:
        continue  # must have Outreach data

    pt = stats["p_total"]
    pr = stats["p_replied"]
    po = stats["p_opened"]
    op = stats["p_opted"]
    ed = stats["e_deliver"]
    ec = stats["e_clicks"]

    rr = round(pr/pt*100, 2) if pt else 0.0
    orr= round(po/pt*100, 2) if pt else 0.0

    # Curators match
    is_cur = (wbs, mu) in curator_set or (wbs, '') in curator_set

    # Pipeline match
    pl = pipeline.get(wbs, {})
    has_pl = bool(pl) and pt > 0

    # Designers link: any designer request with this WBS
    des_wbs_matches = [r for r in des_rows if r["wbs"] == wbs and wbs]

    # Industries / quarters / solutions can be multi-value
    inds = [i.strip() for i in c["industry"].split(",") if i.strip()]
    qtrs = [q.strip() for q in c["quarter"].split(",")   if q.strip()]
    sols = [s.strip() for s in c["solution"].split(",")  if s.strip()]

    AC.append({
        "campaign":    c["campaign"],
        "wbs":         wbs,
        "seq_id":      sid,
        "seq_link":    c["seq_link"],
        "mu":          mu,
        "source":      src,
        "solutions":   sols,
        "status":      c["status"],
        "priority":    c["priority"],
        "industries":  inds,
        "industry":    c["industry"],
        "quarters":    qtrs,
        "quarter":     c["quarter"],
        "executor":    c["executor"],
        "sde":         stats.get("owner",""),
        "target_pl":   c["target_pl"],
        "p_total":     pt,
        "p_replied":   pr,
        "p_opened":    po,
        "p_opted":     op,
        "e_deliver":   ed,
        "e_clicks":    ec,
        "reply_rate":  rr,
        "open_rate":   orr,
        "is_curated":  is_cur,
        "has_pipeline":has_pl,
        "pl_opps":     pl.get("opps",0),
        "pl_acv":      pl.get("acv",0.0),
        "has_designers": len(des_wbs_matches) > 0,
        "des_requests":  len(des_wbs_matches),
    })

print(f"  {len(AC)} campaigns with Outreach data")

# ══════════════════════════════════════════════════════════════════════════════
# 7. DESIGNERS SUMMARY STATS
# ══════════════════════════════════════════════════════════════════════════════
print("Computing Designers summary…")

total_requests = len(des_rows)
total_assets   = sum(r["n_assets"] for r in des_rows)
total_hours    = sum(r["hours"]    for r in des_rows)
matched_opp    = sum(1 for r in des_rows if r["has_opp"])

# Asset category distribution (by request count)
cat_dist = {}
for r in des_rows:
    cat = r["asset_cat"]
    cat_dist[cat] = cat_dist.get(cat, 0) + 1

# Asset type distribution (by request count, top 20)
type_dist = {}
for r in des_rows:
    t = r["asset_type"] or "Unknown"
    type_dist[t] = type_dist.get(t, 0) + 1
type_dist_sorted = dict(sorted(type_dist.items(), key=lambda x: -x[1])[:25])

# Status distribution
status_dist = {}
for r in des_rows:
    s = r["status"] or "Unknown"
    status_dist[s] = status_dist.get(s, 0) + 1

# Purpose distribution
purpose_dist = {}
for r in des_rows:
    p = r["purpose"] or "Unknown"
    purpose_dist[p] = purpose_dist.get(p, 0) + 1

# Region distribution
region_dist = {}
for r in des_rows:
    rg = r["region"] or "Unknown"
    region_dist[rg] = region_dist.get(rg, 0) + 1

# Opp pipeline rows matched from Designers
des_opp_rows = []
seen_opp_ids = set()
for r in des_rows:
    if not r["has_opp"] or not r["opp_id"]:
        continue
    if r["opp_id"] in seen_opp_ids:
        continue
    seen_opp_ids.add(r["opp_id"])
    pl_row = pipeline_by_opp.get(r["opp_id"], {})
    des_opp_rows.append({
        "opp_id":      r["opp_id"],
        "account":     r["pl_account"],
        "description": r["pl_desc"],
        "drm_cat":     r["pl_drm_cat"],
        "opp_status":  r["pl_opp_status"],
        "opp_phase":   r["pl_opp_phase"],
        "region":      r["pl_region"],
        "acv":         r["pl_acv"],
        "wbs":         r["wbs"],
        "project_name":r["name"],
        "asset_type":  r["asset_type"],
        "asset_cat":   r["asset_cat"],
        # Curators overlap: was this WBS in a curated campaign?
        "curators_overlap": any(
            c["is_curated"] for c in AC if c["wbs"] == r["wbs"] and r["wbs"]
        ) if r["wbs"] else False,
    })

# DRM category distribution of Designer opps
drm_dist = {}
for row in des_opp_rows:
    k = row["drm_cat"] or "Unknown"
    drm_dist[k] = drm_dist.get(k, 0) + 1

# Opp phase distribution
phase_dist = {}
for row in des_opp_rows:
    k = row["opp_phase"] or "Unknown"
    phase_dist[k] = phase_dist.get(k, 0) + 1

print(f"  {len(des_opp_rows)} unique opps matched from Designers")
print(f"  Curators overlap: {sum(1 for r in des_opp_rows if r['curators_overlap'])} opps")

# ══════════════════════════════════════════════════════════════════════════════
# 8. BUILD OUTPUT JSON
# ══════════════════════════════════════════════════════════════════════════════
print("Building JSON…")

out = {
    # Outreach campaigns (same structure as v2)
    "campaigns": AC,

    # Designers data
    "designers": {
        "kpis": {
            "total_requests": total_requests,
            "total_assets":   total_assets,
            "total_hours":    round(total_hours, 1),
            "matched_opp":    matched_opp,
            "pct_with_opp":   round(matched_opp / total_requests * 100, 1) if total_requests else 0,
            "unique_opps":    len(des_opp_rows),
        },
        "cat_dist":     cat_dist,
        "type_dist":    type_dist_sorted,
        "status_dist":  status_dist,
        "purpose_dist": purpose_dist,
        "region_dist":  region_dist,
        "drm_dist":     drm_dist,
        "phase_dist":   phase_dist,
        "opp_rows":     des_opp_rows,
    }
}

out_path = "/tmp/dash_v4.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(',',':'))

size_kb = pathlib.Path(out_path).stat().st_size / 1024
print(f"\nDone → {out_path}  ({size_kb:.0f} KB)")
print(f"  Campaigns: {len(AC)}")
print(f"  Designer rows: {len(des_rows)}  |  Opp rows: {len(des_opp_rows)}")
