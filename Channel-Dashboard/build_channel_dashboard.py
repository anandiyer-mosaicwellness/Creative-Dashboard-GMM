"""
PMax Channel/Inventory Level Dashboard Builder
Auto-discovers all weekly Excel files from Drive download folder → MM_Channel_Dashboard.html
To update: run gdown to sync Drive folder, then re-run this script.
"""

import os, json, glob, re
import pandas as pd
from datetime import datetime
from python_calamine.pandas import pandas_monkeypatch

pandas_monkeypatch()

# ── Config ────────────────────────────────────────────────────────────────────
DRIVE_DIR   = os.path.join(os.path.dirname(__file__), "Channel:Inventory Data", "Google-Channel Breakdown-Channel Split")
LEGACY_DIR  = os.path.join(os.path.dirname(__file__), "Channel:Inventory Data")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "MM_Channel_Dashboard.html")

DATA_DIR = DRIVE_DIR if os.path.isdir(DRIVE_DIR) else LEGACY_DIR

CHANNELS = ["Google Search", "YouTube", "Google Display Network",
            "Discover", "Gmail", "Maps", "Search partners"]

CHANNEL_COLORS = {
    "Google Search":          "#4285F4",
    "YouTube":                "#FF0000",
    "Google Display Network": "#34A853",
    "Discover":               "#FBBC05",
    "Gmail":                  "#EA4335",
    "Maps":                   "#1A73E8",
    "Search partners":        "#9E69AF",
}

NUMERIC_COLS = ["Cost", "Impr.", "Clicks", "Conversions",
                "Conv. value", "Cost / conv.", "Conv. value / cost"]

VALID_STATUSES = {"ACTIVE", "PAUSED"}
CATEGORIES     = ["Hair", "Beard", "Nutrition"]

def extract_category(name):
    m = re.search(r'[Pp]max[_\s]([A-Za-z]+)', str(name))
    if not m: return "Other"
    w = m.group(1).title()
    if w in ("Hair",): return "Hair"
    if w in ("Beard",): return "Beard"
    if w in ("Nutrition", "Nutr"): return "Nutrition"
    return "Other"

# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_date_range(raw):
    try:
        parts = str(raw).split(" - ")
        def fmt(d):
            dt = datetime.strptime(d.strip(), "%d %B %Y")
            return dt.strftime("%-d%b")
        return f"{fmt(parts[0])}-{fmt(parts[1])}"
    except Exception:
        return str(raw)

def clean_numeric(series):
    return pd.to_numeric(series.replace("--", None), errors="coerce").fillna(0)

def safe_json(v):
    if v is None or (isinstance(v, float) and (v != v)):
        return None
    return float(round(v, 2)) if isinstance(v, float) else v

# ── Load ──────────────────────────────────────────────────────────────────────
print(f"Scanning: {DATA_DIR}", flush=True)

candidates = []
seen_dates  = set()
for path in glob.glob(os.path.join(DATA_DIR, "*.xlsx")):
    raw = pd.read_excel(path, header=None, nrows=3, engine="calamine")
    date_str = str(raw.iloc[1, 0])
    try:
        start = datetime.strptime(date_str.split(" - ")[0].strip(), "%d %B %Y")
    except Exception:
        print(f"  SKIP (bad date): {os.path.basename(path)}", flush=True)
        continue
    if date_str in seen_dates:
        print(f"  SKIP (duplicate {date_str}): {os.path.basename(path)}", flush=True)
        continue
    seen_dates.add(date_str)
    candidates.append((start, date_str, path))

candidates.sort()
print(f"Found {len(candidates)} unique weeks", flush=True)

frames = []
week_labels = []

for wk_idx, (start, date_str, path) in enumerate(candidates):
    fname = os.path.basename(path)
    label = f"Wk{wk_idx+1} ({parse_date_range(date_str)})"
    week_labels.append(label)

    df = pd.read_excel(path, header=2, engine="calamine")
    df.rename(columns={"Channels": "Channel"}, inplace=True)
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = clean_numeric(df[col])

    # Reconstruct Conv. value from pre-calculated ROAS × Cost when missing
    if "Conv. value" not in df.columns and "Conv. value / cost" in df.columns:
        df["Conv. value"] = df["Conv. value / cost"] * df["Cost"]

    df["_week"]  = wk_idx
    df["_label"] = label
    df["_date"]  = date_str

    # Drop totals row and keep only valid channels + statuses
    df = df[df["Channel"].isin(CHANNELS)]
    df = df[df["Status"].isin(VALID_STATUSES)]
    df = df[df["Campaigns"].notna() & (df["Campaigns"] != "")]
    frames.append(df)
    print(f"  Wk{wk_idx+1}: {date_str} → {fname} ({len(df)} rows)", flush=True)

all_df = pd.concat(frames, ignore_index=True)

# Ensure numeric
for col in ["Cost", "Impr.", "Clicks", "Conversions", "Conv. value"]:
    if col in all_df.columns:
        all_df[col] = pd.to_numeric(all_df[col], errors="coerce").fillna(0)

campaigns = sorted(all_df["Campaigns"].dropna().unique().tolist())

# ── Build campaign-level raw data for JS (filtered client-side) ───────────────
# Columns: week, label, channel, campaign, status, cost, impr, clicks, conv, convVal
camp_rows = []
for _, r in all_df.iterrows():
    camp_rows.append([
        int(r["_week"]),        # 0 F_WK
        r["_label"],            # 1 F_LBL
        r["Channel"],           # 2 F_CH
        r["Campaigns"],         # 3 F_CAMP
        r["Status"],            # 4 F_ST
        float(round(r["Cost"], 2)),           # 5 F_COST
        int(r["Impr."]),                      # 6 F_IMPR
        int(r["Clicks"]),                     # 7 F_CLICKS
        float(round(r["Conversions"], 2)),    # 8 F_CONV
        float(round(r["Conv. value"], 2)),    # 9 F_CVAL
        extract_category(r["Campaigns"]),     # 10 F_CAT
    ])

# ── Latest week summary (all campaigns, active only — for initial render) ────
latest_wk  = all_df["_week"].max()
latest_lbl = all_df[all_df["_week"] == latest_wk]["_label"].iloc[0]

def agg_to_summary(df_sub, prev_df):
    rows = []
    grp = (df_sub.groupby("Channel", as_index=False)
           .agg(Cost=("Cost","sum"), Impressions=("Impr.","sum"),
                Clicks=("Clicks","sum"), Conversions=("Conversions","sum"),
                ConvValue=("Conv. value","sum")))
    prev_grp = (prev_df.groupby("Channel", as_index=False)
                .agg(Cost=("Cost","sum"), Conversions=("Conversions","sum"))
                .set_index("Channel")) if prev_df is not None and not prev_df.empty else None
    grp = grp.set_index("Channel")
    for ch in CHANNELS:
        if ch not in grp.index: continue
        r = grp.loc[ch]
        roas = float(round(r["ConvValue"]/r["Cost"], 2)) if r["Cost"] > 0 else 0
        cpa  = float(round(r["Cost"]/r["Conversions"], 2)) if r["Conversions"] > 0 else 0
        pc   = float(prev_grp.loc[ch,"Cost"])        if (prev_grp is not None and ch in prev_grp.index) else None
        pconv= float(prev_grp.loc[ch,"Conversions"]) if (prev_grp is not None and ch in prev_grp.index) else None
        wow_cost = round((float(r["Cost"])-pc)/pc*100,1) if pc else None
        wow_conv = round((float(r["Conversions"])-pconv)/pconv*100,1) if pconv else None
        rows.append({
            "channel": ch,
            "cost": float(round(r["Cost"],2)), "impressions": int(r["Impressions"]),
            "clicks": int(r["Clicks"]), "conversions": float(round(r["Conversions"],2)),
            "convValue": float(round(r["ConvValue"],2)), "roas": roas, "cpa": cpa,
            "costWoW": wow_cost, "convWoW": wow_conv,
        })
    return rows

prev_wk_df = all_df[all_df["_week"] == latest_wk - 1] if (latest_wk - 1) in all_df["_week"].values else None
summary_rows = agg_to_summary(
    all_df[all_df["_week"] == latest_wk],
    prev_wk_df
)

# ── JS blobs ──────────────────────────────────────────────────────────────────
channels_js    = json.dumps(CHANNELS)
colors_js      = json.dumps(CHANNEL_COLORS)
campaigns_js   = json.dumps(campaigns)
week_labels_js = json.dumps(week_labels)
categories_js  = json.dumps(CATEGORIES)
camp_rows_js   = json.dumps(camp_rows)
summary_js     = json.dumps(summary_rows)
latest_lbl_js  = json.dumps(latest_lbl)

# ── HTML ─────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>  <!-- v2 -->
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PMax Channel / Inventory Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:#0f1117; --card:#1a1d27; --border:#2a2d3a;
    --text:#e8eaf0; --muted:#8b8fa8; --accent:#4285F4;
  }}
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ background:var(--bg); color:var(--text); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; min-height:100vh; }}
  header {{ padding:22px 32px 14px; border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }}
  header h1 {{ font-size:1.3rem; font-weight:700; letter-spacing:-.3px; }}
  header span {{ font-size:.8rem; color:var(--muted); }}
  .toolbar {{ padding:12px 32px; display:flex; gap:20px; flex-wrap:wrap; align-items:center; border-bottom:1px solid var(--border); background:#13161f; }}
  .filter-group {{ display:flex; align-items:center; gap:6px; }}
  .filter-group label {{ font-size:.75rem; color:var(--muted); white-space:nowrap; }}
  select {{ background:#1e2133; color:var(--text); border:1px solid var(--border); border-radius:6px; padding:5px 10px; font-size:.8rem; cursor:pointer; }}
  select:focus {{ outline:none; border-color:var(--accent); }}
  .pill-group {{ display:flex; gap:3px; flex-wrap:wrap; }}
  .pill {{ padding:4px 12px; border-radius:999px; font-size:.76rem; cursor:pointer; border:1px solid var(--border); background:transparent; color:var(--muted); transition:all .12s; white-space:nowrap; }}
  .pill.cat-active      {{ background:#FBBC05; color:#111; border-color:#FBBC05; }}
  .pill.status-active   {{ background:#34A853; color:#fff; border-color:#34A853; }}
  .pill[data-status="PAUSED"].status-active {{ background:#8b8fa8; border-color:#8b8fa8; }}
  .pill[data-status="all"].status-active {{ background:var(--accent); border-color:var(--accent); }}
  /* Week selector */
  .week-selector {{ position:relative; }}
  .week-btn {{ background:#1e2133; color:var(--text); border:1px solid var(--border); border-radius:6px; padding:5px 11px; font-size:.8rem; cursor:pointer; white-space:nowrap; }}
  .week-btn.has-selection {{ border-color:var(--accent); color:var(--accent); }}
  .week-panel {{ position:absolute; top:calc(100% + 5px); left:0; background:#1a1d27; border:1px solid var(--border); border-radius:10px; padding:10px; z-index:200; min-width:230px; max-height:300px; display:none; flex-direction:column; gap:4px; box-shadow:0 8px 28px #00000099; }}
  .week-panel.open {{ display:flex; }}
  .week-panel-hdr {{ display:flex; justify-content:space-between; align-items:center; padding-bottom:7px; border-bottom:1px solid var(--border); margin-bottom:2px; }}
  .week-panel-hdr span {{ font-size:.72rem; color:var(--muted); }}
  .week-panel-hdr div {{ display:flex; gap:5px; }}
  .wk-action {{ font-size:.72rem; padding:3px 9px; border-radius:999px; border:1px solid var(--border); background:transparent; color:var(--muted); cursor:pointer; }}
  .wk-action:hover {{ border-color:var(--accent); color:var(--accent); }}
  .week-list {{ overflow-y:auto; display:flex; flex-direction:column; gap:1px; }}
  .week-check {{ display:flex; align-items:center; gap:8px; padding:5px 6px; border-radius:5px; cursor:pointer; font-size:.78rem; color:var(--text); }}
  .week-check:hover {{ background:#252836; }}
  .week-check input[type=checkbox] {{ cursor:pointer; accent-color:var(--accent); width:13px; height:13px; flex-shrink:0; }}
  /* Compare */
  .compare-pill {{ padding:4px 14px; border-radius:999px; font-size:.76rem; cursor:pointer; border:1px solid var(--border); background:transparent; color:var(--muted); transition:all .15s; white-space:nowrap; }}
  .compare-pill.active {{ background:#EA4335; color:#fff; border-color:#EA4335; }}
  .compare-note {{ font-size:.73rem; color:#EA4335; padding:4px 32px 0; }}
  /* Compare table */
  .cmp-wk-a {{ color:#60a5fa; }}
  .cmp-wk-b {{ color:#f97316; }}
  .cmp-delta-up {{ color:#34A853; font-size:.74rem; }}
  .cmp-delta-dn {{ color:#EA4335; font-size:.74rem; }}
  .cmp-delta-na {{ color:var(--muted); font-size:.74rem; }}
  .section {{ padding:0 32px 32px; }}
  .section-title {{ font-size:.85rem; font-weight:600; color:var(--muted); text-transform:uppercase; letter-spacing:.8px; margin-bottom:14px; padding-top:20px; }}
  /* Summary table */
  .summary-wrap {{ overflow-x:auto; background:var(--card); border:1px solid var(--border); border-radius:12px; }}
  table {{ width:100%; border-collapse:collapse; font-size:.82rem; }}
  thead tr {{ background:#12151e; }}
  th {{ padding:10px 14px; text-align:right; color:var(--muted); font-weight:500; white-space:nowrap; }}
  th:first-child {{ text-align:left; }}
  td {{ padding:9px 14px; text-align:right; border-top:1px solid var(--border); white-space:nowrap; }}
  td:first-child {{ text-align:left; font-weight:600; }}
  tr:hover td {{ background:#1f2235; }}
  .ch-dot {{ display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:7px; }}
  .wow-up {{ color:#34A853; font-size:.74rem; }}
  .wow-dn {{ color:#EA4335; font-size:.74rem; }}
  .wow-na {{ color:var(--muted); font-size:.74rem; }}
  /* Channel trend cards */
  .charts-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(460px,1fr)); gap:18px; }}
  .card {{ background:var(--card); border:1px solid var(--border); border-radius:12px; padding:18px 20px; }}
  .card-title {{ font-size:.88rem; font-weight:700; margin-bottom:6px; display:flex; align-items:center; gap:6px; }}
  .card-legend {{ display:flex; gap:14px; margin-bottom:10px; flex-wrap:wrap; }}
  .legend-item {{ display:flex; align-items:center; gap:5px; font-size:.72rem; color:var(--muted); }}
  .legend-line {{ width:18px; height:2px; border-radius:2px; }}
  .legend-dot  {{ width:8px; height:8px; border-radius:50%; }}
  .card canvas {{ width:100% !important; height:200px !important; }}
  .info-bar {{ font-size:.74rem; color:var(--muted); padding:5px 32px 8px; }}
  .no-data {{ grid-column:1/-1; padding:48px; text-align:center; color:var(--muted); font-size:.9rem; }}
</style>
</head>
<body>

<header>
  <h1>PMax — Channel / Inventory Level Dashboard</h1>
  <span id="latest-label"></span>
</header>

<div class="toolbar">
  <div class="filter-group">
    <label>Week</label>
    <div class="week-selector">
      <button class="week-btn" id="week-btn">All Weeks ▾</button>
      <div class="week-panel" id="week-panel">
        <div class="week-panel-hdr">
          <span id="wk-count-label">0 selected</span>
          <div>
            <button class="wk-action" id="wk-select-all">All</button>
            <button class="wk-action" id="wk-clear">Clear</button>
          </div>
        </div>
        <div class="week-list" id="week-list"></div>
      </div>
    </div>
    <button class="compare-pill" id="compare-btn">⇄ Compare</button>
  </div>
  <div class="filter-group">
    <label>Category</label>
    <div class="pill-group">
      <button class="pill cat-active" data-cat="all">All</button>
      <button class="pill" data-cat="Hair">Hair</button>
      <button class="pill" data-cat="Beard">Beard</button>
      <button class="pill" data-cat="Nutrition">Nutrition</button>
    </div>
  </div>
  <div class="filter-group">
    <label>Campaign</label>
    <div class="week-selector">
      <button class="week-btn" id="camp-btn">All Campaigns ▾</button>
      <div class="week-panel" id="camp-panel">
        <div class="week-panel-hdr">
          <span id="camp-count-label">0 selected</span>
          <div>
            <button class="wk-action" id="camp-select-all">All</button>
            <button class="wk-action" id="camp-clear">Clear</button>
          </div>
        </div>
        <div class="week-list" id="camp-list"></div>
      </div>
    </div>
  </div>
  <div class="filter-group">
    <label>Status</label>
    <div class="pill-group">
      <button class="pill status-active" data-status="all">All</button>
      <button class="pill" data-status="ACTIVE">Active</button>
      <button class="pill" data-status="PAUSED">Paused</button>
    </div>
  </div>
</div>
<div class="compare-note" id="compare-note" style="display:none"></div>
<div class="info-bar" id="info-bar"></div>

<div class="section">
  <div class="section-title" id="summary-title">Summary — Latest Week</div>
  <div class="summary-wrap">
    <table>
      <thead id="summary-head"></thead>
      <tbody id="summary-body"></tbody>
    </table>
  </div>
</div>

<div class="section">
  <div class="section-title">Weekly Trend — Channel Breakdown</div>
  <div class="charts-grid" id="charts-grid"></div>
</div>

<script>
// Row schema: [wk_idx, label, channel, campaign, status, cost, impr, clicks, conv, convVal, category]
const RAW         = {camp_rows_js};
const CHANNELS    = {channels_js};
const COLORS      = {colors_js};
const CAMPAIGNS   = {campaigns_js};
const WEEK_LABELS = {week_labels_js};
const CATEGORIES  = {categories_js};
const LATEST_LBL  = {latest_lbl_js};

const F_WK=0,F_LBL=1,F_CH=2,F_CAMP=3,F_ST=4,F_COST=5,F_IMPR=6,F_CLICKS=7,F_CONV=8,F_CVAL=9,F_CAT=10;

let activeCampaign = "all";
let activeStatus   = "all";
let activeCategory = "all";
let activeWeeks    = new Set(); // empty = all weeks
let compareMode    = false;
const chartInstances = {{}};

const FMT = {{
  cost: v => "₹"+v.toLocaleString("en-IN",{{maximumFractionDigits:0}}),
  conv: v => v.toLocaleString("en-IN",{{maximumFractionDigits:1}}),
  roas: v => v ? v.toFixed(2)+"x" : "–",
}};

document.getElementById("latest-label").textContent = "Latest week: " + LATEST_LBL;

// ── Week panel ─────────────────────────────────────────────────────────────────
const wkBtn   = document.getElementById("week-btn");
const wkPanel = document.getElementById("week-panel");
const wkList  = document.getElementById("week-list");

WEEK_LABELS.forEach((lbl, idx) => {{
  const row = document.createElement("label"); row.className="week-check";
  row.innerHTML = `<input type="checkbox" value="${{idx}}"><span>${{lbl}}</span>`;
  row.querySelector("input").addEventListener("change", e => {{
    e.target.checked ? activeWeeks.add(idx) : activeWeeks.delete(idx);
    updateWeekBtn();
    refresh();
  }});
  wkList.appendChild(row);
}});

document.getElementById("wk-select-all").addEventListener("click", () => {{
  wkList.querySelectorAll("input").forEach(cb => {{ cb.checked=true; activeWeeks.add(+cb.value); }});
  updateWeekBtn(); refresh();
}});
document.getElementById("wk-clear").addEventListener("click", () => {{
  wkList.querySelectorAll("input").forEach(cb => {{ cb.checked=false; }});
  activeWeeks.clear(); updateWeekBtn(); refresh();
}});

wkBtn.addEventListener("click", e => {{ e.stopPropagation(); wkPanel.classList.toggle("open"); }});
document.addEventListener("click", e => {{ if (!wkPanel.contains(e.target)) wkPanel.classList.remove("open"); }});

function updateWeekBtn() {{
  const n = activeWeeks.size;
  document.getElementById("wk-count-label").textContent = n ? `${{n}} selected` : "0 selected";
  wkBtn.textContent = n===0 ? "All Weeks ▾" : n===1 ? WEEK_LABELS[[...activeWeeks][0]]+" ▾" : `${{n}} Weeks ▾`;
  wkBtn.classList.toggle("has-selection", n>0);
}}

// ── Compare toggle ─────────────────────────────────────────────────────────────
const compareBtn  = document.getElementById("compare-btn");
const compareNote = document.getElementById("compare-note");
compareBtn.addEventListener("click", () => {{
  compareMode = !compareMode;
  compareBtn.classList.toggle("active", compareMode);
  refresh();
}});

function updateCompareNote() {{
  if (!compareMode) {{ compareNote.style.display="none"; return; }}
  const n = activeWeeks.size;
  if (n === 0) {{
    compareNote.style.display="block";
    compareNote.textContent = "⚠ Select at least 1 week to compare";
  }} else {{
    compareNote.style.display="none";
  }}
}}

// ── Campaign panel ─────────────────────────────────────────────────────────────
let activeCampaigns = new Set(); // empty = all

const campBtn   = document.getElementById("camp-btn");
const campPanel = document.getElementById("camp-panel");
const campList  = document.getElementById("camp-list");

function getCampOptions() {{
  // Only show campaigns that exist in the current status filter
  return [...new Set(
    RAW.filter(r =>
      (activeStatus   === "all" || r[F_ST]  === activeStatus) &&
      (activeCategory === "all" || r[F_CAT] === activeCategory)
    ).map(r => r[F_CAMP])
  )].sort();
}}

function rebuildCampPanel() {{
  const options = getCampOptions();
  // Drop any selected campaigns that no longer appear
  [...activeCampaigns].forEach(c => {{ if (!options.includes(c)) activeCampaigns.delete(c); }});
  campList.innerHTML = "";
  options.forEach(c => {{
    const row = document.createElement("label"); row.className = "week-check";
    row.innerHTML = `<input type="checkbox" value="${{c}}"${{activeCampaigns.has(c)?" checked":""}}><span>${{c}}</span>`;
    row.querySelector("input").addEventListener("change", e => {{
      e.target.checked ? activeCampaigns.add(c) : activeCampaigns.delete(c);
      updateCampBtn(); refresh();
    }});
    campList.appendChild(row);
  }});
  updateCampBtn();
}}

function updateCampBtn() {{
  const n = activeCampaigns.size;
  document.getElementById("camp-count-label").textContent = n ? `${{n}} selected` : "0 selected";
  campBtn.textContent = n === 0 ? "All Campaigns ▾" : n === 1 ? [...activeCampaigns][0].slice(0,22)+"… ▾" : `${{n}} Campaigns ▾`;
  campBtn.classList.toggle("has-selection", n > 0);
}}

campBtn.addEventListener("click", e => {{ e.stopPropagation(); campPanel.classList.toggle("open"); }});
document.addEventListener("click", e => {{ if (!campPanel.contains(e.target)) campPanel.classList.remove("open"); }});

document.getElementById("camp-select-all").addEventListener("click", () => {{
  campList.querySelectorAll("input").forEach(cb => {{ cb.checked=true; activeCampaigns.add(cb.value); }});
  updateCampBtn(); refresh();
}});
document.getElementById("camp-clear").addEventListener("click", () => {{
  campList.querySelectorAll("input").forEach(cb => {{ cb.checked=false; }});
  activeCampaigns.clear(); updateCampBtn(); refresh();
}});

rebuildCampPanel();

// ── Filter ─────────────────────────────────────────────────────────────────────
function filtered() {{
  return RAW.filter(r =>
    (activeWeeks.size===0    || activeWeeks.has(r[F_WK])) &&
    (activeCategory==="all"  || r[F_CAT]===activeCategory) &&
    (activeCampaigns.size===0 || activeCampaigns.has(r[F_CAMP])) &&
    (activeStatus==="all"    || r[F_ST]===activeStatus)
  );
}}

// ── Per-week totals across ALL channels (for % share) ──────────────────────────
function weekTotals(rows) {{
  const out={{}};
  rows.forEach(r=>{{
    if(!out[r[F_WK]]) out[r[F_WK]]={{cost:0,conv:0}};
    out[r[F_WK]].cost+=r[F_COST]; out[r[F_WK]].conv+=r[F_CONV];
  }});
  return out;
}}

// ── Top 3 campaigns ────────────────────────────────────────────────────────────
function top3(allRows, ch, wkIdx, metric) {{
  const m={{}};
  allRows.filter(r=>r[F_CH]===ch&&r[F_WK]===wkIdx).forEach(r=>{{
    if(!m[r[F_CAMP]]) m[r[F_CAMP]]={{cost:0,conv:0,convVal:0}};
    m[r[F_CAMP]].cost+=r[F_COST]; m[r[F_CAMP]].conv+=r[F_CONV]; m[r[F_CAMP]].convVal+=r[F_CVAL];
  }});
  const score = d => metric==="cac" ? (d.conv>0?d.cost/d.conv:Infinity) : metric==="conv" ? d.conv : d.cost;
  const sorted = Object.entries(m).sort((a,b)=>
    metric==="cac" ? score(a[1])-score(b[1]) : score(b[1])-score(a[1])
  );
  return sorted.slice(0,3);
}}

// ── Channel-week aggregation ───────────────────────────────────────────────────
function buildChannelSeries(rows) {{
  const out={{}};
  rows.forEach(r=>{{
    const ch=r[F_CH],wk=r[F_WK];
    if(!out[ch]) out[ch]={{}};
    if(!out[ch][wk]) out[ch][wk]={{label:r[F_LBL],cost:0,conv:0,convVal:0}};
    out[ch][wk].cost+=r[F_COST]; out[ch][wk].conv+=r[F_CONV]; out[ch][wk].convVal+=r[F_CVAL];
  }});
  return out;
}}

// ── Summary aggregation (single week) ─────────────────────────────────────────
function aggByChannel(rows) {{
  const m={{}};
  rows.forEach(r=>{{
    const ch=r[F_CH];
    if(!m[ch]) m[ch]={{cost:0,impr:0,clicks:0,conv:0,convVal:0}};
    m[ch].cost+=r[F_COST]; m[ch].impr+=r[F_IMPR]; m[ch].clicks+=r[F_CLICKS];
    m[ch].conv+=r[F_CONV]; m[ch].convVal+=r[F_CVAL];
  }});
  return m;
}}

function aggregateSummary(rows) {{
  const allWks=[...new Set(rows.map(r=>r[F_WK]))].sort((a,b)=>a-b);
  if(!allWks.length) return {{summaryRows:[],wkLabel:""}};
  const sumWk  = activeWeeks.size===1 ? [...activeWeeks][0] : allWks[allWks.length-1];
  const prevWk = sumWk-1;
  const wkLabel = WEEK_LABELS[sumWk]||"";
  const cur=aggByChannel(rows.filter(r=>r[F_WK]===sumWk));
  const prv=aggByChannel(rows.filter(r=>r[F_WK]===prevWk));
  const summaryRows=CHANNELS.map(ch=>{{
    const c=cur[ch]||null,p=prv[ch]||null;
    if(!c) return null;
    const roas=c.cost>0?+(c.convVal/c.cost).toFixed(2):0;
    const cac =c.conv>0?+(c.cost/c.conv).toFixed(2):0;
    const wc=(pv,cv)=>(p&&pv>0)?+((cv-pv)/pv*100).toFixed(1):null;
    return {{channel:ch,cost:c.cost,impressions:c.impr,clicks:c.clicks,
             conversions:c.conv,convValue:c.convVal,roas,cac,
             costWoW:wc(p?.cost,c.cost),convWoW:wc(p?.conv,c.conv)}};
  }}).filter(Boolean);
  return {{summaryRows,wkLabel}};
}}

// ── Render: normal summary ─────────────────────────────────────────────────────
function renderSummaryNormal(summaryRows, wkLabel) {{
  document.getElementById("summary-title").textContent = "Summary — "+(wkLabel||"Latest Week");
  document.getElementById("summary-head").innerHTML=`<tr>
    <th>Channel</th><th>Cost (₹)</th><th>Impressions</th><th>Clicks</th>
    <th>Conversions</th><th>Conv. Value (₹)</th><th>ROAS</th><th>CAC (₹)</th>
    <th>Cost WoW</th><th>Conv WoW</th></tr>`;
  const tbody=document.getElementById("summary-body");
  tbody.innerHTML="";
  if(!summaryRows.length){{
    tbody.innerHTML=`<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:24px">No data</td></tr>`;
    return;
  }}
  let tC=0,tI=0,tCl=0,tCo=0,tV=0;
  summaryRows.forEach(r=>{{
    tC+=r.cost;tI+=r.impressions;tCl+=r.clicks;tCo+=r.conversions;tV+=r.convValue;
    const wc=v=>v==null?"wow-na":v>=0?"wow-up":"wow-dn";
    const ws=v=>v==null?"–":(v>=0?"+":"")+v+"%";
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td><span class="ch-dot" style="background:${{COLORS[r.channel]}}"></span>${{r.channel}}</td>
      <td>${{FMT.cost(r.cost)}}</td><td>${{r.impressions.toLocaleString("en-IN")}}</td>
      <td>${{r.clicks.toLocaleString("en-IN")}}</td><td>${{FMT.conv(r.conversions)}}</td>
      <td>${{FMT.cost(r.convValue)}}</td><td>${{FMT.roas(r.roas)}}</td>
      <td>${{r.cac?"₹"+r.cac.toLocaleString("en-IN",{{maximumFractionDigits:0}}):"–"}}</td>
      <td class="${{wc(r.costWoW)}}">${{ws(r.costWoW)}}</td>
      <td class="${{wc(r.convWoW)}}">${{ws(r.convWoW)}}</td>`;
    tbody.appendChild(tr);
  }});
  const totalROAS=tC>0?(tV/tC).toFixed(2)+"x":"–";
  const tf=document.createElement("tr");
  tf.style.cssText="font-weight:700;background:#12151e";
  tf.innerHTML=`<td>Total</td><td>${{FMT.cost(tC)}}</td><td>${{tI.toLocaleString("en-IN")}}</td>
    <td>${{tCl.toLocaleString("en-IN")}}</td><td>${{FMT.conv(tCo)}}</td>
    <td>${{FMT.cost(tV)}}</td><td>${{totalROAS}}</td><td></td><td></td><td></td>`;
  tbody.appendChild(tf);
}}

// ── Render: compare summary (any N weeks vs previous N weeks) ──────────────────
function renderSummaryCompare(rows, selWks) {{
  // selWks: sorted array of selected week indices (Period A)
  // Period B: same count of weeks immediately before the earliest selected
  const n    = selWks.length;
  const minWk = selWks[0];
  const prevWks = Array.from({{length:n}}, (_,i) => minWk - n + i); // e.g. sel=[5,6,7] → prev=[2,3,4]

  const wkLabel = wks => {{
    if (wks.length===1) return WEEK_LABELS[wks[0]]||`Wk${{wks[0]+1}}`;
    const a=WEEK_LABELS[wks[0]]||`Wk${{wks[0]+1}}`, b=WEEK_LABELS[wks[wks.length-1]]||`Wk${{wks[wks.length-1]+1}}`;
    // Show just the week numbers for brevity in multi-week case
    return `Wk${{wks[0]+1}}–Wk${{wks[wks.length-1]+1}} (${{wks.length}}wk)`;
  }};

  const lblA = wkLabel(selWks);
  const lblB = wkLabel(prevWks);
  const prevAvail = prevWks.filter(w => w >= 0 && WEEK_LABELS[w]);

  document.getElementById("summary-title").textContent = `Compare: ${{lblA}} vs ${{lblB}}`;
  document.getElementById("summary-head").innerHTML=`<tr>
    <th>Channel</th>
    <th class="cmp-wk-a">${{lblA}} Cost</th>
    <th class="cmp-wk-b">${{lblB}} Cost</th>
    <th>Δ Cost</th>
    <th class="cmp-wk-a">${{lblA}} Conv</th>
    <th class="cmp-wk-b">${{lblB}} Conv</th>
    <th>Δ Conv</th>
    <th class="cmp-wk-a">${{lblA}} CAC</th>
    <th class="cmp-wk-b">${{lblB}} CAC</th>
    <th>Δ CAC</th>
  </tr>`;

  // Aggregate: Period A = selected weeks, Period B = previous N weeks from RAW (ignore campaign/week filter)
  const baseRows = RAW.filter(r =>
    (activeCategory==="all" || r[F_CAT]===activeCategory) &&
    (activeCampaigns.size===0 || activeCampaigns.has(r[F_CAMP])) &&
    (activeStatus==="all" || r[F_ST]===activeStatus)
  );
  const a = aggByChannel(baseRows.filter(r => selWks.includes(r[F_WK])));
  const b = aggByChannel(baseRows.filter(r => prevAvail.includes(r[F_WK])));

  const tbody=document.getElementById("summary-body"); tbody.innerHTML="";
  const delta=(av,bv)=>bv>0?+((av-bv)/bv*100).toFixed(1):null;
  const dCls =v=>v==null?"cmp-delta-na":v<=0?"cmp-delta-up":"cmp-delta-dn"; // CAC: lower=better
  const dClsN=v=>v==null?"cmp-delta-na":v>=0?"cmp-delta-up":"cmp-delta-dn";
  const ds   =v=>v==null?"–":(v>=0?"+":"")+v+"%";
  let totA={{cost:0,conv:0}}, totB={{cost:0,conv:0}};

  CHANNELS.forEach(ch=>{{
    const ra=a[ch], rb=b[ch];
    if(!ra&&!rb) return;
    const va=ra||{{cost:0,conv:0,convVal:0}}, vb=rb||{{cost:0,conv:0,convVal:0}};
    const cacA=va.conv>0?va.cost/va.conv:null, cacB=vb.conv>0?vb.cost/vb.conv:null;
    const dCost=delta(va.cost,vb.cost), dConv=delta(va.conv,vb.conv);
    const dCac=(cacA&&cacB)?+((cacA-cacB)/cacB*100).toFixed(1):null;
    totA.cost+=va.cost; totA.conv+=va.conv;
    totB.cost+=vb.cost; totB.conv+=vb.conv;
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td><span class="ch-dot" style="background:${{COLORS[ch]}}"></span>${{ch}}</td>
      <td class="cmp-wk-a">${{FMT.cost(va.cost)}}</td>
      <td class="cmp-wk-b">${{FMT.cost(vb.cost)}}</td>
      <td class="${{dClsN(dCost)}}">${{ds(dCost)}}</td>
      <td class="cmp-wk-a">${{FMT.conv(va.conv)}}</td>
      <td class="cmp-wk-b">${{FMT.conv(vb.conv)}}</td>
      <td class="${{dClsN(dConv)}}">${{ds(dConv)}}</td>
      <td class="cmp-wk-a">${{cacA?"₹"+Math.round(cacA):"–"}}</td>
      <td class="cmp-wk-b">${{cacB?"₹"+Math.round(cacB):"–"}}</td>
      <td class="${{dCls(dCac)}}">${{ds(dCac)}}</td>`;
    tbody.appendChild(tr);
  }});

  const dTotCost=delta(totA.cost,totB.cost), dTotConv=delta(totA.conv,totB.conv);
  const tf=document.createElement("tr"); tf.style.cssText="font-weight:700;background:#12151e";
  tf.innerHTML=`<td>Total</td>
    <td class="cmp-wk-a">${{FMT.cost(totA.cost)}}</td>
    <td class="cmp-wk-b">${{FMT.cost(totB.cost)}}</td>
    <td class="${{dClsN(dTotCost)}}">${{ds(dTotCost)}}</td>
    <td class="cmp-wk-a">${{FMT.conv(totA.conv)}}</td>
    <td class="cmp-wk-b">${{FMT.conv(totB.conv)}}</td>
    <td class="${{dClsN(dTotConv)}}">${{ds(dTotConv)}}</td>
    <td></td><td></td><td></td>`;
  tbody.appendChild(tf);

  if (prevAvail.length < n) {{
    const note=document.createElement("tr");
    note.innerHTML=`<td colspan="10" style="text-align:center;color:var(--muted);font-size:.72rem;padding:6px">
      ⚠ Previous period only has ${{prevAvail.length}} of ${{n}} weeks of data</td>`;
    tbody.appendChild(note);
  }}
}}

// ── External DOM tooltip (avoids canvas clipping) ─────────────────────────────
const C_CONV = "#34A853";
const C_CAC  = "#FBBC05";

const extTip = document.createElement("div");
extTip.style.cssText = [
  "position:fixed","pointer-events:none","z-index:9999",
  "background:#1a1d2e","border:1px solid #3a3d52","border-radius:10px",
  "padding:12px 15px","font-size:.78rem","color:#e8eaf0",
  "line-height:1.6","min-width:240px","max-width:360px",
  "box-shadow:0 10px 32px #000000bb","display:none",
].join(";");
document.body.appendChild(extTip);

function externalTooltip(context) {{
  const {{chart, tooltip}} = context;
  if (tooltip.opacity === 0) {{ extTip.style.display="none"; return; }}

  let html = `<div style="color:#8b8fa8;font-weight:600;margin-bottom:8px;font-size:.8rem">${{tooltip.title?.[0]||""}}</div>`;

  tooltip.dataPoints.forEach(dp => {{
    const ds  = dp.dataset;
    const idx = dp.dataIndex;
    const v   = dp.parsed.y;
    if (v === null || v === undefined) return;
    const wk  = ds._wkIndices?.[idx];
    const tot = ds._totals?.[wk] || {{}};

    let metricColor, metricLabel, metricVal;
    if (ds._metric === "cost") {{
      metricColor = chart.data.datasets[0].borderColor;
      const pct = tot.cost>0?(v/tot.cost*100).toFixed(1):"–";
      metricLabel = "Spend";
      metricVal   = `${{FMT.cost(v)}} <span style="color:#8b8fa8;font-size:.72rem">(${{pct}}% of total)</span>`;
    }} else if (ds._metric === "conv") {{
      metricColor = C_CONV;
      const pct = tot.conv>0?(v/tot.conv*100).toFixed(1):"–";
      metricLabel = "Conversions";
      metricVal   = `${{FMT.conv(v)}} <span style="color:#8b8fa8;font-size:.72rem">(${{pct}}% of total)</span>`;
    }} else {{
      metricColor = C_CAC;
      metricLabel = "CAC";
      metricVal   = FMT.cost(v);
    }}

    html += `<div style="color:${{metricColor}};font-weight:600;margin-top:4px">${{metricLabel}}: ${{metricVal}}</div>`;

    if (ds._ch && wk !== undefined) {{
      const camps = top3(ds._allRows, ds._ch, wk, ds._metric);
      if (camps.length) {{
        html += `<div style="color:#6b7080;font-size:.71rem;margin-top:3px;margin-bottom:1px">Top 3 campaigns</div>`;
        camps.forEach(([name, d], i) => {{
          const short = name.length>34 ? name.slice(0,33)+"…" : name;
          let val;
          if (ds._metric==="cost") val=FMT.cost(d.cost);
          else if (ds._metric==="conv") val=FMT.conv(d.conv);
          else val=d.conv>0?FMT.cost(d.cost/d.conv):"–";
          html += `<div style="font-size:.72rem;color:#b0b4cc">${{i+1}}. ${{short}} &nbsp;<strong style="color:#e8eaf0">${{val}}</strong></div>`;
        }});
      }}
    }}
    html += `<div style="height:6px"></div>`;
  }});

  extTip.innerHTML = html;
  extTip.style.display = "block";

  const rect = chart.canvas.getBoundingClientRect();
  let x = rect.left + tooltip.caretX + 14;
  let y = rect.top  + tooltip.caretY - 10;

  // prevent going off-screen right
  const tipW = extTip.offsetWidth;
  if (x + tipW + 10 > window.innerWidth) x = rect.left + tooltip.caretX - tipW - 14;
  // prevent going off-screen bottom
  const tipH = extTip.offsetHeight;
  if (y + tipH + 10 > window.innerHeight) y = window.innerHeight - tipH - 10;
  y = Math.max(10, y);

  extTip.style.left = x + "px";
  extTip.style.top  = y + "px";
}}

// ── Build all channel cards ───────────────────────────────────────────────────

function buildCharts(rows) {{
  Object.keys(chartInstances).forEach(k => {{ chartInstances[k].destroy(); delete chartInstances[k]; }});
  const grid = document.getElementById("charts-grid");
  grid.innerHTML = "";

  const chSeries = buildChannelSeries(rows);
  const totals   = weekTotals(rows);
  let hasAny = false;

  CHANNELS.forEach(ch => {{
    const wkMap = chSeries[ch];
    if (!wkMap || !Object.keys(wkMap).length) return;
    hasAny = true;

    const color  = COLORS[ch];
    const safeCh = ch.replace(/\W/g,"_");
    const wkIdxs = Object.keys(wkMap).map(Number).sort((a,b)=>a-b);
    const labels  = wkIdxs.map(i => wkMap[i].label);
    const costs   = wkIdxs.map(i => +wkMap[i].cost.toFixed(2));
    const convs   = wkIdxs.map(i => +wkMap[i].conv.toFixed(2));
    const cacList = wkIdxs.map(i => {{
      const d = wkMap[i];
      return d.conv>0 ? +(d.cost/d.conv).toFixed(2) : null;
    }});

    const card = document.createElement("div"); card.className="card";
    card.innerHTML = `
      <div class="card-title"><span class="ch-dot" style="background:${{color}}"></span>${{ch}}</div>
      <div class="card-legend">
        <div class="legend-item"><div class="legend-dot" style="background:${{color}}"></div>Spend (bars) — left</div>
        <div class="legend-item"><div class="legend-line" style="background:${{C_CONV}}"></div>Conversions — right</div>
        <div class="legend-item"><div class="legend-line" style="background:${{C_CAC}};border-bottom:2px dashed ${{C_CAC}};height:0"></div>CAC (₹) — right</div>
      </div>
      <canvas id="c-${{safeCh}}" width="460" height="200"></canvas>`;
    grid.appendChild(card);

    requestAnimationFrame(() => {{
      const ctx = document.getElementById("c-"+safeCh); if(!ctx) return;
      chartInstances[ch] = new Chart(ctx, {{
        type: "bar",
        data: {{
          labels,
          datasets: [
            {{
              type: "bar",
              label: "Spend",
              data: costs,
              backgroundColor: color+"99",
              borderColor: color,
              borderWidth: 1,
              borderRadius: 3,
              yAxisID: "y",
              _metric:"cost", _ch:ch, _wkIndices:wkIdxs, _allRows:rows, _totals:totals,
            }},
            {{
              type: "line",
              label: "Conversions",
              data: convs,
              borderColor: C_CONV,
              backgroundColor: "transparent",
              borderWidth: 2,
              pointRadius: 3, pointHoverRadius: 6,
              pointBackgroundColor: C_CONV,
              fill: false, tension: 0.3,
              yAxisID: "y1",
              _metric:"conv", _ch:ch, _wkIndices:wkIdxs, _allRows:rows, _totals:totals,
            }},
            {{
              type: "line",
              label: "CAC",
              data: cacList,
              borderColor: C_CAC,
              backgroundColor: "transparent",
              borderWidth: 2,
              pointRadius: 3, pointHoverRadius: 6,
              pointBackgroundColor: C_CAC,
              borderDash: [4,3],
              fill: false, tension: 0.3,
              yAxisID: "y1",
              _metric:"cac", _ch:ch, _wkIndices:wkIdxs, _allRows:rows, _totals:totals,
            }},
          ]
        }},
        options: {{
          responsive: false,
          maintainAspectRatio: false,
          animation: false,
          interaction: {{ mode:"index", intersect:false }},
          plugins: {{
            legend: {{ display:false }},
            tooltip: {{ enabled:false, external:externalTooltip }},
          }},
          scales: {{
            x: {{
              ticks: {{color:"#8b8fa8", font:{{size:8}}, maxRotation:30, autoSkip:true, maxTicksLimit:9}},
              grid: {{color:"#1e2133"}}
            }},
            y: {{
              position: "left",
              ticks: {{
                color:"#8b8fa8", font:{{size:8}}, maxTicksLimit:5,
                callback: v => v>=100000?"₹"+Math.round(v/1000)+"k" : v>=1000?"₹"+Math.round(v/1000)+"k" : "₹"+v
              }},
              grid: {{color:"#1e2133"}},
              title: {{display:true, text:"Spend (₹)", color:"#8b8fa8", font:{{size:8}}}}
            }},
            y1: {{
              position: "right",
              ticks: {{color:"#8b8fa8", font:{{size:8}}, maxTicksLimit:5}},
              grid: {{drawOnChartArea:false}},
              title: {{display:true, text:"Conv / CAC (₹)", color:"#8b8fa8", font:{{size:8}}}}
            }}
          }}
        }}
      }});
    }});
  }});

  if (!hasAny) {{
    grid.innerHTML = `<div class="no-data">No data for the selected filters.</div>`;
  }}
}}

// ── Info bar ──────────────────────────────────────────────────────────────────
function updateInfoBar(rows) {{
  const nCamps=new Set(rows.map(r=>r[F_CAMP])).size;
  const parts=[];
  if (activeWeeks.size>0) parts.push([...activeWeeks].map(i=>WEEK_LABELS[i]).join(", "));
  if (activeCategory!=="all") parts.push(activeCategory);
  if (activeCampaigns.size>0) parts.push(`${{activeCampaigns.size}} campaign${{activeCampaigns.size>1?"s":""}}`);
  else parts.push(`${{nCamps}} campaigns`);
  if (activeStatus!=="all") parts.push(activeStatus==="ACTIVE"?"Active":"Paused");
  document.getElementById("info-bar").textContent = parts.length?"Showing: "+parts.join(" · "):"";
}}

// ── Master refresh ─────────────────────────────────────────────────────────────
function refresh() {{
  const rows=filtered();
  updateCompareNote();
  // Summary
  if (compareMode && activeWeeks.size>=1) {{
    const selWks=[...activeWeeks].sort((a,b)=>a-b);
    renderSummaryCompare(rows,selWks);
  }} else {{
    const {{summaryRows,wkLabel}}=aggregateSummary(rows);
    renderSummaryNormal(summaryRows,wkLabel);
  }}
  buildCharts(rows);
  updateInfoBar(rows);
}}

refresh();

// ── Controls ───────────────────────────────────────────────────────────────────
document.querySelectorAll("[data-cat]").forEach(btn=>{{
  btn.addEventListener("click",()=>{{
    document.querySelectorAll("[data-cat]").forEach(b=>b.classList.remove("cat-active"));
    btn.classList.add("cat-active"); activeCategory=btn.dataset.cat;
    activeCampaigns.clear(); rebuildCampPanel();
    refresh();
  }});
}});

document.querySelectorAll("[data-status]").forEach(btn=>{{
  btn.addEventListener("click",()=>{{
    document.querySelectorAll("[data-status]").forEach(b=>b.classList.remove("status-active"));
    btn.classList.add("status-active"); activeStatus=btn.dataset.status;
    // Rebuild campaign list to match new status — clears invalid selections automatically
    activeCampaigns.clear(); rebuildCampPanel();
    refresh();
  }});
}});
</script>
</body>
</html>"""

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nDone → {OUTPUT_PATH}", flush=True)
