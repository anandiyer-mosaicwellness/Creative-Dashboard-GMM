import pandas as pd, json, warnings, os, re, time, io
import urllib.request
from collections import defaultdict
from datetime import datetime
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────────────────────
# DIMENSION DETECTION
# Fetches aspect ratio (9:16 or 16:9) for each YouTube video ID from YouTube.
# Results are cached in video_dims_cache.json — only new IDs are fetched.
# Safe: never touches spend/conversion data. Falls back to 'Unknown' on error.
# ─────────────────────────────────────────────────────────────────────────────
DIMS_CACHE_FILE = 'video_dims_cache.json'  # saved alongside build_dashboard.py

def load_dims_cache():
    if os.path.exists(DIMS_CACHE_FILE):
        with open(DIMS_CACHE_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_dims_cache(cache):
    with open(DIMS_CACHE_FILE, 'w') as f:
        json.dump(cache, f, indent=2)

def fetch_dimension(vid_id):
    """Returns '9:16', '16:9', or 'Unknown' for a YouTube video ID."""
    # Try /shorts/ oEmbed first — if valid response, it's a Short (9:16)
    # Then try /watch — if width > height it's 16:9, else 9:16
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    for url_type, url in [
        ('shorts', f'https://www.youtube.com/oembed?url=https://www.youtube.com/shorts/{vid_id}&format=json'),
        ('watch',  f'https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={vid_id}&format=json'),
    ]:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=6) as r:
                data = json.loads(r.read())
                w, h = data.get('width', 0), data.get('height', 0)
                if url_type == 'shorts' and w > 0:
                    return '9:16'   # responded to /shorts/ → it's a Short
                if url_type == 'watch' and w > 0 and h > 0:
                    return '9:16' if h > w else '16:9'
        except Exception:
            continue
    return 'Unknown'

def build_dims_map(video_ids):
    """Load cache, fetch missing IDs, save cache, return full map."""
    cache = load_dims_cache()
    missing = [vid for vid in video_ids if vid not in cache]
    if missing:
        print(f'  Fetching dimensions for {len(missing)} new video(s)...')
        for i, vid in enumerate(missing, 1):
            cache[vid] = fetch_dimension(vid)
            print(f'    [{i}/{len(missing)}] {vid} → {cache[vid]}')
            time.sleep(0.15)   # gentle rate limiting
        save_dims_cache(cache)
        print(f'  Dimensions cached → {DIMS_CACHE_FILE}')
    else:
        print(f'  All {len(video_ids)} video dimensions already cached.')
    return cache

UPLOAD_DIR = os.path.expanduser('~/Downloads')

def fetch_yt_from_gsheet(sheet_id, sheet_name):
    """Fetch YouTube video rows from a GViz sheet tab.
    Returns a list of dicts: {Asset, Campaign, AdGroup, Cost, Conversions}
    Columns: C=Asset, D=AssetType, E=Campaign, F=AdGroup, M=Cost, R=Conversions
    """
    from urllib.parse import quote as _quote
    tq = "select C,E,F,M,R where D='YouTube video' and M>0"
    url = (f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq'
           f'?tqx=out:json&sheet={_quote(sheet_name)}&tq={_quote(tq)}')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read().decode('utf-8')
        raw = raw[raw.index('{'):raw.rindex('}')+1]
        data = json.loads(raw)
    except Exception as e:
        print(f'  ⚠️  GViz fetch failed for "{sheet_name}": {e}')
        return []
    if data.get('status') == 'error':
        print(f'  ⚠️  GViz error for "{sheet_name}": {data.get("errors")}')
        return []
    rows = []
    for row in data['table']['rows']:
        v = [c.get('v') if c else None for c in row['c']]
        asset, camp, adg, cost, conv = v
        if not asset or not camp or not cost:
            continue
        adg = str(adg).strip() if adg and str(adg).strip() not in ('nan', 'None', '--') else ''
        rows.append({
            'Asset': str(asset), 'Campaign': str(camp), 'AdGroup': adg,
            'Cost': float(cost), 'Conversions': float(conv or 0)
        })
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# WEEK GROUPS — edit this section each time you add a new week
# Format: ("date range string", ["file1.xlsx", "file2.xlsx", ...])
# Files from BOTH accounts for the same week go in the same list → spend is summed
# ─────────────────────────────────────────────────────────────────────────────
WEEK_GROUPS = [
    ("23 February 2026 - 1 March 2026", [
        "Asset association report (1).xlsx",
    ]),
    ("2 March 2026 - 8 March 2026", [
        "Asset association report (2).xlsx",
    ]),
    ("9 March 2026 - 15 March 2026", [
        "Asset association report (3).xlsx",
    ]),
    ("16 March 2026 - 22 March 2026", [
        "Asset association report (4).xlsx",
    ]),
    ("23 March 2026 - 29 March 2026", [
        "Asset association report (5).xlsx",
    ]),
    ("30 March 2026 - 5 April 2026", [
        "Asset association report (6).xlsx",
        "Asset association report-631c8624.xlsx",       # 3.0 account
    ]),
    ("6 April 2026 - 12 April 2026", [
        "Asset association report (7).xlsx",
        "Asset association report (1)-bdd4c9cf.xlsx",   # 3.0 account
    ]),
    ("13 April 2026 - 19 April 2026", [
        "Asset association report (8).xlsx",
        "Asset association report (2)-873ee088.xlsx",   # 3.0 account
    ]),
    ("20 April 2026 - 26 April 2026", [
        "Asset association report (9).xlsx",
        "Asset association report (3)-2229d08d.xlsx",   # 3.0 account
    ]),
    ("27 April 2026 - 3 May 2026", [
        "Asset association report (10).xlsx",
        "Asset association report (4)-27ab6a25.xlsx",   # 3.0 account
    ]),
    ("4 May 2026 - 8 May 2026", [
        "Asset association report (11).xlsx",
    ]),
    ("4 May 2026 - 10 May 2026", [
        "Asset association report (12).xlsx",
        "Asset association report (5)-6ac2ff5d.xlsx",   # 3.0 account
    ]),
    ("11 May 2026 - 17 May 2026", [
        "Asset association report 2.0 .xlsx",           # main account
        "Asset association report 3.0 .xlsx",           # 3.0 account
    ]),
    # ── ADD NEW WEEKS BELOW THIS LINE ────────────────────────────────────────
    ("18 May 2026 - 24 May 2026", [
        "Asset association report (Week of 18th May) - 2.0 new.xlsx",       # main account
        "Asset association report (Week of 18th May) - 3.0-91456b62.xlsx", # 3.0 account
    ]),
    ("25 May 2026 - 31 May 2026", [
        "Asset association report - 2.0 (Week of 25th May).xlsx",           # main account
        "Asset association report - 3.0 (Week of 25th May).xlsx",           # 3.0 account
    ]),
    ("1 June 2026 - 7 June 2026", [
        "Asset association report (Week of 7th June) - 2.0.xlsx",           # main account
        "Asset association report (Week of 7th June) - 3.0.xlsx",           # 3.0 account
    ]),
    ("8 June 2026 - 14 June 2026", [
        "Asset association report (Week of 14th June) - 2.0.xlsx",          # main account
        "Asset association report (Week of 14th June) - 3.0.xlsx",          # 3.0 account
    ]),
    ("15 June 2026 - 21 June 2026", [
        "Asset association report (Week of 21st June) - 2.0.xlsx",          # main account
        "Asset association report (Week of 21st June) - 3.0.xlsx",          # 3.0 account
    ]),
    ("22 June 2026 - 28 June 2026", [
        "Asset association report (Week of 28th June) - 2.0.xlsx",          # main account
        "Asset association report (Week of 28th June) - 3.0.xlsx",          # 3.0 account
    ]),
]

# ─────────────────────────────────────────────────────────────────────────────
# GSHEET WEEKS — weeks fetched from Google Sheets (no Excel files needed)
# Format: ("date range string", "Sheet tab name")
# Add new weeks here every Monday going forward
# Sheet: https://docs.google.com/spreadsheets/d/1J6pog_JxT7oJ2jFA4koRw_0eIamcLQ9MYxnFJXrq-J4
# ─────────────────────────────────────────────────────────────────────────────
ASSET_SHEET_ID = '1J6pog_JxT7oJ2jFA4koRw_0eIamcLQ9MYxnFJXrq-J4'

GSHEET_WEEKS = [
    ("29 June 2026 - 5 July 2026",   "Week of 29th June"),
    ("6 July 2026 - 12 July 2026",   "Week of 6th July"),
    ("13 July 2026 - 19 July 2026",  "Week of 13th July"),
    # ── ADD NEW WEEKS BELOW THIS LINE ────────────────────────────────────────
    ("20 July 2026 - 26 July 2026",          "Week of 20th July"),
    ("27 July 2026 - 2 August 2026",         "Week of 27th July"),
    ("3 August 2026 - 9 August 2026",        "Week of 3rd August"),
    ("10 August 2026 - 16 August 2026",      "Week of 10th August"),
    ("17 August 2026 - 23 August 2026",      "Week of 17th August"),
    ("24 August 2026 - 30 August 2026",      "Week of 24th August"),
    ("31 August 2026 - 6 September 2026",    "Week of 31st August"),
    ("7 September 2026 - 13 September 2026", "Week of 7th September"),
    ("14 September 2026 - 20 September 2026","Week of 14th September"),
    ("21 September 2026 - 27 September 2026","Week of 21st September"),
    ("28 September 2026 - 4 October 2026",   "Week of 28th September"),
    ("5 October 2026 - 11 October 2026",     "Week of 5th October"),
    ("12 October 2026 - 18 October 2026",    "Week of 12th October"),
    ("19 October 2026 - 25 October 2026",    "Week of 19th October"),
]

# ─────────────────────────────────────────────────────────────────────────────
# CREATORS TAB — Consolidated sheet config
# Sheet: MM Centralised Sheet (same file used for creator data)
# Tab: "Help Sheet for YT Dashbaord" (note: typo in original sheet name)
#
# Structure (confirmed from actual file):
#   JFM section:  col 0=Category, 1=Name, 2=IG/Profile, 3=Platform
#   AMJ section:  col 5=Category, 6=Name, 7=Profile, 8=Platform
#   Category video sections (each group: Name, INT/INF, Narrative, CreatorType, YT Links):
#     Hair S2:           cols 12-16
#     Hair S3:           cols 18-22
#     Hair Biotin:       cols 24-28
#     Beard:             cols 30-34
#     Nutrition Shilajit: cols 36-40
#     Nutrition Creatine: cols 42-46
#     Nutrition Creatine: cols 48-52
#
# Script tries LOCAL FILE first (place xlsx in same folder), then Google Sheets.
# ─────────────────────────────────────────────────────────────────────────────
CREATOR_SHEET_ID  = '1hlngndF7Q7CpH135uNL629-SGP4MeH3tbIt1Gx106-c'
CREATOR_TAB_NAME  = 'Help Sheet for YT Dashbaord'   # note typo matches actual tab

# If you have the xlsx file locally, set this path. Set to None to always fetch from Google Sheets.
CREATOR_LOCAL_FILE = None   # e.g. 'MM Centralised Sheet _ Anand - Jan 26.xlsx'

# Category sections: (name_col, inf_col, narr_col, type_col, yt_col, category, product)
CREATOR_SECTIONS = [
    (12, 13, 14, 15, 16, 'Hair',      'S2 Kit'),
    (18, 19, 20, 21, 22, 'Hair',      'S3 Kit'),
    (24, 25, 26, 27, 28, 'Hair',      'Biotin'),
    (30, 31, 32, 33, 34, 'Beard',     'Beard'),
    (36, 37, 38, 39, 40, 'Nutrition', 'Shilajit'),
    (42, 43, 44, 45, 46, 'Nutrition', 'Creatine'),
    (48, 49, 50, 51, 52, 'Nutrition', 'Creatine'),
]

def fetch_gsheet_xlsx(sheet_id):
    """Fetch a Google Sheet as an in-memory xlsx file (all tabs)."""
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return io.BytesIO(r.read())
    except Exception as e:
        print(f'  ⚠️  Could not fetch sheet: {e}')
        return None

_YT_ID_RE = re.compile(r'^[A-Za-z0-9_-]{11}$')

def extract_vid_id(url_str):
    """Extract YouTube video ID from full URLs, shorts, youtu.be, or bare 11-char IDs."""
    if not isinstance(url_str, str) or not url_str.strip():
        return None
    url_str = url_str.strip()
    m = re.search(r'youtube\.com/shorts/([A-Za-z0-9_-]{11})', url_str)
    if m: return m.group(1)
    m = re.search(r'[?&]v=([A-Za-z0-9_-]{11})', url_str)
    if m: return m.group(1)
    m = re.search(r'youtu\.be/([A-Za-z0-9_-]{11})', url_str)
    if m: return m.group(1)
    # Bare 11-character YouTube video ID (e.g. "8GZrQOenSBo")
    if _YT_ID_RE.match(url_str):
        return url_str
    return None

CREATOR_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'creator_sheet_cache.xlsx')

def load_creator_xlsx():
    """Load creator xlsx — tries Google Sheets first (always latest data),
    falls back to local cache, then local file if Google Sheets is unreachable."""
    # 1. Try Google Sheets first — always has the most up-to-date creator data
    print(f'  Fetching creator sheet from Google Sheets...')
    buf = fetch_gsheet_xlsx(CREATOR_SHEET_ID)
    if buf is not None:
        # Save a local cache so future offline builds still work
        try:
            with open(CREATOR_CACHE_FILE, 'wb') as f:
                f.write(buf.getvalue())
            print(f'  Creator sheet cached locally → creator_sheet_cache.xlsx')
        except Exception:
            pass
        buf.seek(0)
        return buf
    # 2. Try explicit local file path as fallback
    if CREATOR_LOCAL_FILE and os.path.exists(CREATOR_LOCAL_FILE):
        print(f'  Google Sheets unavailable — using local file: {CREATOR_LOCAL_FILE}')
        with open(CREATOR_LOCAL_FILE, 'rb') as f:
            return io.BytesIO(f.read())
    # 3. Use local cache from last successful fetch
    if os.path.exists(CREATOR_CACHE_FILE):
        print(f'  Google Sheets unavailable — using cached creator sheet (creator_sheet_cache.xlsx)')
        with open(CREATOR_CACHE_FILE, 'rb') as f:
            return io.BytesIO(f.read())
    # 4. Auto-detect any centralised xlsx in current dir as final fallback
    for fname in os.listdir('.'):
        if fname.endswith('.xlsx') and 'centralised' in fname.lower():
            print(f'  Google Sheets unavailable — using local file: {fname}')
            with open(fname, 'rb') as f:
                return io.BytesIO(f.read())
    return None

def build_creator_map():
    """
    Parses Help Sheet for YT Dashbaord tab.
    Returns:
      video_to_creator: { vid_id: {name, inf_int, narrative, category, profile_url} }
      creator_profiles: { name: profile_url }
    """
    video_to_creator = {}
    creator_profiles = {}

    xlsx_buf = load_creator_xlsx()
    if xlsx_buf is None:
        print('  ⚠️  Creator data unavailable — Creators tab will show 0 creators')
        return video_to_creator, creator_profiles

    try:
        xl = pd.ExcelFile(xlsx_buf)
    except Exception as e:
        print(f'  ⚠️  Could not parse file: {e}')
        return video_to_creator, creator_profiles

    # Find the right tab
    target_sheet = None
    for sn in xl.sheet_names:
        if sn.strip().lower() == CREATOR_TAB_NAME.strip().lower():
            target_sheet = sn
            break
    if target_sheet is None:
        for sn in xl.sheet_names:
            if 'help' in sn.lower() and ('yt' in sn.lower() or 'dashb' in sn.lower()):
                target_sheet = sn
                break
    if target_sheet is None:
        print(f'  ⚠️  Tab not found. Available tabs: {xl.sheet_names}')
        return video_to_creator, creator_profiles

    print(f'  Reading tab: "{target_sheet}"')
    raw = xl.parse(target_sheet, header=None)
    data_rows = raw.iloc[4:]   # rows 0-3 are headers/section labels

    # ── Build creator profile map from JFM (cols 0-3) and AMJ (cols 5-8) ──
    # Platform = "Google" or "Youtube" → YT creator
    for _, row in data_rows.iterrows():
        for name_col, profile_col, platform_col in [(1, 2, 3), (6, 7, 8)]:
            if len(row) <= platform_col:
                continue
            name     = str(row.iloc[name_col]).strip()    if pd.notna(row.iloc[name_col])    else ''
            platform = str(row.iloc[platform_col]).strip().lower() if pd.notna(row.iloc[platform_col]) else ''
            profile  = str(row.iloc[profile_col]).strip() if pd.notna(row.iloc[profile_col]) else ''
            if (name and name.lower() not in ('nan', 'none', '')
                    and platform in ('google', 'youtube')
                    and profile and profile not in ('-', 'nan', 'none', '')):
                creator_profiles[name] = profile

    print(f'  YT creator profiles found: {len(creator_profiles)}')

    # ── Build video → creator map from each category section ──
    SKIP_NAMES = {'nan', 'none', '', 'yt creator', 'meta creator', 'int', 'inf', 'ai',
                  'internal', 'static', 'affluence', 'price slash', '-'}
    for name_col, inf_col, narr_col, type_col, yt_col, category, product in CREATOR_SECTIONS:
        last_name         = ''
        last_inf          = ''
        last_narr         = ''
        last_creator_type = ''
        for _, row in data_rows.iterrows():
            if yt_col >= len(row):
                continue
            raw_name = str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ''
            if raw_name.lower() not in SKIP_NAMES:
                last_name = raw_name
                last_inf  = str(row.iloc[inf_col]).strip()  if pd.notna(row.iloc[inf_col])  else ''
                last_narr = str(row.iloc[narr_col]).strip() if pd.notna(row.iloc[narr_col]) else ''
            # Creator Type is on each row (not just name rows) — capture it per video row
            raw_type = str(row.iloc[type_col]).strip() if pd.notna(row.iloc[type_col]) else ''
            if raw_type in ('YT Creator', 'Meta Creator'):
                last_creator_type = raw_type
            elif not raw_name.lower() in SKIP_NAMES and raw_name:
                last_creator_type = ''  # reset when name changes and no type set
            if not last_name:
                continue
            yt_raw = str(row.iloc[yt_col]).strip() if pd.notna(row.iloc[yt_col]) else ''
            vid_id = extract_vid_id(yt_raw)
            if not vid_id:
                continue
            # Blank creator_type → Meta Creator (per user's rule)
            ct = last_creator_type if last_creator_type else 'Meta Creator'
            video_to_creator[vid_id] = {
                'name':         last_name,
                'inf_int':      last_inf,
                'narrative':    last_narr,
                'category':     category,
                'product':      product,
                'creator_type': ct,
                'profile_url':  creator_profiles.get(last_name, ''),
            }

    # ── Also read BQ–BX section (cols 68–75) for creator info ───────────────
    # col68=Category, col69=Product, col70=Inf/Int, col71=Name, col72=Narrative,
    # col73=Creator Type, col74=YT Link, col75=Date Taken Live
    if raw.shape[1] >= 76:
        for _, row in data_rows.iterrows():
            yt_raw = str(row.iloc[74]).strip() if pd.notna(row.iloc[74]) else ''
            vid_id = extract_vid_id(yt_raw)
            if not vid_id or vid_id in video_to_creator:
                continue
            name = str(row.iloc[71]).strip() if pd.notna(row.iloc[71]) else ''
            if not name or name.lower() in SKIP_NAMES:
                continue
            cat_raw = str(row.iloc[68]).strip() if pd.notna(row.iloc[68]) else ''
            prod_raw = str(row.iloc[69]).strip() if pd.notna(row.iloc[69]) else ''
            inf_int = str(row.iloc[70]).strip() if pd.notna(row.iloc[70]) else 'Inf'
            ct_raw  = str(row.iloc[73]).strip() if pd.notna(row.iloc[73]) else 'YT Creator'
            ct = ct_raw if ct_raw in ('YT Creator', 'Meta Creator') else 'YT Creator'
            video_to_creator[vid_id] = {
                'name':         name,
                'inf_int':      inf_int,
                'narrative':    str(row.iloc[72]).strip() if pd.notna(row.iloc[72]) else '',
                'category':     cat_raw or 'All Products',
                'product':      prod_raw,
                'creator_type': ct,
                'profile_url':  creator_profiles.get(name, ''),
            }
    # ── Also load 16:9 video→creator mappings from JSON fallback ──────────────
    # These are 16:9 format videos from BC-BI columns of the Help Sheet.
    # They're NOT in the category sections (12-52), so without this step they
    # wouldn't appear in the Creators tab at all.
    json_16_9 = '16_9_creators.json'
    if os.path.exists(json_16_9):
        try:
            with open(json_16_9) as f:
                raw = json.load(f)
            bc_bi = raw.get('video_to_creator', {})
            added = 0
            for vid_id, info in bc_bi.items():
                if vid_id not in video_to_creator:  # don't overwrite existing mapping
                    name = info.get('name', '')
                    if not name:
                        continue
                    # Prefer profile URL from creator_profiles (JFM/AMJ), fall back to JSON
                    profile = creator_profiles.get(name, '') or info.get('profile_url', '')
                    video_to_creator[vid_id] = {
                        'name':         name,
                        'inf_int':      info.get('inf_int', 'Inf'),
                        'narrative':    info.get('narrative', ''),
                        'category':     info.get('category', ''),
                        'creator_type': info.get('creator_type', 'YT Creator'),
                        'profile_url':  profile,
                    }
                    # Also update creator_profiles if we have a URL
                    if profile and name not in creator_profiles:
                        creator_profiles[name] = profile
                    added += 1
            print(f'  Added {added} 16:9 video→creator mappings from {json_16_9}')
        except Exception as e:
            print(f'  ⚠️  Could not load 16:9 creator JSON: {e}')

    unique_creators = len(set(v['name'] for v in video_to_creator.values()))
    print(f'  Video-to-creator map: {len(video_to_creator)} videos → {unique_creators} creators')
    return video_to_creator, creator_profiles

# ─────────────────────────────────────────────────────────────────────────────
# 16:9 Int VIDEO FILTER
# Reads BC-BI (cols 54-60) of Help Sheet for YT Dashbaord.
# Returns set of video IDs where INT/INF (col 56 = BE) = "Int".
# REVERT NOTE: to remove this feature, delete this function and the call
#   to it in section 2c, the 'fmt' field on videos, and the sFMT filter
#   in the HTML + getF() + render() + clearFilters().
# ─────────────────────────────────────────────────────────────────────────────
def _extract_vid_id_from_row_range(row, start_col, end_col):
    """
    Scan columns [start_col, end_col) in a single row and return the first
    YouTube video ID found — handles full URLs, shorts, youtu.be, and bare 11-char IDs.
    """
    for c in range(start_col, end_col):
        if c >= len(row): break
        cell = row.iloc[c]
        if not pd.notna(cell): continue
        cell_str = str(cell).strip()
        if cell_str in ('', 'nan', '-', 'None'): continue
        vid = extract_vid_id(cell_str)
        if vid:
            return vid
    return None


def build_16_9_int_map():
    """
    Returns set of video IDs from the '16:9 Video All Product' section (BQ–BX, cols 68–75).
    BX (col 75) = Date Taken Live. Video ID scanned across all of BQ–BW (68–74) per row,
    handling full YouTube URLs, youtu.be links, and bare 11-char video IDs.
    Falls back to 16_9_video_ids.json if sheet not available / too narrow.
    """
    ids = set()

    xlsx_buf = load_creator_xlsx()
    if xlsx_buf is not None:
        try:
            xl = pd.ExcelFile(xlsx_buf)
            target_sheet = None
            for sn in xl.sheet_names:
                if sn.strip().lower() == CREATOR_TAB_NAME.strip().lower():
                    target_sheet = sn
                    break
            if target_sheet:
                raw = xl.parse(target_sheet, header=None)
                # BQ=68 … BX=75  → need at least 76 cols
                if raw.shape[1] >= 76:
                    for _, row in raw.iloc[4:].iterrows():
                        vid_id = _extract_vid_id_from_row_range(row, 68, 75)
                        if vid_id:
                            ids.add(vid_id)
                    print(f'  16:9 videos found from sheet (BQ–BX scan): {len(ids)}')
                    return ids
                else:
                    print(f'  ⚠️  16:9 section (BQ–BX) not in xlsx ({raw.shape[1]} cols) — trying JSON fallback')
        except Exception as e:
            print(f'  ⚠️  Could not parse xlsx for 16:9: {e}')

    json_path = '16_9_video_ids.json'
    if os.path.exists(json_path):
        try:
            with open(json_path) as f:
                data = json.load(f)
            ids = set(data.get('video_ids', []))
            print(f'  16:9 videos loaded from JSON fallback: {len(ids)}')
            return ids
        except Exception as e:
            print(f'  ⚠️  Could not read 16:9 JSON: {e}')

    print('  ⚠️  No 16:9 data found. "16:9 — Int" filter will show 0 results.')
    return ids


def build_16_9_dates():
    """
    Reads BQ–BX (cols 68–75) of Help Sheet.
    BX (col 75) = 'Date Taken Live' → rights start date.
    Video ID scanned across all BQ–BW (68–74) per row — handles full URLs and bare IDs.
    Returns {vid_id: 'DD Mon YYYY'}.
    """
    import datetime
    dates = {}
    months_out = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    xlsx_buf = load_creator_xlsx()
    if xlsx_buf is None:
        print('  ⚠️  build_16_9_dates: no xlsx available')
        return dates
    try:
        xl = pd.ExcelFile(xlsx_buf)
        target_sheet = None
        for sn in xl.sheet_names:
            if sn.strip().lower() == CREATOR_TAB_NAME.strip().lower():
                target_sheet = sn
                break
        if not target_sheet:
            return dates
        raw = xl.parse(target_sheet, header=None)
        if raw.shape[1] < 76:
            print(f'  ⚠️  build_16_9_dates: sheet only {raw.shape[1]} cols, need 76+ for BQ–BX')
            return dates

        date_col = 75  # BX = "Date Taken Live" (confirmed by user)
        print(f'  16:9 section: scanning BQ–BW (cols 68–74) for video IDs, BX (col 75) for date')

        # Diagnostic: print first 3 non-empty rows from BQ–BX so we can see raw values
        diag_count = 0
        for ri in range(4, min(raw.shape[0], 60)):
            row_diag = raw.iloc[ri]
            if pd.notna(row_diag.iloc[75]) or any(pd.notna(row_diag.iloc[c]) for c in range(68,75)):
                cells = {f'col{c}': str(row_diag.iloc[c]) for c in range(68, 76) if c < raw.shape[1]}
                print(f'    row{ri} BQ–BX: {cells}')
                diag_count += 1
                if diag_count >= 3:
                    break

        for _, row in raw.iloc[4:].iterrows():
            # Must have a date in BX first
            raw_date = row.iloc[date_col] if pd.notna(row.iloc[date_col]) else None
            if raw_date is None:
                continue
            # Scan BQ–BW for a video ID (URL, youtu.be, or bare 11-char ID)
            vid_id = _extract_vid_id_from_row_range(row, 68, 75)
            if not vid_id:
                continue
            # Parse date — pandas may already give a Timestamp, or it's a string
            try:
                if hasattr(raw_date, 'year'):
                    d = raw_date
                    dates[vid_id] = f"{str(d.day).zfill(2)} {months_out[d.month-1]} {d.year}"
                else:
                    s = str(raw_date).strip()
                    parsed = False
                    for fmt in ('%d %b %Y','%d %B %Y','%Y-%m-%d','%d/%m/%Y','%m/%d/%Y','%d-%m-%Y','%d.%m.%Y'):
                        try:
                            d = datetime.datetime.strptime(s, fmt)
                            dates[vid_id] = f"{str(d.day).zfill(2)} {months_out[d.month-1]} {d.year}"
                            parsed = True
                            break
                        except ValueError:
                            pass
                    if not parsed:
                        print(f'    ⚠️  Unrecognised date format for {vid_id}: "{s}"')
            except Exception as ex:
                print(f'    ⚠️  Date parse error for {vid_id}: {ex}')

        print(f'  16:9 "Date Taken Live" dates extracted: {len(dates)}')
    except Exception as e:
        print(f'  ⚠️  build_16_9_dates error: {e}')
    return dates


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def get_ct(camp):
    for ct in ['Pmax', 'DG', 'UAC', 'Branding']:
        if str(camp).startswith(ct): return ct
    return str(camp).split('_')[0].split('-')[0]

def get_category(camp):
    c = str(camp).lower()
    if any(x in c for x in ['hair', 'dandruff', 'biotin']):               return 'Hair'
    if 'beard' in c:                                                        return 'Beard'
    if any(x in c for x in ['nutrition', 'nutri', 'shilajit', 'creatine',
                              'protein', 'magnesium', 'ibk', 'mvg',
                              'gummies', 'vitamin', 'weight', 'muscle']):  return 'Nutrition'
    return 'All Products'

def fmt_wk(i, dr):
    parts = dr.split(' - ')
    def sh(s):
        t = s.strip().split(' ')
        return t[0] + t[1][:3] if len(t) >= 2 else s
    return f'Wk{i} ({sh(parts[0])}-{sh(parts[1])})'

# ─────────────────────────────────────────────────────────────────────────────
# 1. LOAD & MERGE DATA BY WEEK GROUP
# ─────────────────────────────────────────────────────────────────────────────
week_strs, week_labels, frames = [], [], []

all_week_groups = [(date_str, file_list, 'xlsx') for date_str, file_list in WEEK_GROUPS] + \
                  [(date_str, tab_name, 'gsheet') for date_str, tab_name in GSHEET_WEEKS]

for wk_i, (date_str, source, source_type) in enumerate(all_week_groups, 1):
    week_strs.append(date_str)
    week_labels.append(fmt_wk(wk_i, date_str))

    if source_type == 'gsheet':
        print(f'  Fetching from Google Sheet: {source}...')
        gsheet_rows = fetch_yt_from_gsheet(ASSET_SHEET_ID, source)
        print(f'    {len(gsheet_rows)} YouTube video rows')
        if gsheet_rows:
            df = pd.DataFrame(gsheet_rows)
            df['Week'] = date_str
            df['CT']       = df['Campaign'].apply(get_ct)
            df['Category'] = df['Campaign'].apply(get_category)
            df = df.rename(columns={'AdGroup': 'Ad group'})
            frames.append(df[['Week', 'Asset', 'Campaign', 'CT', 'Category', 'Ad group', 'Cost', 'Conversions']])
    else:
        for fname in source:
            fpath = os.path.join(UPLOAD_DIR, fname)
            if not os.path.exists(fpath):
                print(f'  ⚠️  File not found: {fname} — skipping')
                continue
            # Auto-detect header row (some files have 2 pre-header rows, others start at row 0)
            raw = pd.read_excel(fpath, sheet_name=0, header=None)
            hdr_row = next((i for i, row in raw.iterrows() if 'Asset type' in row.values), 2)
            df = pd.read_excel(fpath, sheet_name=0, header=hdr_row)
            df = df[df['Asset type'] == 'YouTube video'].copy()
            df['Week'] = date_str
            df['CT']       = df['Campaign'].apply(get_ct)
            df['Category'] = df['Campaign'].apply(get_category)
            df['Conversions'] = pd.to_numeric(df['Conversions'], errors='coerce').fillna(0)
            frames.append(df[['Week', 'Asset', 'Campaign', 'CT', 'Category', 'Ad group', 'Cost', 'Conversions']])

combined = pd.concat(frames, ignore_index=True)
combined.columns = ['Week', 'Asset', 'Campaign', 'CT', 'Category', 'AdGroup', 'Cost', 'Conversions']

wk_idx = {w: i for i, w in enumerate(week_strs)}

# ─────────────────────────────────────────────────────────────────────────────
# 2. BUILD PER-VIDEO COMBO SPEND
# ─────────────────────────────────────────────────────────────────────────────
vdata = defaultdict(lambda: defaultdict(lambda: {
    'w':  [0.0] * len(week_strs),
    'wc': [0.0] * len(week_strs),
}))

for _, r in combined.iterrows():
    wi = wk_idx.get(r.Week)
    if wi is not None and r.Cost > 0:
        key = (r.CT, r.Category, r.Campaign, r.AdGroup)
        vdata[r.Asset][key]['w'][wi]  += float(r.Cost)
        vdata[r.Asset][key]['wc'][wi] += float(r.Conversions)

videos = []
for url, combos in vdata.items():
    vid_id = url.split('v=')[-1] if 'v=' in url else url
    combo_list = [
        {'ct': ct, 'cat': cat, 'c': camp, 'ag': ag, 'w': v['w'], 'wc': v['wc']}
        for (ct, cat, camp, ag), v in combos.items()
        if sum(v['w']) > 0
    ]
    total = sum(sum(c['w']) for c in combo_list)
    if total > 0:
        videos.append({'url': url, 'id': vid_id, 'total': round(total, 2), 'combos': combo_list})

videos.sort(key=lambda x: x['total'], reverse=True)

# ─────────────────────────────────────────────────────────────────────────────
# 2b. FETCH / CACHE VIDEO DIMENSIONS
# ─────────────────────────────────────────────────────────────────────────────
print('Checking video dimensions...')
all_vid_ids = [v['id'] for v in videos]
dims_map = build_dims_map(all_vid_ids)
for v in videos:
    v['dim'] = dims_map.get(v['id'], 'Unknown')

# ─────────────────────────────────────────────────────────────────────────────
# 2c. BUILD CREATORS TAB DATA + 16:9 Int tags
# ─────────────────────────────────────────────────────────────────────────────
print('Building Creators Tab data...')
video_to_creator, creator_profiles = build_creator_map()

# Tag each video with its format (16:9 Int or Others) for the Tab 1 filter
print('Building 16:9 Int video map...')
vid_16_9_int = build_16_9_int_map()
for v in videos:
    v['fmt'] = '16:9 Int' if v['id'] in vid_16_9_int else 'Others'

# vid_id → video object (for quick lookup)
vid_lookup = {v['id']: v for v in videos}

# Aggregate per creator
creator_agg = {}   # name → { meta, w[], wc[], videos[] }
for vid_id, cinfo in video_to_creator.items():
    vobj = vid_lookup.get(vid_id)
    if not vobj:
        continue   # video has no spend data — skip
    name = cinfo['name']
    if name not in creator_agg:
        creator_agg[name] = {
            'name':         name,
            'profile_url':  cinfo.get('profile_url', ''),
            'creator_type': cinfo.get('creator_type', 'Meta Creator'),
            'w':  [0.0] * len(week_strs),
            'wc': [0.0] * len(week_strs),
            'videos': [],
        }
    ca = creator_agg[name]
    # If any video is tagged YT Creator, the creator is YT Creator
    if cinfo.get('creator_type') == 'YT Creator':
        ca['creator_type'] = 'YT Creator'
    # Sum weekly spend/conv across all combos for this video
    vid_w  = [0.0] * len(week_strs)
    vid_wc = [0.0] * len(week_strs)
    for combo in vobj['combos']:
        for i in range(len(week_strs)):
            vid_w[i]  += combo['w'][i]
            vid_wc[i] += combo['wc'][i]
    # Add to creator totals
    for i in range(len(week_strs)):
        ca['w'][i]  += vid_w[i]
        ca['wc'][i] += vid_wc[i]
    # Add video entry
    ca['videos'].append({
        'id':        vid_id,
        'url':       vobj['url'],
        'narrative': cinfo['narrative'],
        'product':   cinfo.get('product', ''),
        'total':     round(sum(vid_w), 2),
        'w':         [round(x, 2) for x in vid_w],
        'wc':        [round(x, 2) for x in vid_wc],
    })

# Sort creator videos by total spend desc, sort creators by total spend desc
# Derive category from the creator's videos' campaign categories (majority by spend)
creators_list = []
for ca in creator_agg.values():
    ca['videos'].sort(key=lambda v: v['total'], reverse=True)
    ca['total'] = round(sum(ca['w']), 2)
    ca['w']  = [round(x, 2) for x in ca['w']]
    ca['wc'] = [round(x, 2) for x in ca['wc']]
    if ca['total'] == 0:
        continue
    # Derive top category by spend across all video combos
    cat_spend = defaultdict(float)
    for v in ca['videos']:
        vobj = vid_lookup.get(v['id'])
        if vobj:
            for combo in vobj['combos']:
                cat_spend[combo['cat']] += sum(combo['w'])
    ca['category'] = max(cat_spend, key=cat_spend.get) if cat_spend else 'All Products'
    creators_list.append(ca)

creators_list.sort(key=lambda c: c['total'], reverse=True)
print(f'  Creators with spend data: {len(creators_list)}')

creators_data = {
    'meta': {
        'generated':   datetime.now().strftime('%d %b %Y'),
        'week_labels': week_labels,
        'week_range':  f"{week_strs[0].split(' - ')[0]} → {week_strs[-1].split(' - ')[1]}",
    },
    'creators': creators_list,
}
creators_json_str = json.dumps(creators_data, separators=(',', ':'))

# ─────────────────────────────────────────────────────────────────────────────
# 3. BUILD DROPDOWN OPTIONS
# ─────────────────────────────────────────────────────────────────────────────
all_cts   = sorted({c['ct']  for v in videos for c in v['combos']})
all_cats  = sorted({c['cat'] for v in videos for c in v['combos']})
all_camps = sorted({c['c']   for v in videos for c in v['combos']})
all_ags   = sorted({c['ag']  for v in videos for c in v['combos']})

ags_by_camp = {'All': ['All'] + all_ags}
for camp in all_camps:
    ags_by_camp[camp] = ['All'] + sorted({c['ag'] for v in videos for c in v['combos'] if c['c'] == camp})

data = {
    'meta': {
        'generated':   datetime.now().strftime('%d %b %Y'),
        'total_videos': len(videos),
        'week_labels':  week_labels,
        'week_strs':    week_strs,
        'week_range':   f"{week_strs[0].split(' - ')[0]} → {week_strs[-1].split(' - ')[1]}"
    },
    'opts': {
        'cts':         ['All'] + all_cts,
        'cats':        ['All'] + all_cats,
        'ags_by_camp': ags_by_camp,
    },
    'videos': videos
}

json_str = json.dumps(data, separators=(',', ':'))

# ─────────────────────────────────────────────────────────────────────────────
# 4. HTML TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MM — YouTube Creative Dashboard</title>
<style>
:root{
  --navy:#1F3864;--blue:#2E75B6;--bg:#F0F4F9;
  --white:#fff;--border:#DDE4EE;--text:#1a202c;--muted:#718096;
  --g1:#0e6b2f;--g5:#1a9940;--g15:#5cb85c;--y30:#f0ad4e;--o50:#e67e22;--r99:#c0392b;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
.hdr{background:var(--navy);color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;gap:12px}
.hdr-left h1{font-size:20px;font-weight:700;letter-spacing:.3px}
.hdr-left .sub{font-size:13px;opacity:.7;margin-top:3px}
.hdr-badge{background:rgba(255,255,255,.15);padding:5px 14px;border-radius:20px;font-size:13px;white-space:nowrap}
.fbar{background:var(--white);border-bottom:1px solid var(--border);padding:12px 24px;
      display:flex;align-items:center;gap:14px;flex-wrap:wrap;
      position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.08)}
.fg{display:flex;align-items:center;gap:7px}
.flabel{font-size:12px;font-weight:700;color:var(--navy);white-space:nowrap;text-transform:uppercase;letter-spacing:.5px}
select{padding:7px 26px 7px 10px;border:2px solid var(--border);border-radius:6px;
       font-size:13px;color:var(--text);cursor:pointer;
       background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath fill='%231F3864' d='M0 0l5 6 5-6z'/%3E%3C/svg%3E") no-repeat right 8px center;
       -webkit-appearance:none;appearance:none;min-width:130px;transition:border-color .15s,background .15s}
select:focus{outline:none;border-color:var(--blue)}
select.on{border-color:var(--blue);background-color:#EBF3FF}
.clr-btn{padding:7px 14px;border:2px solid var(--border);border-radius:6px;
         cursor:pointer;font-size:13px;font-weight:700;color:var(--muted);
         background:#fff;transition:all .15s}
.clr-btn:hover{border-color:#e74c3c;color:#e74c3c}
.fstats{font-size:13px;color:var(--muted);white-space:nowrap}
.fstats strong{color:var(--navy)}
.wk-btn{padding:7px 28px 7px 10px;border:2px solid var(--border);border-radius:6px;
        font-size:13px;color:var(--text);cursor:pointer;background:#fff;
        white-space:nowrap;font-weight:400;text-align:left;min-width:130px;
        background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath fill='%231F3864' d='M0 0l5 6 5-6z'/%3E%3C/svg%3E");
        background-repeat:no-repeat;background-position:right 8px center;transition:border-color .15s}
.wk-btn:hover,.wk-btn.on{border-color:var(--blue);background-color:#EBF3FF}
.wk-dd{display:none;position:absolute;top:calc(100% + 4px);left:0;z-index:500;
       background:#fff;border:1px solid var(--border);border-radius:8px;
       box-shadow:0 6px 20px rgba(0,0,0,.12);min-width:220px;max-height:320px;overflow-y:auto}
.wk-dd-row{display:flex;align-items:center;gap:8px;padding:7px 14px;cursor:pointer;font-size:13px}
.wk-dd-row:hover{background:#EBF3FF}
.wk-dd-row.all-row{font-weight:700;border-bottom:1px solid var(--border);color:var(--navy)}
.wk-dd-row input{width:14px;height:14px;cursor:pointer;accent-color:var(--blue)}
.cards{display:flex;gap:14px;padding:16px 24px}
.card{background:var(--white);border-radius:10px;padding:14px 18px;flex:1;
      border:1px solid var(--border);box-shadow:0 1px 3px rgba(0,0,0,.05)}
.card-lbl{font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.6px}
.card-val{font-size:24px;font-weight:800;color:var(--navy);margin:4px 0 2px}
.card-sub{font-size:12px;color:var(--muted)}
.twrap{padding:0 24px 32px;overflow:auto;max-height:calc(100vh - 250px);border-radius:10px;border:1px solid var(--border);box-shadow:0 1px 4px rgba(0,0,0,.06)}
table{width:100%;border-collapse:collapse;background:var(--white)}
thead{position:sticky;top:0;z-index:100}
th{background:var(--navy);color:#fff;padding:10px 8px;text-align:center;
   white-space:nowrap;cursor:pointer;user-select:none;font-size:12px;font-weight:600}
th:first-child{text-align:left;padding-left:12px;min-width:220px}
th.scol{min-width:110px}
th.wcol{min-width:76px;font-size:11px;line-height:1.35}
th:hover{background:#2d4f82}
.sa::after{content:' ↑';opacity:.9}
.sd::after{content:' ↓';opacity:.9}
tbody tr{border-bottom:1px solid #F0F4F8;transition:background .1s}
tbody tr:hover{background:#EBF3FF}
tbody tr:last-child{border-bottom:none}
td{padding:8px 8px;text-align:center;vertical-align:middle;font-size:13px}
td:first-child{text-align:left;padding-left:10px}
.vc{display:flex;align-items:center;gap:9px}
.thumb{width:72px;height:40px;border-radius:5px;object-fit:cover;background:#e2e8f0;flex-shrink:0;display:block}
.vinfo{min-width:0}
.vid-id{font-weight:600;font-size:13px;color:var(--navy);
        display:block;text-overflow:ellipsis;overflow:hidden;white-space:nowrap;max-width:240px}
.vid-link{font-size:12px;color:var(--blue);text-decoration:none;display:inline-block;margin-top:2px}
.vid-link:hover{text-decoration:underline}
.spend{font-weight:700;color:var(--navy);font-size:13px}
.rk{display:inline-block;min-width:44px;padding:3px 6px;border-radius:5px;
    font-weight:700;font-size:13px;text-align:center;cursor:default;position:relative}
.rk[data-tip]:hover::after{
  content:attr(data-tip);position:absolute;bottom:calc(100% + 5px);left:50%;
  transform:translateX(-50%);background:#333;color:#fff;padding:3px 8px;
  border-radius:4px;font-size:10px;font-weight:400;white-space:nowrap;
  pointer-events:none;z-index:999}
.r1{background:var(--g1);color:#fff}
.r5{background:var(--g5);color:#fff}
.r15{background:var(--g15);color:#fff}
.r30{background:var(--y30);color:#333}
.r50{background:var(--o50);color:#fff}
.r99{background:var(--r99);color:#fff}
.rno{color:#ccc;font-weight:400;font-size:11px}
.vbadge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700;white-space:nowrap;cursor:default}
#vtip{position:fixed;background:#1a202c;color:#fff;padding:7px 12px;border-radius:6px;font-size:11px;font-weight:400;line-height:1.5;max-width:260px;pointer-events:none;z-index:9999;display:none;box-shadow:0 4px 12px rgba(0,0,0,.3)}
.v-scale{background:#d4edda;color:#155724}
.v-keep{background:#cce5ff;color:#004085}
.v-monitor{background:#fff3cd;color:#856404}
.v-pause{background:#f8d7da;color:#721c24}
.v-caution{background:#ffe5b4;color:#7d3c00;border:1px solid #f0a500}
.v-inactive{background:#e2e3e5;color:#383d41}
.v-new{background:#d1ecf1;color:#0c5460}
.fmt-badge{background:#E6F1FB;color:#0C447C;font-size:10px;padding:1px 6px;border-radius:8px;font-weight:600;margin-left:5px;vertical-align:middle}
.vid-meta{display:flex;align-items:center;gap:5px;margin-top:3px;flex-wrap:wrap}
.vid-cr{font-size:11px;color:var(--muted);font-weight:500}
.vid-prod{font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;display:inline-block}
.nodata{text-align:center;padding:48px;color:var(--muted)}
.nodata h3{font-size:16px;margin-bottom:6px}
.guide{background:var(--white);border:1px solid var(--border);border-radius:10px;margin:0 24px 24px;padding:18px 24px}
.guide-toggle{cursor:pointer;font-size:13px;font-weight:700;color:var(--navy);list-style:none;display:flex;align-items:center;gap:8px}
.guide-toggle::before{content:'📋';font-size:16px}
.guide-body{margin-top:14px;display:none}
details[open] .guide-body{display:block}
.guide-body ol{padding-left:20px;line-height:1.8;color:var(--text)}
.guide-body code{background:#f4f6f9;padding:2px 6px;border-radius:4px;font-size:12px;font-family:monospace}
.guide-body .note{background:#FFF3CD;border:1px solid #FFEAA7;border-radius:6px;padding:10px 14px;margin-top:10px;font-size:12px;color:#856404}
/* ── Tab navigation ── */
.tab-nav{background:var(--white);border-bottom:2px solid var(--border);padding:0 24px;display:flex;gap:0}
.tab-btn{padding:10px 20px;font-size:13px;font-weight:700;cursor:pointer;color:var(--muted);
         background:none;border:none;border-bottom:3px solid transparent;margin-bottom:-2px;transition:color .15s}
/* ── Insight Tab ── */
.ins-section{padding:0 24px 28px}
.ins-h2{font-size:15px;font-weight:700;color:var(--navy);margin:0 0 12px;padding-bottom:8px;border-bottom:2px solid var(--border);display:flex;align-items:center;gap:8px}
.ins-cat-btn{padding:5px 14px;border:2px solid var(--border);border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;color:var(--muted);background:#fff;transition:all .15s}
.ins-cat-btn:hover,.ins-cat-btn.active{border-color:var(--blue);color:var(--blue);background:#EBF3FF}
.mom-bar-wrap{display:flex;align-items:center;gap:8px}
.mom-bar{height:10px;border-radius:5px;background:#1a9940;display:inline-block;min-width:2px}
.mom-num{font-size:12px;font-weight:700;color:var(--navy);min-width:28px}
.brief-card{background:#f8fafd;border:1px solid var(--border);border-left:4px solid;border-radius:8px;padding:12px 16px;margin-bottom:10px}
.brief-card h3{font-size:13px;font-weight:700;color:var(--navy);margin:0 0 4px}
.brief-card p{font-size:12px;color:var(--muted);line-height:1.55;margin:0}
.brief-tag{display:inline-block;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;margin-right:6px;margin-bottom:4px}
.matrix-cell{font-size:11px;text-align:center;padding:4px 6px;border-radius:4px;cursor:default;white-space:nowrap}
.ins-generate-btn{padding:10px 24px;background:var(--navy);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;transition:background .15s}
.ins-generate-btn:hover{background:#2d4f82}
.tab-btn:hover{color:var(--navy)}
.tab-btn.active{color:var(--navy);border-bottom-color:var(--blue)}
/* ── Creators tab — creator row ── */
.crow{border-bottom:1px solid #F0F4F8;cursor:pointer;transition:background .1s}
.crow:hover td,.crow.open td{background:#EBF3FF}
.crow td{padding:8px 8px;text-align:center;vertical-align:middle}
.crow td:first-child{text-align:left;padding-left:10px}
.creator-cell{display:flex;align-items:center;gap:9px}
.c-avatar{width:36px;height:36px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700}
.c-info{min-width:0}
.c-name{font-weight:700;font-size:14px;color:var(--navy)}
.c-tags{display:flex;gap:4px;margin-top:3px;flex-wrap:wrap;align-items:center}
.cbadge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700}
.cb-inf{background:#cce5ff;color:#004085}
.cb-int{background:#d4edda;color:#155724}
.cb-hair{background:#fff3cd;color:#856404}
.cb-beard{background:#d1ecf1;color:#0c5460}
.cb-nutr{background:#e8d5f5;color:#4a235a}
.cb-all{background:#e2e3e5;color:#383d41}
.c-chevron{font-size:11px;color:var(--muted);padding-left:4px!important;width:20px}
/* ── Creators tab — video sub-row ── */
.vrow{border-bottom:1px solid #F0F4F8;background:#F7FAFF}
.vrow td{padding:6px 8px;text-align:center;vertical-align:middle;font-size:13px}
.vrow td:first-child{text-align:left;padding-left:32px}
.vid-cell{display:flex;align-items:center;gap:8px}
.vthumb{width:60px;height:34px;border-radius:4px;object-fit:cover;background:#e2e8f0;flex-shrink:0}
.vid-code{font-size:12px;color:var(--blue);font-family:monospace;display:block}
.vid-narr{font-size:11px;color:var(--muted)}
</style>
</head>
<body>

<div id="vtip"></div>
<div class="hdr">
  <div class="hdr-left">
    <h1>📊 Man Matters — YouTube Creative Performance Dashboard</h1>
    <div class="sub" id="hdr-sub"></div>
  </div>
  <div class="hdr-badge" id="hdr-badge"></div>
</div>

<div class="tab-nav">
  <button class="tab-btn active" data-tab="tab1" onclick="switchTab('tab1')">📊 YT Performance</button>
  <button class="tab-btn" data-tab="tab2" onclick="switchTab('tab2')">👥 Creators</button>
  <button class="tab-btn" data-tab="tab3" onclick="switchTab('tab3')">🧠 Insights</button>
</div>

<div id="tab1">
<div style="background:var(--white);border-bottom:1px solid var(--border);padding:10px 24px 10px;position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.08)">

  <!-- Filter Row 1: Dropdowns -->
  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px">
    <div class="fg"><span class="flabel">Campaign Type</span><select id="sCT" onchange="onCTChange()"></select></div>
    <div class="fg"><span class="flabel">Category</span><select id="sCAT" onchange="onCATChange()"></select></div>
    <div class="fg"><span class="flabel">Campaign</span><select id="sCAMP" onchange="onCampChange()"></select></div>
    <div class="fg"><span class="flabel">Ad Group</span><select id="sAG" onchange="render()"></select></div>
    <div class="fg"><span class="flabel">Status</span>
      <select id="sSTATUS" onchange="render()">
        <option value="All">All</option>
        <option value="Active">🟢 Active</option>
        <option value="Inactive">⚫ Inactive</option>
      </select>
    </div>
    <div class="fg"><span class="flabel">Video Format</span>
      <select id="sFMT" onchange="render()">
        <option value="All">All</option>
        <option value="16:9 Int">16:9 — Int</option>
        <option value="Others">Others</option>
      </select>
    </div>
    <div class="fg" style="position:relative" id="wk-wrap">
      <span class="flabel">Weeks</span>
      <button class="wk-btn" id="wk-btn" onclick="toggleWkDD(event)">All Weeks</button>
      <div class="wk-dd" id="wk-dd">
        <label class="wk-dd-row all-row" onclick="event.stopPropagation()">
          <input type="checkbox" id="wk-all" onchange="toggleAllWks(this)" checked> Select All
        </label>
      </div>
    </div>
  </div>

  <!-- Filter Row 2: Toggles + Search + Stats -->
  <div style="display:flex;align-items:center;gap:12px">
    <label style="display:flex;align-items:center;gap:6px;font-size:13px;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap">
      <input type="checkbox" id="sHIDE500" onchange="render()" style="width:14px;height:14px;cursor:pointer">
      Hide spend below ₹500
    </label>
    <button class="clr-btn" onclick="clearFilters()">✕ Clear</button>
    <div style="margin-left:auto;display:flex;align-items:center;gap:14px">
      <div style="display:flex;align-items:center;gap:6px;border:2px solid var(--border);border-radius:6px;padding:5px 10px;background:#fff;min-width:200px;transition:border-color .15s" id="search-wrap">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="color:var(--muted);flex-shrink:0"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input id="sSearch" type="text" placeholder="Search video ID…"
          oninput="render()" onfocus="this.closest('#search-wrap').style.borderColor='var(--blue)'" onblur="this.closest('#search-wrap').style.borderColor='var(--border)'"
          style="border:none;outline:none;font-size:12px;color:var(--text);background:transparent;width:150px">
        <span id="search-clear" onclick="clearSearch()" style="cursor:pointer;color:var(--muted);font-size:14px;display:none;line-height:1">✕</span>
      </div>
      <div class="fstats">
        Showing&nbsp;<strong id="st-count">–</strong>&nbsp;videos
        &nbsp;|&nbsp;Total Spend:&nbsp;<strong id="st-spend">–</strong>
      </div>
    </div>
  </div>

</div>

<div class="cards">
  <div class="card">
    <div class="card-lbl">Active Videos</div>
    <div class="card-val" id="cv-count">–</div>
    <div class="card-sub">matching current filter</div>
  </div>
  <div class="card">
    <div class="card-lbl">Total Spend (₹)</div>
    <div class="card-val" id="cv-spend">–</div>
    <div class="card-sub" id="cv-range">–</div>
  </div>
  <div class="card">
    <div class="card-lbl">Weeks Tracked</div>
    <div class="card-val" id="cv-weeks">–</div>
    <div class="card-sub" id="cv-wkrange">–</div>
  </div>
  <div class="card">
    <div class="card-lbl">🏆 Top Video This Period</div>
    <div class="card-val" id="cv-top" style="font-size:13px;word-break:break-all">–</div>
    <div class="card-sub" id="cv-topspend">–</div>
  </div>
</div>

<div class="twrap">
  <table>
    <thead><tr id="thead"></tr></thead>
    <tbody id="tbody"></tbody>
  </table>
  <div id="nodata" class="nodata" style="display:none">
    <h3>No videos found</h3>
    <p>No videos had spend under the current filter combination.</p>
  </div>
</div>

<details class="guide">
  <summary class="guide-toggle">How to Refresh This Dashboard Weekly</summary>
  <div class="guide-body">
    <ol>
      <li>Download the weekly <strong>Asset Association Report</strong> from Google Ads (master account level)</li>
      <li>Rename the file: <code>Asset association report (N).xlsx</code> — where N is the next number</li>
      <li>Place it in the uploads folder alongside all existing files</li>
      <li>Open <code>build_dashboard.py</code>, scroll to the <strong>WEEK_GROUPS</strong> section, and add a new line:<br>
          <code>("18 May 2026 - 24 May 2026", ["Asset association report (N).xlsx"])</code></li>
      <li>Run: <code>python build_dashboard.py</code></li>
      <li>Open the new <code>MM_YT_Dashboard.html</code> — done!</li>
    </ol>
    <div class="note">
      💡 From next week onwards you'll download from the master account level which combines all accounts automatically — just one file per week needed.
    </div>
  </div>
</details>
</div><!-- end tab1 -->

<!-- ═══════════════════════════════════════════════════════════════ TAB 2 -->
<div id="tab2" style="display:none">

<div style="background:var(--white);border-bottom:1px solid var(--border);padding:10px 24px;position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.08)">

  <!-- Creator Filter Row 1: Dropdowns + Weeks -->
  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px">
    <div class="fg"><span class="flabel">Category</span><select id="cCAT" onchange="cRender()"></select></div>
    <div class="fg"><span class="flabel">Type</span>
      <select id="cTYPE" onchange="cRender()">
        <option value="All">All</option>
        <option value="Inf">Inf</option>
        <option value="Int">Int</option>
      </select>
    </div>
    <div class="fg"><span class="flabel">Creator Type</span>
      <select id="cCT_TYPE" onchange="cRender()">
        <option value="All">All</option>
        <option value="YT Creator">📺 YT Creator</option>
        <option value="Meta Creator">📱 Meta Creator</option>
      </select>
    </div>
    <div class="fg"><span class="flabel">Status</span>
      <select id="cSTATUS" onchange="cRender()">
        <option value="All">All</option>
        <option value="Active">🟢 Active</option>
        <option value="Inactive">⚫ Inactive</option>
      </select>
    </div>
    <div class="fg" style="position:relative" id="cwk-wrap">
      <span class="flabel">Weeks</span>
      <button class="wk-btn" id="cwk-btn" onclick="toggleCWkDD(event)">All Weeks</button>
      <div class="wk-dd" id="cwk-dd">
        <label class="wk-dd-row all-row" onclick="event.stopPropagation()">
          <input type="checkbox" id="cwk-all" onchange="toggleAllCWks(this)" checked> Select All
        </label>
      </div>
    </div>
  </div>

  <!-- Creator Filter Row 2: Toggles + Search + Stats -->
  <div style="display:flex;align-items:center;gap:12px">
    <label style="display:flex;align-items:center;gap:6px;font-size:13px;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap">
      <input type="checkbox" id="cHIDE500" onchange="cRender()" style="width:14px;height:14px;cursor:pointer">
      Hide spend below ₹500
    </label>
    <button class="clr-btn" onclick="cClearFilters()">✕ Clear</button>
    <div style="margin-left:auto;display:flex;align-items:center;gap:14px">
      <div style="display:flex;align-items:center;gap:6px;border:2px solid var(--border);border-radius:6px;padding:5px 10px;background:#fff;min-width:200px;transition:border-color .15s" id="c-search-wrap">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="color:var(--muted);flex-shrink:0"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input id="cSearch" type="text" placeholder="Search creator name…"
          oninput="cRender()" onfocus="this.closest('#c-search-wrap').style.borderColor='var(--blue)'" onblur="this.closest('#c-search-wrap').style.borderColor='var(--border)'"
          style="border:none;outline:none;font-size:12px;color:var(--text);background:transparent;width:150px">
        <span id="c-search-clear" onclick="clearCSearch()" style="cursor:pointer;color:var(--muted);font-size:14px;display:none;line-height:1">✕</span>
      </div>
      <div class="fstats">
        Showing&nbsp;<strong id="c-count">–</strong>&nbsp;creators
        &nbsp;|&nbsp;Total Spend:&nbsp;<strong id="c-spend">–</strong>
      </div>
    </div>
  </div>

</div>

<div class="cards">
  <div class="card"><div class="card-lbl">Creators</div><div class="card-val" id="cc-count">–</div><div class="card-sub" id="cc-infsplit">–</div></div>
  <div class="card"><div class="card-lbl">Total Videos</div><div class="card-val" id="cc-vids">–</div><div class="card-sub">with spend data</div></div>
  <div class="card"><div class="card-lbl">Total Spend (₹)</div><div class="card-val" id="cc-spend">–</div><div class="card-sub" id="cc-range">–</div></div>
  <div class="card"><div class="card-lbl">🏆 Top Creator</div><div class="card-val" id="cc-top" style="font-size:14px">–</div><div class="card-sub" id="cc-topspend">–</div></div>
</div>

<div class="twrap">
  <table>
    <thead><tr id="c-thead"></tr></thead>
    <tbody id="c-tbody"></tbody>
  </table>
  <div id="c-nodata" class="nodata" style="display:none">
    <h3>No creators found</h3>
    <p>No creators matched the current filter.</p>
  </div>
</div>

</div><!-- end tab2 -->

<!-- ═══════════════════════════════════════════════════════════════ TAB 3 -->
<div id="tab3" style="display:none">

<!-- Category filter bar -->
<div style="background:var(--white);border-bottom:1px solid var(--border);padding:12px 24px;position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.06);display:flex;align-items:center;gap:10px">
  <span class="flabel">Category</span>
  <button class="ins-cat-btn active" id="ins-all"  onclick="insSetCat('All',this)">All</button>
  <button class="ins-cat-btn" id="ins-hair" onclick="insSetCat('Hair',this)">Hair</button>
  <button class="ins-cat-btn" id="ins-beard" onclick="insSetCat('Beard',this)">Beard</button>
  <button class="ins-cat-btn" id="ins-nutr" onclick="insSetCat('Nutrition',this)">Nutrition</button>
  <span style="margin-left:auto;font-size:12px;color:var(--muted)">Based on all active weeks</span>
</div>

<!-- Summary cards -->
<div class="cards">
  <div class="card"><div class="card-lbl">Top Hook by CAC</div><div class="card-val" id="ins-c1" style="font-size:16px">–</div><div class="card-sub" id="ins-c1s">–</div></div>
  <div class="card"><div class="card-lbl">Untapped Hook</div><div class="card-val" id="ins-c2" style="font-size:16px">–</div><div class="card-sub" id="ins-c2s">–</div></div>
  <div class="card"><div class="card-lbl">Highest Momentum</div><div class="card-val" id="ins-c3" style="font-size:14px">–</div><div class="card-sub" id="ins-c3s">–</div></div>
  <div class="card"><div class="card-lbl">Content Gap Alert</div><div class="card-val" id="ins-c4" style="font-size:14px">–</div><div class="card-sub" id="ins-c4s">–</div></div>
</div>

<!-- Section 1: Narrative Intelligence -->
<div class="ins-section">
  <div class="ins-h2">1 — Hook & Narrative Intelligence</div>
  <div style="font-size:11px;color:var(--muted);margin-bottom:8px">Click any row to see which creators ran that narrative and their individual CAC.</div>
  <div class="twrap" style="max-height:420px">
    <table>
      <thead><tr>
        <th style="text-align:left;padding-left:12px;min-width:30px;width:36px"></th>
        <th style="text-align:left;min-width:180px">Narrative / Hook</th>
        <th>Videos</th><th>Total Spend</th><th>Conversions</th><th>Avg CAC</th><th>Scale Rate</th><th style="cursor:default">Signal</th>
      </tr></thead>
      <tbody id="ins-narr-body"></tbody>
    </table>
  </div>
</div>

<!-- Section 4: Momentum Leaders -->
<div class="ins-section">
  <div class="ins-h2">4 — Momentum Leaders <span style="font-size:11px;font-weight:400;color:var(--muted)">(based on spend + conv trend, last 6 weeks)</span></div>
  <div style="background:#EBF3FF;border:1px solid #b8d4f0;border-radius:8px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:#185FA5;line-height:1.6">
    <strong>How is Momentum Score (0–10) calculated?</strong><br>
    Each video gets a score based on two signals over the last 6 weeks:<br>
    <strong>Spend trend</strong> — is the algo spending more or less recently vs before? <strong>Conversion trend</strong> — are conversions improving?<br>
    🟢 8–10 = Scaling fast &nbsp;|&nbsp; 🟡 5–7 = Stable &nbsp;|&nbsp; 🟠 3–4 = Softening &nbsp;|&nbsp; 🔴 0–2 = Declining<br>
    <em>Hover over any row's score bar to see the breakdown.</em>
  </div>
  <div class="twrap" style="max-height:380px">
    <table>
      <thead><tr>
        <th style="text-align:left;padding-left:12px;min-width:200px">Video</th>
        <th style="text-align:left">Category</th>
        <th style="text-align:left">Narrative</th>
        <th style="min-width:160px">Momentum Score</th>
        <th>Wk15 Spend</th><th>Wk16 Spend</th><th style="cursor:default">Verdict</th>
      </tr></thead>
      <tbody id="ins-mom-body"></tbody>
    </table>
  </div>
</div>

<!-- Section 5: Video Brain -->
<div class="ins-section">
  <div class="ins-h2">5 — Video Brain <span style="font-size:11px;font-weight:400;color:var(--muted)">Predictive brief generator based on your performance patterns</span></div>

  <!-- Remarks / Your Context -->
  <div style="background:#f8fafd;border:1px solid var(--border);border-radius:10px;padding:14px 18px;margin-bottom:16px">
    <div style="font-size:13px;font-weight:700;color:var(--navy);margin-bottom:4px">📝 Your Notes & Context</div>
    <div style="font-size:12px;color:var(--muted);margin-bottom:8px">Add your observations, corrections or context here. The brief generator will factor these in when you click Generate. Notes are saved automatically.</div>
    <textarea id="ins-remarks"
      placeholder="e.g. 'Problem Based is working really well in June. Beard needs more Myth Buster. Jovin is our best Hair creator right now. Nutrition Harsh Fit dropped due to campaign issue not creative..'"
      style="width:100%;height:90px;border:2px solid var(--border);border-radius:6px;padding:8px 10px;font-size:12px;color:var(--text);resize:vertical;font-family:inherit;background:#fff;transition:border-color .15s"
      onfocus="this.style.borderColor='var(--blue)'" onblur="this.style.borderColor='var(--border)'"
      oninput="saveInsRemarks()"></textarea>
    <div style="font-size:11px;color:var(--muted);margin-top:4px" id="ins-remarks-status">Notes auto-saved in browser.</div>
  </div>

  <div style="margin-bottom:14px">
    <button class="ins-generate-btn" onclick="insGenerateBriefs()">✨ Generate Creative Briefs</button>
  </div>
  <div id="ins-briefs">
    <div style="color:var(--muted);font-size:13px;padding:8px 0">Click "Generate" to analyse patterns and get creative brief recommendations.</div>
  </div>
</div>

</div><!-- end tab3 -->

<script>
const D = __DATA__;
const C = __CREATORS_DATA__;
const RIGHTS_16_9_DATES = __RIGHTS_16_9_DATA__;

let sortCol = -2, sortDir = 'desc';
let selectedWeeks = null;        // Tab 1 week filter
let selectedCreatorWeeks = null; // Tab 2 week filter

const $ = id => document.getElementById(id);
const inr = n => '₹' + Math.round(n).toLocaleString('en-IN');
const inrS = n => {
  if(n>=1e7) return '₹'+Math.round(n/1e7)+'Cr';
  if(n>=1e5) return '₹'+Math.round(n/1e5)+'L';
  if(n>=1e3) return '₹'+Math.round(n/1e3)+'K';
  return '₹'+Math.round(n);
};
const rkCls = r => {
  if(!r)    return 'rno';
  if(r===1) return 'r1';
  if(r<=5)  return 'r5';
  if(r<=15) return 'r15';
  if(r<=30) return 'r30';
  if(r<=50) return 'r50';
  return 'r99';
};
const getF = () => ({
  ct:      $('sCT').value,
  cat:     $('sCAT').value,
  camp:    $('sCAMP').value,
  ag:      $('sAG').value,
  status:  $('sSTATUS').value,
  fmt:     $('sFMT').value,
  hide500: $('sHIDE500').checked,
  search:  $('sSearch').value.trim().toLowerCase()
});

function weeklySpend(video, f) {
  const w = new Array(D.meta.week_labels.length).fill(0);
  for(const c of video.combos){
    if(f.ct   !=='All' && c.ct  !==f.ct)   continue;
    if(f.cat  !=='All' && c.cat !==f.cat)  continue;
    if(f.camp !=='All' && c.c   !==f.camp) continue;
    if(f.ag   !=='All' && c.ag  !==f.ag)   continue;
    for(let i=0;i<c.w.length;i++) w[i]+=c.w[i];
  }
  return w;
}
function weeklyConv(video, f) {
  const wc = new Array(D.meta.week_labels.length).fill(0);
  for(const c of video.combos){
    if(f.ct   !=='All' && c.ct  !==f.ct)   continue;
    if(f.cat  !=='All' && c.cat !==f.cat)  continue;
    if(f.camp !=='All' && c.c   !==f.camp) continue;
    if(f.ag   !=='All' && c.ag  !==f.ag)   continue;
    if(c.wc) for(let i=0;i<c.wc.length;i++) wc[i]+=c.wc[i];
  }
  return wc;
}

// Returns top campaign dependency for a video under current filter.
// dep = { topCamp, depPct, totalConv, numCamps }
function getCampDep(video, f) {
  const campConv = {};
  for(const c of video.combos){
    if(f.ct   !=='All' && c.ct  !==f.ct)   continue;
    if(f.cat  !=='All' && c.cat !==f.cat)  continue;
    if(f.camp !=='All' && c.c   !==f.camp) continue;
    if(f.ag   !=='All' && c.ag  !==f.ag)   continue;
    const cv = (c.wc||[]).reduce((a,b)=>a+b, 0);
    if(cv > 0) campConv[c.c] = (campConv[c.c]||0) + cv;
  }
  const entries = Object.entries(campConv).sort((a,b)=>b[1]-a[1]);
  const totalConv = entries.reduce((s,[,v])=>s+v, 0);
  if(totalConv === 0 || entries.length === 0) return null;
  const [topCamp, topConv] = entries[0];
  return {
    topCamp,
    depPct:    Math.round(topConv / totalConv * 100),
    totalConv: Math.round(totalConv * 10) / 10,
    numCamps:  entries.length
  };
}

function buildHeader(){
  const visIdxs = selectedWeeks !== null ? selectedWeeks : D.meta.week_labels.map((_,i)=>i);
  const isPeriod = selectedWeeks !== null && selectedWeeks.length < D.meta.week_labels.length;
  const spendLbl = isPeriod ? 'Period Spend ₹' : 'Total Spend ₹';
  let h = `<th onclick="setSort(-1)">Video<span></span></th>
            <th class="scol" onclick="setSort(-2)">${spendLbl}<span></span></th>`;
  visIdxs.forEach(wi=>
    h+=`<th class="wcol" onclick="setSort(${wi})">${D.meta.week_labels[wi]}<br>Rank<span></span></th>`
  );
  h += `<th class="scol" style="min-width:160px;cursor:default">Verdict<span></span></th>`;
  $('thead').innerHTML = h;
}

function setSort(col){
  if(sortCol===col){ sortDir = sortDir==='asc'?'desc':'asc'; }
  else{ sortCol=col; sortDir = col>=0?'asc':'desc'; }
  render();
}

function render(){
  const f = getF();
  const wks = D.meta.week_labels;
  const visIdxs = selectedWeeks !== null ? selectedWeeks : wks.map((_,i)=>i);
  const isPeriod = selectedWeeks !== null && selectedWeeks.length < wks.length;

  buildHeader(); // rebuild to update period label + visible columns

  let rows = D.videos.map(v=>{
    const w  = weeklySpend(v,f);
    const wc = weeklyConv(v,f);
    // Period total = sum of selected (or all) weeks
    const total = visIdxs.reduce((s,i)=>s+(w[i]||0), 0);
    return {...v,w,wc,total:Math.round(total*100)/100};
  }).filter(v=>{
    if(v.total===0) return false;
    if(f.fmt     !== 'All' && v.fmt !== f.fmt) return false;
    if(f.hide500 && v.total < 500)             return false;
    if(f.status==='Active')   return v.w[v.w.length-1] > 0;
    if(f.status==='Inactive') return v.w[v.w.length-1] === 0;
    if(f.search){
      const vid = v.id.toLowerCase();
      const allIds = D.videos.map(x=>x.id.toLowerCase());
      if(allIds.includes(f.search)) return vid === f.search;
      return vid.includes(f.search);
    }
    return true;
  });

  // Ranks per visible week position (each rank cell = that week's spend rank)
  const ranksByPos = visIdxs.map(wi=>{
    const spends = rows.map(v=>v.w[wi]);
    const sorted = [...spends].filter(s=>s>0).sort((a,b)=>b-a);
    return spends.map(s=> s===0 ? null : sorted.indexOf(s)+1);
  });
  rows = rows.map((v,i)=>({...v, ranks: ranksByPos.map(rk=>rk[i])}));

  rows.sort((a,b)=>{
    if(sortCol===-1) return sortDir==='asc'?a.url.localeCompare(b.url):b.url.localeCompare(a.url);
    if(sortCol===-2) return sortDir==='asc'?a.total-b.total:b.total-a.total;
    const pos = visIdxs.indexOf(sortCol);
    if(pos===-1) return 0;
    const ra=a.ranks[pos]??99999, rb=b.ranks[pos]??99999;
    return sortDir==='asc'?ra-rb:rb-ra;
  });

  const totalSpend = rows.reduce((s,v)=>s+v.total,0);
  $('cv-count').textContent = rows.length;
  $('cv-spend').textContent = inrS(totalSpend);
  $('cv-range').textContent = isPeriod ? `${visIdxs.length} of ${wks.length} weeks` : D.meta.week_range;
  $('cv-weeks').textContent = visIdxs.length;
  $('cv-wkrange').textContent = isPeriod ? `${visIdxs.length} weeks selected` : D.meta.week_range;
  if(rows.length>0){
    $('cv-top').textContent = rows[0].id;
    $('cv-topspend').textContent = inr(rows[0].total)+(isPeriod?' period':' total');
  }
  $('st-count').textContent = rows.length;
  $('st-spend').textContent = inr(totalSpend);

  document.querySelectorAll('thead th').forEach((th,i)=>{
    th.classList.remove('sa','sd');
    const sp=th.querySelector('span'); if(sp) sp.textContent='';
    const lastCol = 2 + visIdxs.length;
    let col;
    if(i===0) col=-1;
    else if(i===1) col=-2;
    else if(i===lastCol) col=null;
    else col=visIdxs[i-2]; // map display position → actual week index
    if(col!==null && col===sortCol){ th.classList.add(sortDir==='asc'?'sa':'sd'); if(sp) sp.textContent=sortDir==='asc'?' ↑':' ↓'; }
  });
  ['sCT','sCAT','sCAMP','sAG','sSTATUS','sFMT'].forEach(id=>{
    const el=$(id); if(el) el.classList.toggle('on',el.value!=='All');
  });
  ['cCAT','cTYPE','cCT_TYPE','cSTATUS'].forEach(id=>{
    const el=$(id); if(el) el.classList.toggle('on',el.value!=='All');
  });
  const csv = $('cSearch').value.trim();
  const csc = $('c-search-clear');
  if(csc) csc.style.display = csv ? 'inline' : 'none';
  const csw = $('c-search-wrap');
  if(csw) csw.style.borderColor = csv ? 'var(--blue)' : 'var(--border)';
  const sv = $('sSearch').value.trim();
  $('search-clear').style.display = sv ? 'inline' : 'none';
  $('search-wrap').style.borderColor = sv ? 'var(--blue)' : 'var(--border)';

  const tbody=$('tbody'), nodata=$('nodata');
  if(rows.length===0){ tbody.innerHTML=''; nodata.style.display='block'; return; }
  nodata.style.display='none';

  tbody.innerHTML = rows.map(v=>{
    const thumb=`https://img.youtube.com/vi/${v.id}/mqdefault.jpg`;
    const rkCells = v.ranks.map((r,pos)=>{
      const wi = visIdxs[pos]; // actual week index
      if(!r) return `<td><span class="rk rno">–</span></td>`;
      const tip=`Spend: ${inr(v.w[wi])} | ${wks[wi]}`;
      return `<td><span class="rk ${rkCls(r)}" data-tip="${tip}">#${r}</span></td>`;
    }).join('');
    const dep = getCampDep(v, f);
    const vd = getVerdict(v.w, v.wc, dep); // verdict always uses full weekly data
    const isExact = f.search && v.id.toLowerCase()===f.search;
    const fmtBadge = v.fmt==='16:9 Int' ? `<span class="fmt-badge">16:9 Int</span>` : '';
    let crTag = '', prodTag = '';
    try {
      const crInfo = VID_CREATOR_MAP[v.id];
      if (crInfo) crTag = `<span class="vid-cr">${crInfo.name}</span>`;
      const prod = getVideoProduct(v);
      if (prod && prod !== 'Other') prodTag = `<span class="vid-prod" style="${getProdBadgeStyle(prod)}">${prod}</span>`;
    } catch(e) {}
    return `<tr style="${isExact?'background:#FFFDE7;':''}">
      <td><div class="vc">
        <img class="thumb" src="${thumb}" loading="lazy" onerror="this.style.visibility='hidden'" alt="">
        <div class="vinfo">
          <span class="vid-id" title="${v.url}">${v.id}${fmtBadge}</span>
          <div class="vid-meta">${crTag}${prodTag}</div>
          <a href="${v.url}" target="_blank" class="vid-link">▶ Open on YouTube</a>
        </div>
      </div></td>
      <td class="spend">${inr(v.total)}</td>
      ${rkCells}
      <td><span class="vbadge ${vd.cls}" data-tip="${vd.tip}">${vd.label}</span></td>
    </tr>`;
  }).join('');
}

// ── Verdict logic ─────────────────────────────────────────────────────────────
// Uses spend (w), conversions (wc), and campaign dependency (dep).
// dep = { topCamp, depPct, totalConv, numCamps } — null if no conversion data.
// HIGH DEPENDENCY (>=65% of conv from one campaign): Pause → downgraded to Caution
// so you don't blindly pause a video that's propping up a key campaign's numbers.
function getVerdict(w, wc, dep) {
  const n = w.length;
  const fmt  = v => '₹' + Math.round(v).toLocaleString('en-IN');
  const spct = (a,b) => Math.round((a/b - 1)*100);
  const dpct = (a,b) => Math.round((1 - a/b)*100);

  // Dependency note appended to every tooltip
  const depNote = dep
    ? (dep.numCamps === 1
        ? ` | 100% of conv from ${dep.topCamp} — single campaign dependency.`
        : dep.depPct >= 65
          ? ` | ⚠️ ${dep.depPct}% of conv from ${dep.topCamp} — heavy dependency.`
          : dep.depPct >= 40
            ? ` | ${dep.depPct}% of conv from ${dep.topCamp}.`
            : ` | Conv spread across ${dep.numCamps} campaigns.`)
    : '';

  const activeIdx = w.reduce((acc,v,i)=>{ if(v>0) acc.push(i); return acc; }, []);
  if(activeIdx.length === 0) return {label:'⚫ Inactive', cls:'v-inactive', tip:'No spend recorded' + depNote};

  const firstActive = activeIdx[0];
  const lastActive  = activeIdx[activeIdx.length - 1];

  // Last 3 weeks
  const r0=w[n-3]||0,  r1=w[n-2]||0,  r2=w[n-1]||0;
  const c0=(wc&&wc[n-3])||0, c1=(wc&&wc[n-2])||0, c2=(wc&&wc[n-1])||0;

  // Helper: apply dependency override — if video has heavy campaign dependency,
  // downgrade a Pause to Caution so you don't accidentally kill a key campaign.
  const highDep = dep && dep.depPct >= 65;
  function pauseOrCaution(tip) {
    if (highDep) {
      return {label:'🟠 Caution', cls:'v-caution',
              tip: tip + ` BUT ${dep.depPct}% of conversions depend on ${dep.topCamp} — pausing may cause heavy conv drop. Review with team first.`};
    }
    return {label:'🔴 Pause', cls:'v-pause', tip: tip + depNote};
  }

  // INACTIVE: no spend last 2 weeks
  if (r1===0 && r2===0) {
    const gap = n - 1 - lastActive;
    return {label:'⚫ Inactive', cls:'v-inactive',
            tip:`No spend for ${gap} week${gap>1?'s':''}. May have been paused or eliminated by algo.${depNote}`};
  }

  // NEW ENTRY: only 1-2 weeks of spend
  if (firstActive >= n - 2) {
    const convNote = c2>0 ? ` ${c2} conv last week.` : ' 0 conv so far.';
    return {label:'🆕 New Entry', cls:'v-new',
            tip:`Only ${activeIdx.length} week(s) of data.${convNote} Monitor 2 more weeks before deciding.${depNote}`};
  }

  // PAUSE/CAUTION: low spend (<₹1K) for 2 consecutive weeks with zero conversions
  // Not a new entry — been live long enough but not gaining traction
  if (r1>0 && r1<1000 && r2>0 && r2<1000 && c1===0 && c2===0) {
    return pauseOrCaution(`Spend under ₹1,000/week for 2 weeks in a row with 0 conversions. Not gaining traction — pause.`);
  }

  // PAUSE/CAUTION: high spend, zero conversions last week
  if (r2>0 && c2===0 && r2>1000) {
    return pauseOrCaution(`${fmt(r2)} spent last week with 0 conversions. Burning budget — review creative.`);
  }
  // PAUSE/CAUTION: conversions collapsed while spend held
  if (r1>0 && r2>0 && c1>2 && c2<c1*0.35 && r2>=r1*0.6) {
    return pauseOrCaution(`Conversions dropped ${c1} to ${c2} while spend held at ${fmt(r2)}. Efficiency tanking.`);
  }
  // PAUSE/CAUTION: both spend AND conversions declining 3 weeks straight
  if (r0>0&&r1>0&&r2>0 && r2<r1&&r1<r0 && c2<=c1&&c1<=c0 && !(c0===0&&c1===0&&c2===0)) {
    const sd=dpct(r2,r1);
    return pauseOrCaution(`Spend and conversions both declining 3 weeks running. Spend down ${sd}% last week.`);
  }
  // PAUSE/CAUTION: spend dropped >40% and conversions also dropping
  if (r1>0&&r2>0 && r2<r1*0.60 && c2<=c1) {
    const sd=dpct(r2,r1);
    return pauseOrCaution(`Spend fell ${sd}% (${fmt(r1)} to ${fmt(r2)}) with conv at ${c2} vs ${c1} prior. Algo pulling back.`);
  }

  // SCALE UP: spend growing 3 weeks + conversions holding or improving
  if (r0>0&&r1>0&&r2>0 && r2>r1&&r1>r0 && c2>=c1) {
    const sp=spct(r2,r1);
    return {label:'🟢 Scale Up', cls:'v-scale',
            tip:`Spend +${sp}% week-on-week, 3rd consecutive week. ${c2} conv last week. Increase budget.${depNote}`};
  }
  // SCALE UP: strong spend jump with conversions
  if (r1>0&&r2>r1*1.30 && c2>0) {
    const sp=spct(r2,r1);
    return {label:'🟢 Scale Up', cls:'v-scale',
            tip:`Spend jumped +${sp}% to ${fmt(r2)} with ${c2} conversions. Algo finding scale — increase budget.${depNote}`};
  }
  // SCALE UP: conversions growing strongly
  if (r1>0&&r2>0 && c1>0&&c2>c1*1.30) {
    const cp=spct(c2,c1);
    return {label:'🟢 Scale Up', cls:'v-scale',
            tip:`Conversions up +${cp}% (${c1} to ${c2}) at ${fmt(r2)} spend. Efficiency improving — scale up.${depNote}`};
  }

  // MONITOR: dropped to 0 last week
  if (r2===0 && r1>0) {
    return {label:'⚠️ Monitor', cls:'v-monitor',
            tip:`Spend went to 0 last week (was ${fmt(r1)}). Check if paused intentionally or algo dropped it.${depNote}`};
  }
  // MONITOR: spend OK but zero conversions
  if (r2>0 && c2===0) {
    return {label:'⚠️ Monitor', cls:'v-monitor',
            tip:`${fmt(r2)} spent last week, 0 conversions. May need more time or creative is underperforming.${depNote}`};
  }
  // MONITOR: spend declining but still converting
  if (r1>0&&r2>0 && r2<r1*0.70 && c2>0) {
    const sd=dpct(r2,r1);
    return {label:'⚠️ Monitor', cls:'v-monitor',
            tip:`Spend down ${sd}% but still got ${c2} conversions. Could be algo optimising — watch 1 more week.${depNote}`};
  }

  // KEEP: stable spend + conversions present
  if (r1>0&&r2>0 && r2>=r1*0.70 && c2>0) {
    const chg=spct(r2,r1), sign=chg>=0?'+':'';
    return {label:'✅ Keep', cls:'v-keep',
            tip:`Spend ${sign}${chg}% (${fmt(r2)}) with ${c2} conversions last week. Performing well.${depNote}`};
  }

  return {label:'⚠️ Monitor', cls:'v-monitor',
          tip:'Mixed signals — watch spend and conversion trend this week.' + depNote};
}

// ── Dropdown helpers ──────────────────────────────────────────────────────────
function uniqueSorted(arr){ return ['All',...new Set(arr)].sort((a,b)=>a==='All'?-1:b==='All'?1:a.localeCompare(b)); }
function populate(sel,items){ sel.innerHTML=items.map(v=>`<option value="${v}">${v}</option>`).join(''); sel.value='All'; }

// Cascade: CT → Category → Campaign → AdGroup
function onCTChange(){
  const ct = $('sCT').value;
  const cats = uniqueSorted(
    D.videos.flatMap(v=>v.combos)
      .filter(c=> ct==='All' || c.ct===ct)
      .map(c=>c.cat)
  );
  populate($('sCAT'), cats);
  onCATChange();
}
function onCATChange(){
  const ct=$('sCT').value, cat=$('sCAT').value;
  const camps = uniqueSorted(
    D.videos.flatMap(v=>v.combos)
      .filter(c=>(ct==='All'||c.ct===ct) && (cat==='All'||c.cat===cat))
      .map(c=>c.c)
  );
  populate($('sCAMP'), camps);
  onCampChange();
}
function onCampChange(){
  const ct=$('sCT').value, cat=$('sCAT').value, camp=$('sCAMP').value;
  const ags = uniqueSorted(
    D.videos.flatMap(v=>v.combos)
      .filter(c=>(ct==='All'||c.ct===ct)&&(cat==='All'||c.cat===cat)&&(camp==='All'||c.c===camp))
      .map(c=>c.ag)
  );
  populate($('sAG'), ags);
  render();
}
function clearSearch(){
  $('sSearch').value='';
  render();
}
function clearFilters(){
  $('sCT').value='All';
  $('sSTATUS').value='All';
  $('sFMT').value='All';
  $('sHIDE500').checked=false;
  $('sSearch').value='';
  resetWkFilter();
  onCTChange();
}

// ── Week filter dropdown ──────────────────────────────────────────────────────
function buildWkFilter(){
  const dd = $('wk-dd');
  const allRow = dd.querySelector('.all-row');
  // Clear existing week rows
  [...dd.querySelectorAll('.wk-item')].forEach(el=>el.remove());
  D.meta.week_labels.forEach((lbl,i)=>{
    const row = document.createElement('label');
    row.className='wk-dd-row wk-item';
    row.onclick = e=>e.stopPropagation();
    row.innerHTML=`<input type="checkbox" class="wk-cb" data-idx="${i}" onchange="onWkChange()" checked> ${lbl}`;
    dd.appendChild(row);
  });
}
function toggleWkDD(e){
  e.stopPropagation();
  const dd=$('wk-dd');
  dd.style.display = dd.style.display==='none'||!dd.style.display ? 'block' : 'none';
}
function toggleAllWks(cb){
  document.querySelectorAll('.wk-cb').forEach(c=>c.checked=cb.checked);
  onWkChange();
}
function onWkChange(){
  const all=[...document.querySelectorAll('.wk-cb')];
  const checked=all.filter(c=>c.checked).map(c=>+c.dataset.idx);
  selectedWeeks = checked.length===all.length ? null : (checked.length===0?[0]:checked);
  const allBox=$('wk-all');
  allBox.checked = selectedWeeks===null;
  allBox.indeterminate = selectedWeeks!==null && checked.length>0 && checked.length<all.length;
  updateWkBtn(checked.length, all.length);
  render();
}
function updateWkBtn(n, total){
  const btn=$('wk-btn');
  btn.textContent = n===total ? 'All Weeks' : `${n} Week${n!==1?'s':''}`;
  btn.classList.toggle('on', n!==total);
}
function resetWkFilter(){
  document.querySelectorAll('.wk-cb').forEach(c=>c.checked=true);
  $('wk-all').checked=true; $('wk-all').indeterminate=false;
  selectedWeeks=null;
  updateWkBtn(D.meta.week_labels.length, D.meta.week_labels.length);
}
document.addEventListener('click', e=>{
  if(!e.target.closest('#wk-wrap')){ const dd=$('wk-dd'); if(dd) dd.style.display='none'; }
});

// ── Verdict tooltip (fixed position — never clipped by scroll container) ──────
(function(){
  const tip = document.getElementById('vtip');
  document.addEventListener('mouseover', e => {
    const el = e.target.closest('.vbadge[data-tip]');
    if(!el){ tip.style.display='none'; return; }
    tip.textContent = el.dataset.tip;
    tip.style.display = 'block';
  });
  document.addEventListener('mousemove', e => {
    if(tip.style.display==='none') return;
    const pad = 12, tw = tip.offsetWidth, th = tip.offsetHeight;
    let x = e.clientX - tw/2;
    let y = e.clientY - th - pad;
    // keep within viewport
    if(x < pad) x = pad;
    if(x + tw > window.innerWidth - pad) x = window.innerWidth - tw - pad;
    if(y < pad) y = e.clientY + pad;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
  });
  document.addEventListener('mouseout', e => {
    if(!e.target.closest('.vbadge[data-tip]')) tip.style.display='none';
  });
})();

// ── Init ──────────────────────────────────────────────────────────────────────
(function init(){
  $('hdr-sub').textContent = `${D.meta.week_range}  ·  ${D.meta.total_videos} unique videos tracked`;
  $('hdr-badge').textContent = `Generated: ${D.meta.generated}`;
  populate($('sCT'), D.opts.cts);
  buildHeader();
  onCTChange();
})();

// ═══════════════════════════════════════════════════════════════════════════
// TAB SWITCHING
// ═══════════════════════════════════════════════════════════════════════════
function switchTab(t){
  ['tab1','tab2','tab3'].forEach(id=>{ const el=$(id); if(el) el.style.display = t===id?'':'none'; });
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.tab===t));
  if(t==='tab2') cInit();
  if(t==='tab3') insightInit();
}

// ═══════════════════════════════════════════════════════════════════════════
// CREATORS TAB — JavaScript
// ═══════════════════════════════════════════════════════════════════════════
let cInitDone = false;
const catColors = {
  'Hair':       'cb-hair',
  'Beard':      'cb-beard',
  'Nutrition':  'cb-nutr',
  'All Products':'cb-all',
};
function catBadge(cat){ return `<span class="cbadge ${catColors[cat]||'cb-all'}">${cat}</span>`; }

function cGetF(){
  return {
    cat:     $('cCAT').value,
    type:    $('cTYPE').value,
    ct_type: $('cCT_TYPE').value,
    status:  $('cSTATUS').value,
    hide500: $('cHIDE500').checked,
    search:  $('cSearch').value.trim().toLowerCase(),
  };
}
function clearCSearch(){ $('cSearch').value=''; cRender(); }

function cBuildHeader(){
  const wks = C.meta.week_labels;
  const cVisIdxs = selectedCreatorWeeks !== null ? selectedCreatorWeeks : wks.map((_,i)=>i);
  const cIsPeriod = selectedCreatorWeeks !== null && selectedCreatorWeeks.length < wks.length;
  const spendLbl = cIsPeriod ? 'Period Spend ₹' : 'Total Spend ₹';
  let h = `<th onclick="cSetSort(-1)">Creator<span></span></th>
            <th class="scol" onclick="cSetSort(-2)">${spendLbl}<span></span></th>`;
  cVisIdxs.forEach(wi=> h+=`<th class="wcol" onclick="cSetSort(${wi})">${wks[wi]}<br>Rank<span></span></th>`);
  h += `<th class="wcol" style="cursor:default">Conv<span></span></th>`;
  h += `<th class="wcol" style="cursor:default">CAC ₹<span></span></th>`;
  h += `<th class="scol" style="min-width:140px;cursor:default">Verdict<span></span></th>`;
  $('c-thead').innerHTML = h;
}

let cSortCol = -2, cSortDir = 'desc';
function cSetSort(col){
  if(cSortCol===col){ cSortDir = cSortDir==='asc'?'desc':'asc'; }
  else{ cSortCol=col; cSortDir = col>=0?'asc':'desc'; }
  cRender();
}

function cRender(){
  const f  = cGetF();
  const wks = C.meta.week_labels;
  const cVisIdxs = selectedCreatorWeeks !== null ? selectedCreatorWeeks : wks.map((_,i)=>i);
  const cIsPeriod = selectedCreatorWeeks !== null && selectedCreatorWeeks.length < wks.length;

  cBuildHeader(); // rebuild to update period label + visible columns

  // Recalculate period total per creator based on selected weeks
  const crWithTotal = C.creators.map(cr=>{
    const periodTotal = cVisIdxs.reduce((s,i)=>s+(cr.w[i]||0), 0);
    const periodConv  = cVisIdxs.reduce((s,i)=>s+(cr.wc[i]||0), 0);
    return {...cr, total: Math.round(periodTotal*100)/100, periodConv};
  });

  let rows = crWithTotal.filter(cr=>{
    if(f.cat     !=='All' && cr.category     !== f.cat)     return false;
    if(f.type    !=='All' && cr.inf_int      !== f.type)    return false;
    if(f.ct_type !=='All' && cr.creator_type !== f.ct_type) return false;
    if(f.hide500 && cr.total < 500)                         return false;
    if(f.search && !cr.name.toLowerCase().includes(f.search)) return false;
    if(f.status==='Active')   return cr.w[cr.w.length-1] > 0;
    if(f.status==='Inactive') return cr.w[cr.w.length-1] === 0;
    return true;
  });

  // Ranks per visible week
  const ranksByPos = cVisIdxs.map(wi=>{
    const spends = rows.map(cr=>cr.w[wi]);
    const sorted = [...spends].filter(s=>s>0).sort((a,b)=>b-a);
    return spends.map(s=> s===0 ? null : sorted.indexOf(s)+1);
  });
  rows = rows.map((cr,i)=>({...cr, ranks: ranksByPos.map(rk=>rk[i])}));

  // Sort
  rows.sort((a,b)=>{
    if(cSortCol===-1) return cSortDir==='asc'?a.name.localeCompare(b.name):b.name.localeCompare(a.name);
    if(cSortCol===-2) return cSortDir==='asc'?a.total-b.total:b.total-a.total;
    const pos = cVisIdxs.indexOf(cSortCol);
    if(pos===-1) return 0;
    const ra=a.ranks[pos]??99999, rb=b.ranks[pos]??99999;
    return cSortDir==='asc'?ra-rb:rb-ra;
  });

  // Summary cards
  const totalSpend = rows.reduce((s,cr)=>s+cr.total,0);
  const totalVids  = rows.reduce((s,cr)=>s+cr.videos.length,0);
  const infCount   = rows.filter(cr=>cr.inf_int==='Inf').length;
  const intCount   = rows.filter(cr=>cr.inf_int==='Int').length;
  $('cc-count').textContent  = rows.length;
  $('cc-infsplit').textContent = `${infCount} Inf · ${intCount} Int`;
  $('cc-vids').textContent   = totalVids;
  $('cc-spend').textContent  = inrS(totalSpend);
  $('cc-range').textContent  = cIsPeriod ? `${cVisIdxs.length} of ${wks.length} weeks` : C.meta.week_range;
  $('c-count').textContent   = rows.length;
  $('c-spend').textContent   = inr(totalSpend);
  if(rows.length>0){
    $('cc-top').textContent      = rows[0].name;
    $('cc-topspend').textContent = inr(rows[0].total)+(cIsPeriod?' period':' total');
  }

  // Sort indicators
  document.querySelectorAll('#c-thead th').forEach((th,i)=>{
    th.classList.remove('sa','sd');
    const sp=th.querySelector('span'); if(sp) sp.textContent='';
    const lastDataCol = 2 + cVisIdxs.length;
    let col;
    if(i===0) col=-1;
    else if(i===1) col=-2;
    else if(i>=lastDataCol) col=null;
    else col=cVisIdxs[i-2];
    if(col!==null && col===cSortCol){ th.classList.add(cSortDir==='asc'?'sa':'sd'); if(sp) sp.textContent=cSortDir==='asc'?' ↑':' ↓'; }
  });

  const tbody=$('c-tbody'), nodata=$('c-nodata');
  if(rows.length===0){ tbody.innerHTML=''; nodata.style.display='block'; return; }
  nodata.style.display='none';

  tbody.innerHTML = rows.map((cr,ri)=>{
    const initials = cr.name.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
    const avatarColors = ['#cce5ff:#004085','#d4edda:#155724','#fff3cd:#856404','#d1ecf1:#0c5460','#e8d5f5:#4a235a','#fde8d8:#7d3c00'];
    const [bg,fg] = avatarColors[ri % avatarColors.length].split(':');
    const hasImg = cr.profile_url && cr.profile_url !== 'nan' && cr.profile_url.startsWith('http');
    const rkCells = cr.ranks.map((r,pos)=>{
      const wi = cVisIdxs[pos];
      if(!r) return `<td><span class="rk rno">–</span></td>`;
      return `<td><span class="rk ${rkCls(r)}" data-tip="Spend: ${inr(cr.w[wi])} | ${wks[wi]}">#${r}</span></td>`;
    }).join('');
    const totalConv = Math.round(cr.periodConv);
    const cac = totalConv > 0 ? Math.round(cr.total / totalConv) : 0;
    const vd = getVerdict(cr.w, cr.wc, null); // verdict uses full data
    const rowId = `cvrow-${ri}`;
    const avatarHTML = hasImg
      ? `<img src="${cr.profile_url}" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0" onerror="this.style.display='none';this.nextSibling.style.display='flex'" alt="${cr.name}"><div class="c-avatar" style="background:${bg};color:${fg};display:none">${initials}</div>`
      : `<div class="c-avatar" style="background:${bg};color:${fg}">${initials}</div>`;
    const profileLink = hasImg ? ` <a href="${cr.profile_url}" target="_blank" style="font-size:9px;color:var(--blue);text-decoration:none" title="Open profile">↗</a>` : '';

    // Video sub-rows — also use selected weeks for period spend
    const visibleVids = f.hide500 ? cr.videos.filter(v => {
      const vPeriod = cVisIdxs.reduce((s,i)=>s+(v.w[i]||0),0);
      return vPeriod >= 500;
    }) : cr.videos;
    const vidRows = visibleVids.map(v=>{
      const vPeriodTotal = cVisIdxs.reduce((s,i)=>s+(v.w[i]||0), 0);
      const vPeriodConv  = cVisIdxs.reduce((s,i)=>s+(v.wc[i]||0), 0);
      const vCac  = vPeriodConv > 0 ? Math.round(vPeriodTotal / vPeriodConv) : 0;
      const vVd   = getVerdict(v.w, v.wc, null); // verdict uses full data
      const vRkCells = cVisIdxs.map(wi=>{
        const ws = v.w[wi]||0;
        if(!ws) return `<td><span class="rk rno">–</span></td>`;
        return `<td style="font-size:13px;color:var(--navy);font-weight:600">${inrS(ws)}</td>`;
      }).join('');
      return `<tr class="vrow" id="${rowId}" style="display:none">
        <td><div class="vid-cell">
          <img class="vthumb" src="https://img.youtube.com/vi/${v.id}/mqdefault.jpg" loading="lazy" onerror="this.style.visibility='hidden'" alt="">
          <div>
            <a href="${v.url}" target="_blank" class="vid-code">${v.id}</a>
            <span class="vid-narr">${v.narrative||'–'}</span>
          </div>
        </div></td>
        <td class="spend">${inr(vPeriodTotal)}</td>
        ${vRkCells}
        <td>${Math.round(vPeriodConv)}</td>
        <td>${vCac > 0 ? inr(vCac) : '–'}</td>
        <td><span class="vbadge ${vVd.cls}" data-tip="${vVd.tip}">${vVd.label}</span></td>
      </tr>`;
    }).join('');

    return `<tr class="crow" onclick="cToggle(this,'${rowId}')">
      <td><div class="creator-cell">
        <span class="c-chevron" id="chev-${rowId}">▶</span>
        ${avatarHTML}
        <div class="c-info">
          <div class="c-name">${cr.name}${profileLink}</div>
          <div class="c-tags">
            ${catBadge(cr.category)}
            <span class="cbadge" style="background:${cr.creator_type==='YT Creator'?'#f8d7da':'#e2e3e5'};color:${cr.creator_type==='YT Creator'?'#721c24':'#383d41'}">${cr.creator_type}</span>
            <span style="font-size:9px;color:var(--muted)">${visibleVids.length} video${visibleVids.length>1?'s':''}</span>
          </div>
        </div>
      </div></td>
      <td class="spend">${inr(cr.total)}</td>
      ${rkCells}
      <td>${totalConv}</td>
      <td>${cac > 0 ? inr(cac) : '–'}</td>
      <td><span class="vbadge ${vd.cls}" data-tip="${vd.tip}">${vd.label}</span></td>
    </tr>${vidRows}`;
  }).join('');
}

function cToggle(row, rowId){
  row.classList.toggle('open');
  const isOpen = row.classList.contains('open');
  const chev = document.getElementById(`chev-${rowId}`);
  if(chev) chev.textContent = isOpen ? '▼' : '▶';
  document.querySelectorAll(`#${rowId}`).forEach(r=> r.style.display = isOpen ? '' : 'none');
}

function cClearFilters(){
  $('cCAT').value='All'; $('cTYPE').value='All'; $('cCT_TYPE').value='All'; $('cSTATUS').value='All';
  $('cHIDE500').checked = false;
  $('cSearch').value = '';
  resetCWkFilter();
  cRender();
}

// ── Creator week filter dropdown ──────────────────────────────────────────────
function cBuildWkFilter(){
  const dd = $('cwk-dd');
  [...dd.querySelectorAll('.cwk-item')].forEach(el=>el.remove());
  C.meta.week_labels.forEach((lbl,i)=>{
    const row = document.createElement('label');
    row.className='wk-dd-row cwk-item';
    row.onclick = e=>e.stopPropagation();
    row.innerHTML=`<input type="checkbox" class="cwk-cb" data-idx="${i}" onchange="onCWkChange()" checked> ${lbl}`;
    dd.appendChild(row);
  });
}
function toggleCWkDD(e){
  e.stopPropagation();
  const dd=$('cwk-dd');
  dd.style.display = dd.style.display==='none'||!dd.style.display ? 'block' : 'none';
}
function toggleAllCWks(cb){
  document.querySelectorAll('.cwk-cb').forEach(c=>c.checked=cb.checked);
  onCWkChange();
}
function onCWkChange(){
  const all=[...document.querySelectorAll('.cwk-cb')];
  const checked=all.filter(c=>c.checked).map(c=>+c.dataset.idx);
  selectedCreatorWeeks = checked.length===all.length ? null : (checked.length===0?[0]:checked);
  const allBox=$('cwk-all');
  allBox.checked = selectedCreatorWeeks===null;
  allBox.indeterminate = selectedCreatorWeeks!==null && checked.length>0 && checked.length<all.length;
  updateCWkBtn(checked.length, all.length);
  cRender();
}
function updateCWkBtn(n, total){
  const btn=$('cwk-btn');
  btn.textContent = n===total ? 'All Weeks' : `${n} Week${n!==1?'s':''}`;
  btn.classList.toggle('on', n!==total);
}
function resetCWkFilter(){
  document.querySelectorAll('.cwk-cb').forEach(c=>c.checked=true);
  const allBox=$('cwk-all');
  if(allBox){ allBox.checked=true; allBox.indeterminate=false; }
  selectedCreatorWeeks=null;
  updateCWkBtn(C.meta.week_labels.length, C.meta.week_labels.length);
}
document.addEventListener('click', e=>{
  if(!e.target.closest('#cwk-wrap')){ const dd=$('cwk-dd'); if(dd) dd.style.display='none'; }
});

function cInit(){
  if(cInitDone) return;
  cInitDone = true;
  const cats = ['All', ...new Set(C.creators.map(c=>c.category))].sort((a,b)=>a==='All'?-1:b==='All'?1:a.localeCompare(b));
  populate($('cCAT'), cats);
  cBuildWkFilter();
  // Default: YT Creator filter on
  $('cCT_TYPE').value = 'YT Creator';
  // Default: Hide spend below ₹500 ticked
  $('cHIDE500').checked = true;
  // Default: hide Wk1–Wk8 (indices 0-7), show Wk9 onwards
  const total = C.meta.week_labels.length;
  const hideFirst = 8;
  if(total > hideFirst){
    document.querySelectorAll('.cwk-cb').forEach((cb,i)=>{ cb.checked = i >= hideFirst; });
    $('cwk-all').checked = false;
    $('cwk-all').indeterminate = true;
    const checked = [...Array(total).keys()].filter(i => i >= hideFirst);
    selectedCreatorWeeks = checked;
    updateCWkBtn(checked.length, total);
  }
  cBuildHeader();
  cRender();
}

// ═══════════════════════════════════════════════════════════════════════════
// INSIGHT TAB — JavaScript
// ═══════════════════════════════════════════════════════════════════════════
let insInitDone = false;
let insCat = 'All';

function buildVidNarrMap(){
  const map = {};
  C.creators.forEach(cr=>{ cr.videos.forEach(v=>{ map[v.id]={narr:v.narrative||'Unknown',cat:cr.category,creator:cr.name}; }); });
  return map;
}

const INS_MIN_SPEND = 1000; // Minimum spend to include a video in CAC analysis

function insNarrStats(cat){
  const stats = {};
  C.creators.forEach(cr=>{
    if(cat!=='All'&&cr.category!==cat) return;
    cr.videos.forEach(v=>{
      const narr=v.narrative||'Unknown';
      if(!stats[narr]) stats[narr]={spend:0,conv:0,count:0,scaleUp:0};
      const s=v.w.reduce((a,b)=>a+b,0), c=v.wc.reduce((a,b)=>a+b,0);
      if(s>=INS_MIN_SPEND){  // Only meaningful spend videos count for CAC
        stats[narr].spend+=s; stats[narr].conv+=c; stats[narr].count++;
        const n=v.w.length;
        if(n>=2&&v.w[n-1]>0&&v.w[n-1]>=v.w[n-2]*0.9&&c>0) stats[narr].scaleUp++;
      }
    });
  });
  return stats;
}

function calcMomentum(w,wc){
  const n=w.length;
  if(n<4) return 5;
  const l3s=w.slice(-3).reduce((a,b)=>a+b,0), p3s=w.slice(-6,-3).reduce((a,b)=>a+b,0);
  const l3c=wc.slice(-3).reduce((a,b)=>a+b,0), p3c=wc.slice(-6,-3).reduce((a,b)=>a+b,0);
  if(l3s===0) return 0;
  let score=5;
  if(p3s>0){const r=l3s/p3s; score+=r>1.5?2:r>1.0?1:r<0.5?-2:r<0.8?-1:0;} else score+=1;
  if(p3c>0){const r=l3c/p3c; score+=r>1.5?2:r>1.0?1:r<0.5?-2:r<0.8?-1:0;} else if(l3c>0) score+=1;
  if(w[n-1]>0) score+=1;
  return Math.max(0,Math.min(10,score));
}

function momColor(s){return s>=8?'#0e6b2f':s>=6?'#1a9940':s>=4?'#f0ad4e':s>=2?'#e67e22':'#c0392b';}

function insSetCat(cat,btn){
  insCat=cat;
  document.querySelectorAll('.ins-cat-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  insRender();
}

function insRender(){
  const cat=insCat;
  const vidNarr=buildVidNarrMap();
  const wk15=14, wk16=15;

  // Narrative table
  const ns=insNarrStats(cat);
  const rows=Object.entries(ns)
    .map(([n,d])=>({narr:n,...d,cac:d.conv>0?Math.round(d.spend/d.conv):0,scaleRate:d.count>0?Math.round(d.scaleUp/d.count*100):0}))
    .filter(d=>d.count>0&&d.narr!=='Unknown').sort((a,b)=>b.spend-a.spend);

  const byCac=[...rows].filter(r=>r.conv>0).sort((a,b)=>a.cac-b.cac);
  const untapped=[...rows].filter(r=>r.count<4&&r.conv>0&&r.cac<600).sort((a,b)=>a.cac-b.cac);
  if(byCac.length>0){$('ins-c1').textContent=byCac[0].narr.slice(0,18);$('ins-c1s').textContent=`₹${byCac[0].cac} CAC · ${byCac[0].count} videos`;}
  if(untapped.length>0){$('ins-c2').textContent=untapped[0].narr.slice(0,18);$('ins-c2s').textContent=`₹${untapped[0].cac} CAC · only ${untapped[0].count} tested`;}

  $('ins-narr-body').innerHTML=rows.map(r=>{
    const [bg,fg]=(r.cac>0&&r.cac<400?'#d4edda:#155724':r.cac<700?'#fff3cd:#856404':'#f8d7da:#721c24').split(':');
    const sig=r.cac>0&&r.cac<400?'🟢 Top CAC':r.scaleRate>40?'🟢 High Scale':r.cac<600?'✅ Stable':'⚠️ Monitor';
    const rid=`narr-${r.narr.replace(/[^a-z0-9]/gi,'_')}`;
    return `<tr id="${rid}" style="cursor:pointer" onclick="insExpandNarr('${r.narr}','${cat}','${rid}',this)">
      <td style="text-align:center;font-size:11px;color:var(--blue)" id="${rid}-icon">▸</td>
      <td style="text-align:left;font-weight:600">${r.narr}</td>
      <td>${r.count}</td><td class="spend">${inrS(r.spend)}</td><td>${Math.round(r.conv)}</td>
      <td><span style="background:${bg};color:${fg};padding:2px 8px;border-radius:8px;font-size:11px;font-weight:700">₹${r.cac}</span></td>
      <td>${r.scaleRate}%</td><td>${sig}</td>
    </tr>`;
  }).join('');

  // Momentum leaders
  const momRows=[];
  D.videos.forEach(v=>{
    const info=vidNarr[v.id];
    if(cat!=='All'&&(!info||info.cat!==cat)) return;
    const allW=v.combos.reduce((arr,c)=>{c.w.forEach((s,i)=>arr[i]=(arr[i]||0)+s);return arr;},[]);
    const allWc=v.combos.reduce((arr,c)=>{(c.wc||[]).forEach((s,i)=>arr[i]=(arr[i]||0)+s);return arr;},[]);
    const mom=calcMomentum(allW,allWc);
    const s15=allW[wk15]||0, s16=allW[wk16]||0;
    if((s15+s16)<500) return;
    momRows.push({id:v.id,url:v.url,cat:info?.cat||'–',narr:info?.narr||'–',creator:info?.creator||'–',mom,s15,s16,w:allW,wc:allWc});
  });
  momRows.sort((a,b)=>b.mom-a.mom);

  if(momRows.length>0){$('ins-c3').textContent=momRows[0].id;$('ins-c3s').textContent=`Score ${momRows[0].mom}/10 · ${momRows[0].narr}`;}

  $('ins-mom-body').innerHTML=momRows.slice(0,15).map(r=>{
    const thumb=`https://img.youtube.com/vi/${r.id}/mqdefault.jpg`;
    const barW=Math.round(r.mom/10*120);
    const vd=getVerdict(r.w,r.wc,null);
    const catCls=r.cat==='Hair'?'cb-hair':r.cat==='Beard'?'cb-beard':'cb-nutr';
    // Build tooltip
    const n=r.w.length;
    const l3s=r.w.slice(-3).reduce((a,b)=>a+b,0), p3s=r.w.slice(-6,-3).reduce((a,b)=>a+b,0);
    const l3c=r.wc.slice(-3).reduce((a,b)=>a+b,0), p3c=r.wc.slice(-6,-3).reduce((a,b)=>a+b,0);
    const spendTrend=p3s>0?Math.round((l3s/p3s-1)*100):100;
    const convTrend=p3c>0?Math.round((l3c/p3c-1)*100):100;
    const label=r.mom>=8?'Scaling fast 🚀':r.mom>=6?'Growing steadily 🟢':r.mom>=4?'Stable ✅':r.mom>=2?'Softening ⚠️':'Declining 🔴';
    const tip=`Score ${r.mom}/10 — ${label} | Spend trend: ${spendTrend>=0?'+':''}${spendTrend}% | Conv trend: ${convTrend>=0?'+':''}${convTrend}% | Creator: ${r.creator}`;
    return `<tr><td style="text-align:left;padding-left:10px"><div style="display:flex;align-items:center;gap:8px"><img src="${thumb}" style="width:52px;height:30px;border-radius:3px;object-fit:cover" onerror="this.style.visibility='hidden'" alt=""><div><a href="${r.url}" target="_blank" style="font-size:11px;font-weight:600;color:var(--navy);text-decoration:none">${r.id}</a><div style="font-size:10px;color:var(--muted)">${r.creator}</div></div></div></td><td style="text-align:left"><span class="cbadge ${catCls}">${r.cat}</span></td><td style="text-align:left;font-size:11px;color:var(--muted)">${r.narr}</td><td><div class="mom-bar-wrap vbadge" data-tip="${tip}" style="cursor:default"><div class="mom-bar" style="width:${barW}px;background:${momColor(r.mom)}"></div><span class="mom-num" style="color:${momColor(r.mom)}">${r.mom}/10</span></div></td><td class="spend" style="font-size:11px">${r.s15>0?inrS(r.s15):'–'}</td><td class="spend" style="font-size:11px">${r.s16>0?inrS(r.s16):'–'}</td><td><span class="vbadge ${vd.cls}" style="font-size:10px">${vd.label}</span></td></tr>`;
  }).join('');

  // Content gap card
  const allRows=insNarrStats('All');
  const allNarrsList=Object.entries(allRows).filter(([n,d])=>d.conv>0&&d.cac>0&&n!=='Unknown').sort((a,b)=>b[1].spend-a[1].spend).map(([n])=>n);
  const catNarrSet=new Set(Object.keys(insNarrStats(cat==='All'?'Hair':cat)));
  const gapNarr=allNarrsList.find(n=>!catNarrSet.has(n)||insNarrStats(cat==='All'?'Hair':cat)[n]?.count<2);
  if(gapNarr){$('ins-c4').textContent=gapNarr.slice(0,18);$('ins-c4s').textContent=`Underrepresented hook — brief needed`;}
}

function saveInsRemarks(){
  try{
    localStorage.setItem('mm_vb_remarks', $('ins-remarks').value);
    const s=$('ins-remarks-status');
    if(s){ s.textContent='✓ Saved'; setTimeout(()=>{ s.textContent='Notes auto-saved in browser.'; },1500); }
  }catch(e){}
}
function loadInsRemarks(){
  try{
    const saved=localStorage.getItem('mm_vb_remarks');
    if(saved&&$('ins-remarks')) $('ins-remarks').value=saved;
  }catch(e){}
}

function insGenerateBriefs(){
  const cat=insCat;
  const ns=insNarrStats(cat);
  const remarks=($('ins-remarks')?.value||'').trim();
  const briefs=[];
  const sorted=Object.entries(ns).map(([n,d])=>({narr:n,...d,cac:d.conv>0?Math.round(d.spend/d.conv):0})).filter(d=>d.count>0&&d.cac>0&&d.narr!=='Unknown').sort((a,b)=>a.cac-b.cac);

  if(sorted.length>0){const t=sorted[0];briefs.push({border:'#1a9940',tagBg:'#d4edda',tagFg:'#155724',tag:'🟢 Scale Up',title:`Double down on "${t.narr}"`,desc:`Best CAC hook at ₹${t.cac}${cat!=='All'?' in '+cat:''} across ${t.count} active video${t.count>1?'s':''}. Brief 2–3 more videos using this exact hook angle. Keep the same tone — the algo is already trained on it.`});}
  const untapped=sorted.filter(d=>d.count<4&&d.cac<600);
  if(untapped.length>0){const u=untapped[0];briefs.push({border:'#2E75B6',tagBg:'#cce5ff',tagFg:'#004085',tag:'🔬 Explore',title:`Test more "${u.narr}" — promising, under-tested`,desc:`₹${u.cac} CAC with only ${u.count} video${u.count>1?'s':''} tested. Brief a different creator on the same hook to see if the performance replicates before committing big budget.`});}
  const fatigue=sorted.filter(d=>d.spend>100000&&d.cac>650);
  if(fatigue.length>0){const f=fatigue[0];briefs.push({border:'#e67e22',tagBg:'#fff3cd',tagFg:'#856404',tag:'⚠️ Refresh',title:`Refresh "${f.narr}" — possible creative fatigue`,desc:`₹${inrS(f.spend)} spent but CAC is ₹${f.cac}. The hook is sound but existing creatives may be wearing out. Brief a fresh execution of the same hook with a new creator or new opening frame.`});}
  const knownNarrs=['Problem Based','Myth Busting','Testimonial','Before/After','Science','Genetics'];
  const allN=new Set(sorted.map(d=>d.narr));
  const miss=knownNarrs.filter(n=>!allN.has(n));
  if(miss.length>0){briefs.push({border:'#c0392b',tagBg:'#f8d7da',tagFg:'#721c24',tag:'🔴 Gap',title:`No active "${miss[0]}" videos${cat!=='All'?' in '+cat:''}`,desc:`This hook type is missing from your${cat!=='All'?' '+cat:''} pipeline. It works in other categories — worth testing here. One well-executed video on this angle could unlock a new audience segment.`});}

  // Append remarks as context if provided
  if(remarks){
    briefs.push({border:'#888780',tagBg:'#f1efe8',tagFg:'#444441',tag:'📝 Your Context',
      title:'Notes you added',
      desc:remarks
    });
  }

  $('ins-briefs').innerHTML=briefs.length===0
    ?'<div style="color:var(--muted);font-size:13px;padding:12px 0">Not enough data for this category. Select All or a category with more videos.</div>'
    :briefs.map(b=>`<div class="brief-card" style="border-left-color:${b.border}"><div style="margin-bottom:6px"><span class="brief-tag" style="background:${b.tagBg};color:${b.tagFg}">${b.tag}</span></div><h3>${b.title}</h3><p>${b.desc}</p></div>`).join('');
}

function insExpandNarr(narr, cat, rowId, rowEl){
  const existId = rowId+'-sub';
  const existing = document.getElementById(existId);
  const icon = document.getElementById(rowId+'-icon');
  if(existing){ existing.remove(); if(icon) icon.textContent='▸'; return; }
  if(icon) icon.textContent='▾';

  // Build creator data for this narrative
  const cData = {};
  C.creators.forEach(cr=>{
    if(cat!=='All'&&cr.category!==cat) return;
    cr.videos.forEach(v=>{
      if(v.narrative!==narr) return;
      const s=v.w.reduce((a,b)=>a+b,0), c=v.wc.reduce((a,b)=>a+b,0);
      if(s<INS_MIN_SPEND) return; // exclude low-spend videos from CAC
      if(!cData[cr.name]) cData[cr.name]={spend:0,conv:0,count:0,cat:cr.category};
      cData[cr.name].spend+=s; cData[cr.name].conv+=c; cData[cr.name].count++;
    });
  });
  const cRows=Object.entries(cData).map(([name,d])=>({name,...d,cac:d.conv>0?Math.round(d.spend/d.conv):0})).sort((a,b)=>a.cac-b.cac);
  if(cRows.length===0) return;

  const colSpan = 8;
  const subHtml=`<tr id="${existId}" style="background:#f0f4f9">
    <td colspan="${colSpan}" style="padding:0 0 6px 36px">
      <div style="border-left:3px solid var(--blue);padding-left:12px;margin:4px 0">
        <div style="font-size:11px;font-weight:700;color:var(--navy);margin:6px 0 4px">Creators running "${narr}"</div>
        <table style="width:80%;border-collapse:collapse;font-size:11px">
          <thead><tr>
            <th style="text-align:left;padding:3px 8px;color:var(--muted);font-weight:600;background:#e8edf5">Creator</th>
            <th style="background:#e8edf5">Videos</th>
            <th style="background:#e8edf5">Total Spend</th>
            <th style="background:#e8edf5">Conversions</th>
            <th style="background:#e8edf5">Avg CAC</th>
            <th style="background:#e8edf5">Signal</th>
          </tr></thead>
          <tbody>
            ${cRows.map(r=>{
              const [bg,fg]=(r.cac>0&&r.cac<400?'#d4edda:#155724':r.cac<700?'#fff3cd:#856404':'#f8d7da:#721c24').split(':');
              const sig=r.cac>0&&r.cac<400?'🟢 Best CAC':r.cac<600?'✅ Efficient':'⚠️ Review';
              return `<tr><td style="text-align:left;padding:4px 8px;font-weight:600;color:var(--navy)">${r.name}</td><td>${r.count}</td><td class="spend">${inrS(r.spend)}</td><td>${Math.round(r.conv)}</td><td><span style="background:${bg};color:${fg};padding:2px 8px;border-radius:8px;font-size:11px;font-weight:700">₹${r.cac}</span></td><td style="font-size:12px">${sig}</td></tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </td>
  </tr>`;
  rowEl.insertAdjacentHTML('afterend', subHtml);
}

function insightInit(){
  if(insInitDone) return;
  insInitDone=true;
  loadInsRemarks();
  insRender();
}

// ── Init ──────────────────────────────────────────────────────────────────────
(function init(){
  $('hdr-sub').textContent = `${D.meta.week_range}  ·  ${D.meta.total_videos} unique videos tracked`;
  $('hdr-badge').textContent = `Generated: ${D.meta.generated}`;
  populate($('sCT'), D.opts.cts);
  buildWkFilter();
  // Default: hide Wk1–Wk8 (indices 0-7), show Wk9 onwards
  const total = D.meta.week_labels.length;
  const hideFirst = 8;
  if(total > hideFirst){
    document.querySelectorAll('.wk-cb').forEach((cb,i)=>{ cb.checked = i >= hideFirst; });
    $('wk-all').checked = false;
    $('wk-all').indeterminate = true;
    const checked = [...Array(total).keys()].filter(i => i >= hideFirst);
    selectedWeeks = checked;
    updateWkBtn(checked.length, total);
  }
  buildHeader();
  onCTChange();
})();
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────────────────────
# 5. INJECT DATA & SAVE
# ─────────────────────────────────────────────────────────────────────────────
rights_16_9 = build_16_9_dates()
rights_16_9_json = json.dumps(rights_16_9, ensure_ascii=False)
html = HTML.replace('__DATA__', json_str).replace('__CREATORS_DATA__', creators_json_str).replace('__RIGHTS_16_9_DATA__', rights_16_9_json)
out  = 'MM_YT_Dashboard 0807 V2.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'✅  Saved → {out}')
print(f'    Videos: {len(videos)}  |  Creators: {len(creators_list)}  |  Weeks: {len(week_labels)}  |  Size: {os.path.getsize(out)//1024} KB')

# ─────────────────────────────────────────────────────────────────────────────
# 6. AUTO-PATCH CUSTOM TABS (tabs 4-8, access control, drag, AI Generate)
# ─────────────────────────────────────────────────────────────────────────────
import subprocess, sys as _sys
_patch = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'patch_custom.py')
if os.path.exists(_patch):
    print('\nApplying custom tab patch…')
    result = subprocess.run([_sys.executable, _patch], capture_output=True, text=True)
    print(result.stdout.strip())
    if result.returncode != 0:
        print('⚠ Patch error:', result.stderr.strip())
else:
    print('⚠ patch_custom.py not found — custom tabs not applied')
