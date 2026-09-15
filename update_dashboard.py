"""
update_dashboard.py
===================
Full weekly automation for MM_YT_Dashboard 0807 V2.html

Updates all 4 data tabs in one run:
  - Tab 1: YT Performance (const D)
  - Tab 2: Creators       (const C)
  - Tab 5: Text Assets    (TA_WEEKS / TA_RAW)
  - Tab 6: Image Assets   (IA_WEEKS / IA_RAW)

Runs every Monday at 12:30 PM via Claude scheduled task.
Previous week data is ALWAYS preserved — new weeks are only appended.

Sheet: https://docs.google.com/spreadsheets/d/1J6pog_JxT7oJ2jFA4koRw_0eIamcLQ9MYxnFJXrq-J4
"""

import urllib.request, json, re, os, datetime as dt, sys
from urllib.parse import quote
from collections import defaultdict

# ─── CONFIG ──────────────────────────────────────────────────────────────────
SHEET_ID  = '1J6pog_JxT7oJ2jFA4koRw_0eIamcLQ9MYxnFJXrq-J4'
DASHBOARD = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         'MM_YT_Dashboard 0807 V3.html')
LOG_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         'update_dashboard.log')

TEXT_TYPES  = ['Headline', 'Description', 'Long headline']
IMAGE_TYPES = ['Horizontal image', 'Square image', 'Vertical image']
YT_TYPES    = ['YouTube video']

MONTH_MAP = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
             'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
MONTHS    = ['January','February','March','April','May','June',
             'July','August','September','October','November','December']
MON_SHORT = ['Jan','Feb','Mar','Apr','May','Jun',
             'Jul','Aug','Sep','Oct','Nov','Dec']

# ─── LOGGING ─────────────────────────────────────────────────────────────────
def log(msg):
    ts = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')

# ─── GVIZ ────────────────────────────────────────────────────────────────────
def gviz(sheet_name, tq, timeout=30):
    url = (f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq'
           f'?tqx=out:json&sheet={quote(sheet_name)}&tq={quote(tq)}')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read().decode('utf-8')
    raw = raw[raw.index('{'):raw.rindex('}')+1]
    data = json.loads(raw)
    if data.get('status') == 'error':
        raise RuntimeError(f"GViz error on '{sheet_name}': {data.get('errors')}")
    return data['table']['rows']

def esc(s):
    return str(s).replace('\\','\\\\').replace('"','\\"')

def norm_adg(s):
    s = str(s).strip() if s else ''
    return '' if s in ('--', 'nan', 'None', '-') else s

# ─── TAB DISCOVERY ───────────────────────────────────────────────────────────
def make_tab_name(d):
    day = d.day
    suf = 'th' if 11 <= day <= 13 else {1:'st',2:'nd',3:'rd'}.get(day % 10,'th')
    return f'Week of {day}{suf} {MONTHS[d.month-1]}'

def tab_name_to_label(name):
    m = re.search(r'week of\s+(\d+)\w*\s+(\w+)', name.lower())
    return f"{m.group(1)} {m.group(2).capitalize()[:3]}" if m else name

def parse_week_date(name):
    m = re.search(r'week of\s+(\d+)\w*\s+(\w+)', name.lower())
    if not m: return 0
    day = int(m.group(1))
    mon = MONTH_MAP.get(m.group(2)[:3], 0)
    return dt.datetime.now().year * 10000 + mon * 100 + day if mon else 0

def week_label_for_D(tab_name):
    """Generate 'WkN (6Jul-12Jul)' style label from tab name + next week count."""
    m = re.search(r'week of\s+(\d+)\w*\s+(\w+)', tab_name.lower())
    if not m: return tab_name
    day = int(m.group(1))
    mon = MONTH_MAP.get(m.group(2)[:3], 0)
    if not mon: return tab_name
    start = dt.date(dt.datetime.now().year, mon, day)
    end   = start + dt.timedelta(days=6)
    s_str = f"{start.day}{MON_SHORT[start.month-1]}"
    e_str = f"{end.day}{MON_SHORT[end.month-1]}"
    return f"(${s_str}-{e_str})"  # WkN prefix added later

def get_fingerprint(sheet_name):
    try:
        rows = gviz(sheet_name, "select sum(M) where D='Headline' and M>0")
        if rows and rows[0]['c'][0]:
            return rows[0]['c'][0].get('v')
    except: pass
    return None

def discover_new_tabs(existing_labels):
    default_fp = get_fingerprint('__INVALID__')
    log(f"  Default fingerprint: {default_fp}")

    now = dt.datetime.now()
    # Probe weekly steps going back 6 months (~26 probes), plus ±2 days around
    # each step to catch any day-of-week variation. Deduplicate tab names.
    seen, candidates = set(), []
    for week in range(27):
        for offset in range(-2, 3):
            d = now - dt.timedelta(weeks=week, days=offset)
            name = make_tab_name(d)
            if name not in seen:
                seen.add(name)
                candidates.append((name, parse_week_date(name)))
    candidates.sort(key=lambda x: -x[1])

    distinct = []
    for name, dk in candidates:
        lbl = tab_name_to_label(name)
        if lbl in existing_labels: continue
        fp = get_fingerprint(name)
        if fp is not None and fp != default_fp:
            distinct.append((name, dk))
            log(f"  ✓ New tab: {name}")

    # Back-calculate oldest week (it's the default/first sheet tab)
    if distinct:
        oldest_dk   = min(dk for _, dk in distinct)
        oldest_name = next(n for n, dk in distinct if dk == oldest_dk)
        m = re.search(r'week of\s+(\d+)\w*\s+(\w+)', oldest_name.lower())
        if m:
            day = int(m.group(1)); mon = MONTH_MAP.get(m.group(2)[:3], 0)
            if mon:
                prev_d    = dt.date(now.year, mon, day) - dt.timedelta(weeks=1)
                prev_name = make_tab_name(dt.datetime(prev_d.year, prev_d.month, prev_d.day))
                prev_lbl  = tab_name_to_label(prev_name)
                if prev_lbl not in existing_labels:
                    fp = get_fingerprint(prev_name)
                    if fp is not None:
                        distinct.append((prev_name, parse_week_date(prev_name)))
                        log(f"  ✓ Oldest (default sheet): {prev_name}")

    return sorted(distinct, key=lambda x: x[1])

# ─── LOAD HTML ───────────────────────────────────────────────────────────────
def load_html():
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        return f.read()

def save_html(html):
    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(html)

def extract_json_block(html, var_name, after_var=None):
    """Extract a JSON object/array assigned to a JS const."""
    start_idx = html.index(f'const {var_name} = {{') + len(f'const {var_name} = ')
    after = html[start_idx:]
    depth=0; in_str=False; esc_next=False
    for i, ch in enumerate(after):
        if esc_next: esc_next=False; continue
        if ch=='\\' and in_str: esc_next=True; continue
        if ch=='"' and not esc_next: in_str=not in_str; continue
        if in_str: continue
        if ch=='{': depth+=1
        elif ch=='}':
            depth-=1
            if depth==0:
                return json.loads(after[:i+1]), start_idx, start_idx+i+1
    raise ValueError(f"Could not extract {var_name}")

# ─── TAB 1: YT PERFORMANCE ───────────────────────────────────────────────────
def fetch_yt(sheet_name):
    """Returns dict (url, camp, adg) -> [cost, conv]."""
    result = defaultdict(lambda: [0.0, 0.0])
    rows = gviz(sheet_name, "select C,E,F,M,R where D='YouTube video' and M>0")
    for row in rows:
        v = [c.get('v') if c else None for c in row['c']]
        url, camp, adg, cost, conv = v
        if not url or not camp or not cost: continue
        na = norm_adg(adg)
        key = (str(url), str(camp), na)
        result[key][0] += float(cost)
        result[key][1] += float(conv or 0)
    log(f"    YT rows: {len(rows)}, combos: {len(result)}")
    return result

def get_camp_type(c):
    if 'Pmax' in c: return 'Pmax'
    if c.startswith('UAC'): return 'UAC'
    if c.startswith('DG'): return 'DG'
    if c.startswith('Search'): return 'Search'
    return 'Other'

def get_category(c):
    cl = c.lower()
    if 'beard' in cl: return 'Beard'
    if 'hair' in cl: return 'Hair'
    if 'nutrition' in cl or 'shilajit' in cl or 'creatine' in cl: return 'Nutrition'
    return 'Other'

def update_D(html, new_tabs):
    D, d_start, d_end = extract_json_block(html, 'D')
    weeks   = D['meta']['week_labels']
    videos  = D['videos']
    n_weeks = len(weeks)

    # Ensure every video has a 'url' key (AI-added videos may only have 'id')
    for v in videos:
        if 'url' not in v:
            v['url'] = f"https://www.youtube.com/watch?v={v.get('id', '')}"

    # Build URL index for fast lookup
    url_idx = {v['url']: i for i, v in enumerate(videos)}
    # Also build (url, camp, norm_adg) -> combo index within video
    combo_idx = {}
    for vi, v in enumerate(videos):
        for ci, combo in enumerate(v['combos']):
            key = (v['url'], combo['c'], norm_adg(combo['ag']))
            combo_idx[key] = (vi, ci)

    for tab_name, _ in new_tabs:
        log(f"  Fetching YT data: {tab_name}...")
        yt_data = fetch_yt(tab_name)

        # Generate week label
        m = re.search(r'week of\s+(\d+)\w*\s+(\w+)', tab_name.lower())
        day = int(m.group(1)); mon = MONTH_MAP.get(m.group(2)[:3], 0)
        start = dt.date(dt.datetime.now().year, mon, day)
        end   = start + dt.timedelta(days=6)
        wk_lbl = f"Wk{n_weeks+1} ({start.day}{MON_SHORT[start.month-1]}-{end.day}{MON_SHORT[end.month-1]})"
        weeks.append(wk_lbl)
        new_idx = n_weeks
        n_weeks += 1

        # Extend all existing combos with a zero for the new week
        for v in videos:
            for combo in v['combos']:
                combo['w'].append(0.0)
                combo['wc'].append(0.0)

        # Fill in new week data — try exact match first, then (url, camp) fallback
        url_camp_lookup = defaultdict(lambda: [0.0, 0.0])
        for (url, camp, adg), vals in yt_data.items():
            url_camp_lookup[(url, camp)][0] += vals[0]
            url_camp_lookup[(url, camp)][1] += vals[1]

        new_videos = 0
        for (url, camp, adg), (cost, conv) in yt_data.items():
            exact_key = (url, camp, adg)
            if exact_key in combo_idx:
                vi, ci = combo_idx[exact_key]
                videos[vi]['combos'][ci]['w'][new_idx]  = round(cost, 2)
                videos[vi]['combos'][ci]['wc'][new_idx] = round(conv, 4)
            elif url in url_idx:
                # Video exists but combo doesn't — check by (url, camp) match
                vi = url_idx[url]
                matched = False
                for ci, combo in enumerate(videos[vi]['combos']):
                    if combo['c'] == camp:
                        combo['w'][new_idx]  = round(cost, 2)
                        combo['wc'][new_idx] = round(conv, 4)
                        combo_idx[exact_key] = (vi, ci)
                        matched = True
                        break
                if not matched:
                    # New combo on existing video
                    new_combo = {
                        'ct': get_camp_type(camp),
                        'cat': get_category(camp),
                        'c': camp,
                        'ag': adg if adg else '--',
                        'w':  [0.0] * (n_weeks - 1) + [round(cost, 2)],
                        'wc': [0.0] * (n_weeks - 1) + [round(conv, 4)]
                    }
                    videos[vi]['combos'].append(new_combo)
                    combo_idx[exact_key] = (vi, len(videos[vi]['combos'])-1)
            else:
                # Brand new video
                vid_id = re.search(r'[?&]v=([^&]+)', url)
                vid_id = vid_id.group(1) if vid_id else url.split('/')[-1]
                new_vid = {
                    'url': url, 'id': vid_id, 'total': 0,
                    'combos': [{
                        'ct': get_camp_type(camp),
                        'cat': get_category(camp),
                        'c': camp,
                        'ag': adg if adg else '--',
                        'w':  [0.0] * (n_weeks - 1) + [round(cost, 2)],
                        'wc': [0.0] * (n_weeks - 1) + [round(conv, 4)]
                    }],
                    'dim': '', 'fmt': ''
                }
                url_idx[url] = len(videos)
                combo_idx[exact_key] = (len(videos), 0)
                videos.append(new_vid)
                new_videos += 1

        log(f"    New videos added: {new_videos}")

    # Recalculate totals
    for v in videos:
        v['total'] = round(sum(w for combo in v['combos'] for w in combo['w']), 2)

    D['meta']['week_labels'] = weeks
    D['meta']['total_videos'] = len(videos)
    D['meta']['generated'] = dt.datetime.now().strftime('%d %b %Y')
    D['videos'] = videos

    d_json = json.dumps(D, separators=(',', ':'))
    html = html[:d_start] + d_json + html[d_end:]
    log(f"  ✓ Tab 1 updated: {len(videos)} videos, {len(weeks)} weeks")
    return html

# ─── TAB 2: CREATORS ─────────────────────────────────────────────────────────
def update_C(html):
    # Re-extract D (already updated)
    D, _, _ = extract_json_block(html, 'D')
    weeks   = D['meta']['week_labels']
    n_weeks = len(weeks)
    wk_idx  = n_weeks - 1

    # Build url -> wk_idx spend/conv from D
    url_spend = defaultdict(float)
    url_conv  = defaultdict(float)
    for v in D['videos']:
        for combo in v['combos']:
            if wk_idx < len(combo['w']):
                url_spend[v['url']] += combo['w'][wk_idx]
                url_conv[v['url']]  += combo['wc'][wk_idx]

    # Extract C
    c_start_idx = html.index('const C = {') + len('const C = ')
    after = html[c_start_idx:]
    depth=0; in_str=False; esc_c=False
    for i, ch in enumerate(after):
        if esc_c: esc_c=False; continue
        if ch=='\\' and in_str: esc_c=True; continue
        if ch=='"' and not esc_c: in_str=not in_str; continue
        if in_str: continue
        if ch=='{': depth+=1
        elif ch=='}':
            depth-=1
            if depth==0:
                C = json.loads(after[:i+1])
                c_end_abs = c_start_idx + i + 1
                break

    # Extend creator arrays if needed and update latest week
    for cr in C['creators']:
        # Extend if shorter than D weeks
        while len(cr['w'])  < n_weeks: cr['w'].append(0.0)
        while len(cr['wc']) < n_weeks: cr['wc'].append(0.0)
        creator_sp = 0.0; creator_cv = 0.0
        for vid in cr.get('videos', []):
            while len(vid['w'])  < n_weeks: vid['w'].append(0.0)
            while len(vid['wc']) < n_weeks: vid['wc'].append(0.0)
            sp = round(url_spend.get(vid['url'], 0), 2)
            cv = round(url_conv.get(vid['url'], 0), 4)
            vid['w'][wk_idx]  = sp
            vid['wc'][wk_idx] = cv
            vid['total'] = round(sum(vid['w']), 2)
            creator_sp += sp; creator_cv += cv
        cr['w'][wk_idx]  = round(creator_sp, 2)
        cr['wc'][wk_idx] = round(creator_cv, 4)
        cr['total'] = round(sum(cr['w']), 2)

    c_json = json.dumps(C, separators=(',', ':'))
    html = html[:c_start_idx] + c_json + html[c_end_abs:]
    log(f"  ✓ Tab 2 updated: {len(C['creators'])} creators")
    return html

# ─── TABS 5 & 6: TEXT / IMAGE ASSETS ─────────────────────────────────────────
def fetch_assets(sheet_name, asset_types):
    result = defaultdict(lambda: [0.0, 0.0])
    for atype in asset_types:
        try:
            rows = gviz(sheet_name, f"select C,D,E,F,M,R where D='{atype}' and M>0")
        except Exception as e:
            log(f"    Warning: {atype} failed: {e}"); continue
        for row in rows:
            v = [c.get('v') if c else None for c in row['c']]
            asset, typ, camp, adg, cost, conv = v
            if not asset or not camp or not cost: continue
            na = norm_adg(adg)
            key = (str(asset), str(typ), str(camp), na)
            result[key][0] += float(cost)
            result[key][1] += float(conv or 0)
    return result

def read_raw_const(html, const_name):
    m = re.search(r'const ' + const_name + r'\s*=\s*(\[[^\]]*\])', html)
    weeks = json.loads(m.group(1)) if m else []
    raw_m = re.search(r'const ' + const_name.replace('WEEKS','RAW') + r'\s*=\s*\[(.*?)\n\];', html, re.DOTALL)
    rows = {}
    if raw_m:
        for line in raw_m.group(1).strip().splitlines():
            line = line.strip().rstrip(',')
            if not line or not line.startswith('['): continue
            try:
                row = json.loads(line)
                rows[(row[0], row[1], row[2], row[3])] = row[4:]
            except: pass
    return weeks, rows

def merge_assets(existing_weeks, existing_rows, new_week_data):
    n = len(existing_weeks)
    all_weeks = existing_weeks + [lbl for lbl, _ in new_week_data]
    merged = {k: list(v) + [0]*(max(0, 2*n - len(v))) for k, v in existing_rows.items()}
    for wk_i, (lbl, week_dict) in enumerate(new_week_data):
        for key in list(merged.keys()):
            merged[key].extend(week_dict.get(key, [0, 0.0]))
        for key, (cost, conv) in week_dict.items():
            if key not in merged:
                backfill = [0] * (2*(n + wk_i))
                rest     = [0.0] * (2*(len(new_week_data)-wk_i-1))
                merged[key] = backfill + [round(cost), round(conv, 2)] + rest
    return all_weeks, merged

def build_raw_js(all_weeks, merged, weeks_name, raw_name):
    rows = sorted(merged.items(), key=lambda kv: -sum(kv[1][i] for i in range(0, len(kv[1]), 2)))
    lines = ['["{0}","{1}","{2}","{3}",{4}]'.format(
        esc(a), esc(t), esc(c), esc(g), ','.join(str(v) for v in cols))
        for (a, t, c, g), cols in rows]
    js  = f'const {weeks_name} = {json.dumps(all_weeks)};\n'
    js += f'const {raw_name} = [\n' + ',\n'.join(lines) + '\n];'
    return js

START_MARKER = ('// ═══════════════════════════════════════════════════════════════════════════\n'
                '// TEXT ASSETS & IMAGE ASSETS — DATA\n'
                '// ═══════════════════════════════════════════════════════════════════════════\n')

def update_assets(html, new_tabs):
    ta_weeks, ta_rows = read_raw_const(html, 'TA_WEEKS')
    ia_weeks, ia_rows = read_raw_const(html, 'IA_WEEKS')
    log(f"  Existing: TA {len(ta_weeks)} weeks/{len(ta_rows)} rows, IA {len(ia_weeks)} weeks/{len(ia_rows)} rows")

    new_ta, new_ia = [], []
    for tab_name, _ in new_tabs:
        lbl = tab_name_to_label(tab_name)
        log(f"  Fetching asset data: {tab_name}...")
        ta_d = fetch_assets(tab_name, TEXT_TYPES)
        ia_d = fetch_assets(tab_name, IMAGE_TYPES)
        log(f"    TA: {len(ta_d)} rows, IA: {len(ia_d)} rows")
        new_ta.append((lbl, ta_d))
        new_ia.append((lbl, ia_d))

    all_ta_weeks, merged_ta = merge_assets(ta_weeks, ta_rows, new_ta)
    all_ia_weeks, merged_ia = merge_assets(ia_weeks, ia_rows, new_ia)

    ta_js = build_raw_js(all_ta_weeks, merged_ta, 'TA_WEEKS', 'TA_RAW')
    ia_js = build_raw_js(all_ia_weeks, merged_ia, 'IA_WEEKS', 'IA_RAW')

    new_block = START_MARKER + ta_js + '\n\n' + ia_js
    block_start = html.index(START_MARKER)
    ia_open  = html.index('const IA_RAW = [', block_start)
    ia_close = html.index('\n];', ia_open) + 3
    html = html[:block_start] + new_block + html[ia_close:]
    log(f"  ✓ Tab 5 updated: {len(merged_ta)} TA rows across {all_ta_weeks}")
    log(f"  ✓ Tab 6 updated: {len(merged_ia)} IA rows across {all_ia_weeks}")
    return html

# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    log("=" * 60)
    log("MM Dashboard Weekly Update")
    log("=" * 60)

    html = load_html()

    # Read existing week labels from TA_WEEKS (simplest proxy for what's loaded)
    m = re.search(r'const TA_WEEKS\s*=\s*(\[[^\]]*\])', html)
    existing_labels = json.loads(m.group(1)) if m else []
    log(f"Existing weeks in dashboard: {existing_labels}")

    log("\nDiscovering new weekly tabs...")
    new_tabs = discover_new_tabs(existing_labels)

    if not new_tabs:
        log("Dashboard is already up to date — no new weeks found.")
        log("=" * 60)
        return

    log(f"New weeks to add: {[tab_name_to_label(n) for n, _ in new_tabs]}")

    log("\n[Tab 1 & 2] Updating YT Performance + Creators...")
    html = update_D(html, new_tabs)
    html = update_C(html)

    log("\n[Tab 5 & 6] Updating Text + Image Assets...")
    html = update_assets(html, new_tabs)

    save_html(html)
    log(f"\n✅ All done! Dashboard updated with {[tab_name_to_label(n) for n, _ in new_tabs]}")
    log("Hard-refresh the dashboard (Cmd+Shift+R) to see new data.")
    log("=" * 60)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        log(f"❌ ERROR: {e}")
        log(traceback.format_exc())
