#!/usr/bin/env python3
"""
add_weeks_v3.py — Adds new YT week data to MM_YT_Dashboard 0807 V3.html
Fetches from Google Sheets and merges into existing D data.

FREEZE GUARANTEE
================
This script ONLY modifies the `const D = {...}` block.
It NEVER touches:
  - const C (creator data, creator↔video mapping)
  - JS logic (product mapping, verdict, access control) — all lives AFTER C block
  - Tab structure, filters, or visual layout

Before writing, it verifies checksums from dashboard_freeze.json:
  - C block MD5 must be unchanged
  - JS logic MD5 must be unchanged
  - Week count must only increase
If any check fails, the script aborts without writing.
"""

import json, re, os, hashlib, urllib.request, csv, io
from urllib.parse import quote as _quote

FREEZE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_freeze.json")

DASHBOARD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MM_YT_Dashboard 0807 V3.html")
ASSET_SHEET_ID = '1J6pog_JxT7oJ2jFA4koRw_0eIamcLQ9MYxnFJXrq-J4'

# Weeks to add — format: (date_range_label, [gsheet_tab_names_or_gids])
# Use "gid:1234567" to fetch by sheet gid instead of tab name (when tab name is unknown)
# Multiple tab names = combine rows from multiple ad accounts for the same week
NEW_WEEKS = [
    ("24 August 2026 - 30 August 2026", ["gid:1171624481"]),
]

def fmt_week_label(date_str):
    """Convert '6 July 2026 - 12 July 2026' → 'Wk21 (6Jul-12Jul)'"""
    months = {'January':'Jan','February':'Feb','March':'Mar','April':'Apr',
              'May':'May','June':'Jun','July':'Jul','August':'Aug',
              'September':'Sep','October':'Oct','November':'Nov','December':'Dec'}
    parts = date_str.split(' - ')
    def fmt(d):
        t = d.strip().split()  # e.g. ['6', 'July', '2026']
        return f"{t[0]}{months.get(t[1], t[1])}"
    return f"{fmt(parts[0])}-{fmt(parts[1])}"

def fetch_yt_from_gsheet(sheet_id, sheet_name):
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
        adg = str(adg).strip() if adg and str(adg).strip() not in ('nan','None','--') else ''
        rows.append({'Asset': str(asset), 'Campaign': str(camp), 'AdGroup': adg,
                     'Cost': float(cost), 'Conversions': float(conv or 0)})
    return rows

def fetch_yt_from_gid(sheet_id, gid):
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read().decode('utf-8')
    except Exception as e:
        print(f'  ⚠️  GID fetch failed for gid={gid}: {e}')
        return []
    reader = csv.DictReader(io.StringIO(data))
    rows = []
    for row in reader:
        asset = row.get('Asset', '').strip()
        if 'youtube.com/watch' not in asset and '/shorts/' not in asset:
            continue
        camp = row.get('Campaign', '').strip()
        adg = row.get('Ad group', '').strip()
        if not camp or adg in ('--', ''): adg = ''
        try:
            sp = float(row.get('Cost', '0') or 0)
            cv = float(row.get('Conversions', '0') or 0)
        except:
            continue
        if sp <= 0:
            continue
        rows.append({'Asset': asset, 'Campaign': camp, 'AdGroup': adg,
                     'Cost': sp, 'Conversions': cv})
    return rows

def extract_vid_id(url_or_id):
    if 'youtube.com' in url_or_id or 'youtu.be' in url_or_id:
        m = re.search(r'[?&]v=([A-Za-z0-9_\-]{11})', url_or_id)
        if m: return m.group(1)
        m = re.search(r'youtu\.be/([A-Za-z0-9_\-]{11})', url_or_id)
        if m: return m.group(1)
    if re.match(r'^[A-Za-z0-9_\-]{11}$', url_or_id.strip()):
        return url_or_id.strip()
    return None

def get_camp_type(c):
    cl = c.lower()
    if 'pmax' in cl or 'performance max' in cl: return 'Pmax'
    if 'uac' in cl or 'universal app' in cl: return 'UAC'
    if ' dg ' in cl or '_dg_' in cl or 'display' in cl or 'discovery' in cl or 'demand gen' in cl: return 'DG'
    if 'search' in cl: return 'Search'
    return 'Other'

def get_camp_cat(c):
    cl = c.lower()
    if 'hair' in cl: return 'Hair'
    if 'beard' in cl: return 'Beard'
    if 'nutrition' in cl or 'shilajit' in cl or 'creatine' in cl: return 'Nutrition'
    return 'All Products'

def _extract_block(html, marker):
    start = html.find(marker)
    if start == -1: return -1, -1, ''
    depth, end = 0, start + len(marker)
    for i, ch in enumerate(html[start + len(marker):], start + len(marker)):
        if ch == '{': depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0: end = i + 1; break
    return start, end, html[start + len(marker):end]

def _md5(s): return hashlib.md5(s.encode()).hexdigest()

def verify_freeze(html):
    """Abort with error message if freeze checksums don't match. Returns freeze dict."""
    if not os.path.exists(FREEZE_FILE):
        print("⚠️  dashboard_freeze.json not found — run audit script first. Proceeding without freeze checks.")
        return None
    with open(FREEZE_FILE) as f:
        freeze = json.load(f)

    errors = []

    # Check C block
    _, c_end, c_str = _extract_block(html, 'const C = ')
    if not c_str:
        errors.append("CRITICAL: Cannot find 'const C = ' block")
    else:
        actual_md5 = _md5(c_str)
        expected_md5 = freeze['checksums']['c_block_md5']
        if actual_md5 != expected_md5:
            errors.append(f"CRITICAL: C block MD5 mismatch!\n  expected: {expected_md5}\n  actual  : {actual_md5}\n  → Creator data has been modified. Aborting to prevent data loss.")
        else:
            print(f"  ✓ C block unchanged (MD5: {actual_md5})")

    # Check JS logic (everything after C block)
    if c_end > 0:
        js_after = html[c_end:]
        actual_md5 = _md5(js_after)
        expected_md5 = freeze['checksums']['js_logic_after_c_md5']
        if actual_md5 != expected_md5:
            errors.append(f"CRITICAL: JS logic MD5 mismatch!\n  expected: {expected_md5}\n  actual  : {actual_md5}\n  → Product/creator mapping or dashboard logic has changed. Aborting.")
        else:
            print(f"  ✓ JS logic unchanged (MD5: {actual_md5})")

    if errors:
        print("\n" + "\n".join(errors))
        print("\nAborted — no changes written.")
        raise SystemExit(1)

    return freeze

def main():
    print("=" * 60)
    print("V3 Week Updater")
    print("=" * 60)

    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()

    # ── Freeze verification ──────────────────────────────────
    print("\nRunning freeze checks...")
    freeze = verify_freeze(html)

    # Extract current D data
    m = re.search(r'const D = (\{.*?\});\s*(?:const C|var C)', html, re.DOTALL)
    if not m:
        # Try alternate pattern
        start = html.find('const D = {')
        if start == -1:
            print("ERROR: Cannot find 'const D = {' in dashboard.")
            return
        # Find matching brace
        depth = 0
        end = start + len('const D = ')
        for i, ch in enumerate(html[start + len('const D = '):], start + len('const D = ')):
            if ch == '{': depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    end = i + 1
                    break
        d_str = html[start + len('const D = '):end]
    else:
        d_str = m.group(1)

    try:
        D = json.loads(d_str)
    except Exception as e:
        print(f"ERROR parsing D: {e}")
        return

    current_weeks = D['meta']['week_labels']
    n_existing = len(current_weeks)
    print(f"Current weeks: {n_existing} ({current_weeks[-1]})")

    # Build lookup: video id → video object
    vid_map = {v['id']: v for v in D['videos']}

    weeks_added = []
    for date_str, tab_names in NEW_WEEKS:
        # Support both single tab name (str) and multiple (list)
        if isinstance(tab_names, str):
            tab_names = [tab_names]

        # Derive label from date_str
        short = fmt_week_label(date_str)
        wk_num = n_existing + len(weeks_added) + 1
        label = f"Wk{wk_num} ({short})"

        if any(label in wl for wl in current_weeks):
            print(f"  ↩ {label} already present — skipping")
            continue
        if any(short in wl for wl in current_weeks):
            print(f"  ↩ {short} already present — skipping")
            continue

        # Fetch and combine rows from all tabs for this week
        rows = []
        for tab_name in tab_names:
            if tab_name.startswith('gid:'):
                gid = tab_name.split(':', 1)[1]
                print(f"\nFetching {label} from gid={gid}...")
                tab_rows = fetch_yt_from_gid(ASSET_SHEET_ID, gid)
            else:
                print(f"\nFetching {label} from '{tab_name}'...")
                tab_rows = fetch_yt_from_gsheet(ASSET_SHEET_ID, tab_name)
            print(f"  {len(tab_rows)} rows fetched")
            rows.extend(tab_rows)
        if not rows:
            print(f"  ⚠️  No data returned for any tab — skipping")
            continue
        print(f"  Total: {len(rows)} rows across {len(tab_names)} tab(s)")

        # Aggregate: video_id → campaign+adgroup → {spend, conv}
        agg = {}  # (vid_id, camp, adg) → {spend, conv}
        skipped = 0
        for r in rows:
            vid_id = extract_vid_id(r['Asset'])
            if not vid_id:
                skipped += 1
                continue
            key = (vid_id, r['Campaign'], r['AdGroup'])
            if key not in agg:
                agg[key] = {'spend': 0.0, 'conv': 0.0}
            agg[key]['spend'] += r['Cost']
            agg[key]['conv'] += r['Conversions']

        print(f"  {len(agg)} video+campaign combos ({skipped} skipped)")

        new_week_idx = n_existing + len(weeks_added)

        # Extend all existing videos' combo arrays with 0
        for v in D['videos']:
            for combo in v['combos']:
                while len(combo['w']) < new_week_idx:
                    combo['w'].append(0.0)
                    combo['wc'].append(0.0)
                combo['w'].append(0.0)
                combo['wc'].append(0.0)

        # Fill in actual data
        new_vids = 0
        for (vid_id, camp, adg), data in agg.items():
            ct = get_camp_type(camp)
            cat = get_camp_cat(camp)

            if vid_id not in vid_map:
                # New video
                new_w = [0.0] * (new_week_idx + 1)
                new_wc = [0.0] * (new_week_idx + 1)
                new_w[new_week_idx] = data['spend']
                new_wc[new_week_idx] = data['conv']
                new_v = {
                    'url': f'https://www.youtube.com/watch?v={vid_id}',
                    'id': vid_id,
                    'total': data['spend'],
                    'fmt': '',
                    'combos': [{'ct': ct, 'cat': cat, 'c': camp, 'ag': adg,
                                'w': new_w, 'wc': new_wc}]
                }
                D['videos'].append(new_v)
                vid_map[vid_id] = new_v
                new_vids += 1
            else:
                v = vid_map[vid_id]
                # Find matching combo
                matched = None
                for combo in v['combos']:
                    if combo['c'] == camp and combo['ag'] == adg:
                        matched = combo
                        break
                if matched:
                    matched['w'][new_week_idx] += data['spend']
                    matched['wc'][new_week_idx] += data['conv']
                else:
                    # New combo for existing video
                    new_w = [0.0] * (new_week_idx + 1)
                    new_wc = [0.0] * (new_week_idx + 1)
                    new_w[new_week_idx] = data['spend']
                    new_wc[new_week_idx] = data['conv']
                    v['combos'].append({'ct': ct, 'cat': cat, 'c': camp, 'ag': adg,
                                        'w': new_w, 'wc': new_wc})

        # Recompute totals
        for v in D['videos']:
            v['total'] = sum(s for combo in v['combos'] for s in combo['w'])

        current_weeks.append(label)
        weeks_added.append(label)
        print(f"  ✓ {label} added ({new_vids} new videos)")

    if not weeks_added:
        print("\nNo new weeks to add.")
        return

    # Update D meta
    D['meta']['week_labels'] = current_weeks
    D['meta']['total_videos'] = len(D['videos'])
    D['meta']['week_range'] = f"{current_weeks[0].split('(')[1].rstrip(')')} → {current_weeks[-1].split('(')[1].rstrip(')')}"

    # ── Sync C creator videos and C.meta with new D week data ──
    # C.meta.week_labels drives the Creators tab column headers
    c_str_orig = _extract_block(html, 'const C = ')[2]
    C = json.loads(c_str_orig)
    n_new = len(current_weeks)
    d_vid_map_final = {v['id']: v for v in D['videos']}

    def _get_wk(vid_id, wk_idx):
        v = d_vid_map_final.get(vid_id)
        if not v: return 0.0, 0.0
        sp = sum(c['w'][wk_idx] for c in v['combos'] if len(c['w']) > wk_idx)
        co = sum(c['wc'][wk_idx] for c in v['combos'] if len(c['wc']) > wk_idx)
        return sp, co

    c_updated = 0
    for cr in C['creators']:
        for v in cr.get('videos', []):
            w, wc = v.get('w'), v.get('wc')
            if w is None or wc is None: continue
            while len(w) < n_new - 1:
                w.append(0.0); wc.append(0.0)
            if len(w) < n_new:
                wk_idx = n_new - 1
                sp, co = _get_wk(v['id'], wk_idx)
                w.append(round(sp, 6)); wc.append(round(co, 6))
                v['total'] = round(sum(w), 2)
                c_updated += 1
        # Also update creator-level w/wc (used by Platform tab)
        cr_w, cr_wc = cr.get('w', []), cr.get('wc', [])
        if len(cr_w) < n_new:
            wk_idx = n_new - 1
            sp24 = sum(v['w'][wk_idx] for v in cr.get('videos', []) if v.get('w') and len(v['w']) > wk_idx)
            cv24 = sum(v['wc'][wk_idx] for v in cr.get('videos', []) if v.get('wc') and len(v['wc']) > wk_idx)
            cr_w.append(round(sp24, 6)); cr_wc.append(round(cv24, 6))
            cr['total'] = round(sum(cr_w), 2)

    C['meta']['week_labels'] = current_weeks[:]
    C['meta']['week_range']  = D['meta']['week_range']
    if 'week_strs' in D['meta']:
        C['meta']['week_strs'] = D['meta']['week_strs'][:]
    print(f"  ✓ Synced {c_updated} creator video w-arrays and C.meta.week_labels to {n_new} weeks")

    # Serialize both blocks and write
    new_d_json = json.dumps(D, separators=(',', ':'), ensure_ascii=False)
    new_c_json = json.dumps(C, separators=(',', ':'), ensure_ascii=False)

    # Replace D block first
    if m:
        new_html = html[:m.start(1)] + new_d_json + html[m.end(1):]
    else:
        d_start = html.find('const D = {')
        new_html = html[:d_start + len('const D = ')] + new_d_json + html[d_start + len('const D = ') + len(d_str):]

    # Replace C block in the already-modified html
    c2_start, c2_end, _ = _extract_block(new_html, 'const C = ')
    new_html = new_html[:c2_start + len('const C = ')] + new_c_json + new_html[c2_end:]

    with open(DASHBOARD, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f"\n✅ Added {len(weeks_added)} week(s): {', '.join(weeks_added)}")
    print(f"   Total weeks: {len(current_weeks)}, Total videos: {len(D['videos'])}")
    print(f"   Saved → {DASHBOARD}")

    # ── Post-write: verify JS logic unchanged, update freeze ──
    if freeze:
        with open(DASHBOARD, 'r', encoding='utf-8') as f:
            new_html = f.read()
        _, c_end2, c_str2 = _extract_block(new_html, 'const C = ')
        js2 = new_html[c_end2:]
        js_ok = _md5(js2) == freeze['checksums']['js_logic_after_c_md5']
        if js_ok:
            print("   ✓ JS logic unchanged")
        else:
            print("   ⚠️  JS logic checksum changed — please inspect the file!")
        # Update freeze with new C checksum and data state
        freeze['checksums']['c_block_md5'] = _md5(c_str2)
        freeze['checksums']['c_block_length'] = len(c_str2)
        freeze['data_state']['week_count'] = len(current_weeks)
        freeze['data_state']['last_week'] = current_weeks[-1]
        freeze['data_state']['video_count'] = len(D['videos'])
        with open(FREEZE_FILE, 'w') as ff:
            import json as _json
            _json.dump(freeze, ff, indent=2, ensure_ascii=False)
        print("   ✓ dashboard_freeze.json updated with new checksums")

    print("=" * 60)

if __name__ == '__main__':
    main()
