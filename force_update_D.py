"""
Force-update D (YT Performance) and C (Creators) for Week of 7th September.
Used when update_asset_tabs.py already added the week to TA_WEEKS,
so update_dashboard.py exits early thinking nothing is new.
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import everything from update_dashboard
from update_dashboard import (
    load_html, save_html, update_D, update_C, log, parse_week_date
)

html = load_html()

# The tab we need to add to D
new_tabs = [("Week of 7th September", parse_week_date("Week of 7th September"))]

log("Force-updating D (YT Performance) for Week of 7th September...")
html = update_D(html, new_tabs)

log("Updating C (Creators)...")
html = update_C(html)

save_html(html)
log("Done. Hard-refresh (Cmd+Shift+R) to see changes.")
