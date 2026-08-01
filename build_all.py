#!/usr/bin/env python3
"""
Regenerate owner portals whose data changed since the last run.
Run daily by GitHub Actions (.github/workflows/update-portals.yml).

Inputs:  turo.xlsx (downloaded by the workflow), state.json, hosting.json
Outputs: <slug>/index.html for each changed owner, updated state.json/hosting.json
"""
import json, os, re, sys, uuid, datetime
import portal_generator as pg

# Sanity check: make sure the download is a real spreadsheet, not a Google
# error/login page (happens if the sheet's link sharing gets turned off).
with open(pg.XLSX, 'rb') as f:
    if f.read(2) != b'PK':
        sys.exit('ERROR: turo.xlsx is not a valid spreadsheet. '
                 'Check that the Google Sheet is shared as "Anyone with the link: Viewer".')

today = datetime.datetime.now()
roster, stmts, exps, trips = pg.load(today)
new_state = pg.snapshot(roster, stmts, exps, trips)

old_state = json.load(open('state.json')) if os.path.exists('state.json') else {}
hosting = json.load(open('hosting.json')) if os.path.exists('hosting.json') else {}

changed = []
for owner, info in new_state.items():
    already_published = owner in hosting and os.path.exists(os.path.join(hosting.get(owner, ''), 'index.html'))
    if old_state.get(owner, {}).get('hash') == info['hash'] and already_published:
        continue
    if owner not in hosting:
        slug = re.sub(r'[^a-z0-9]+', '-', owner.lower()).strip('-') + '-' + uuid.uuid4().hex[:8]
        hosting[owner] = slug
    slug = hosting[owner]
    try:
        html = pg.render(owner, roster, stmts, exps, trips, today)
    except SystemExit as e:
        print(f'skip {owner}: {e}')
        continue
    os.makedirs(slug, exist_ok=True)
    with open(os.path.join(slug, 'index.html'), 'w') as f:
        f.write(html)
    changed.append(owner)

with open('state.json', 'w') as f:
    json.dump(new_state, f, indent=1)
with open('hosting.json', 'w') as f:
    json.dump(hosting, f, indent=1, sort_keys=True)

print('updated:', ', '.join(changed) if changed else 'none')
