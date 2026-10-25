"""Create a fresh Kobo XLS export (same options as the Downloads page), wait for it, save it.

Works for a shared user: it only needs permission to view submissions, plus that user's API token.
Env: KOBO_TOKEN (secret), KOBO_SERVER, KOBO_ASSET_UID.
"""
import os, sys, time, requests

server = os.environ.get('KOBO_SERVER', 'https://eu.kobotoolbox.org').rstrip('/')
uid = os.environ['KOBO_ASSET_UID']
H = {'Authorization': 'Token ' + os.environ['KOBO_TOKEN'], 'Accept': 'application/json'}
base = '%s/api/v2/assets/%s/exports/' % (server, uid)

# Same choices as the export the dashboard was built from:
# XLS, English labels, no group names in headers, all versions, select-multiple as single + separate columns.
opts = {'type': 'xls', 'lang': 'English (en)', 'hierarchy_in_labels': False, 'group_sep': '/',
        'fields_from_all_versions': True, 'multiple_select': 'both', 'include_media_url': True}

r = requests.post(base, json=opts, headers=H, timeout=60)
if r.status_code in (401, 403):
    sys.exit('Kobo refused the request (%s). Check KOBO_TOKEN and that this account can view submissions.' % r.status_code)
r.raise_for_status()
exp = r.json()
url = exp.get('url') or base + exp['uid'] + '/'

for _ in range(60):  # up to about 5 minutes
    exp = requests.get(url, headers=H, timeout=60).json()
    if exp.get('status') == 'complete':
        break
    if exp.get('status') == 'error':
        sys.exit('Kobo export failed: %s' % exp.get('messages'))
    time.sleep(5)
else:
    sys.exit('Kobo export did not finish in time.')

f = requests.get(exp['result'], headers={'Authorization': H['Authorization']}, timeout=180)
f.raise_for_status()
if not f.content.startswith(b'PK'):
    sys.exit('Kobo did not return an XLSX file.')
open(sys.argv[1], 'wb').write(f.content)
print('saved', len(f.content), 'bytes')

requests.delete(url, headers=H, timeout=60)  # keep the Downloads list tidy
