"""Download the latest Kobo export (XLSX, labels, all versions) to the path given."""
import os, sys, requests

url, token = os.environ['KOBO_EXPORT_URL'], os.environ['KOBO_TOKEN']
r = requests.get(url, headers={'Authorization': 'Token ' + token}, timeout=180)
r.raise_for_status()
if not r.content.startswith(b'PK'):
    sys.exit('Kobo did not return an XLSX file. Check KOBO_EXPORT_URL and KOBO_TOKEN. First bytes: %r' % r.content[:200])
open(sys.argv[1], 'wb').write(r.content)
print('saved', len(r.content), 'bytes')
