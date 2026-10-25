"""Turn the Kobo export of the Investor Survey into the JSON the dashboard reads."""
import json, re, sys, datetime as dt
from collections import Counter
import openpyxl

NOW = dt.datetime.utcnow() + dt.timedelta(hours=6)  # Dhaka time

SRC, GEO, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
wb = openpyxl.load_workbook(SRC)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))
H, R = rows[0], [r for r in rows[1:] if any(v not in (None, '') for v in r)]

OWN_FIX = {'Foreign / MNC ( / (MNC))': 'Foreign / MNC', 'MNC ( )': 'MNC'}
BN = re.compile(r'[ঀ-৿]+')
def clean(v):
    if v is None: return None
    s = str(v)
    s = BN.sub('', s)
    s = re.sub(r'^\s*\d+\s*=\s*', '', s)
    s = s.replace(';', '').replace('\n', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    s = OWN_FIX.get(s, s)
    for _ in range(3):
        s = re.sub(r'\([\s/,\-]*\)', '', s).strip()
        m = re.search(r'\(([^()]*)\)\s*$', s)
        if m and m.group(1).strip() and m.group(1).strip().lower() in s[:m.start()].lower():
            s = s[:m.start()].strip()
    s = re.sub(r'\s+', ' ', s).strip(' ,')
    return s or None

def num(v):
    if v in (None, ''): return None
    try: return float(str(v).replace(',', '').strip())
    except ValueError: return None

def money(v):
    """Tk value from free text. Returns (value or None, problem or None)."""
    if v in (None, ''): return None, None
    s = str(v).strip().lower().replace(',', '')
    if re.search(r'us|\$', s): return None, 'Value given in US dollars'
    m = re.match(r'^([\d.]+)\s*(crore|cr|lakh|lac)?\s*(bdt|tk|taka)?$', s)
    if not m: return None, 'Value is not a number'
    x = float(m.group(1)) * {'crore': 1e7, 'cr': 1e7, 'lakh': 1e5, 'lac': 1e5, None: 1}[m.group(2)]
    if x == 0: return None, 'Value is zero'
    return x, None

# ---------- question schema ----------
SEV = ['none', 'minor', 'moderate', 'serious', 'very serious']
Q = []  # each: key, col, section, label, type, suffix
def add(key, col, sec, label, typ='single', **kw):
    Q.append(dict(key=key, col=col, section=sec, label=label, type=typ, **kw))

A, B, C, E, F, G = ('A. Firm profile', 'B. Trade and imports', 'C. Tariff and customs',
                    'E. Investment climate', 'F. Effects on investment', 'G. Reform priorities and outlook')
add('div', 5, A, 'Location (division)')
add('zone', 8, A, 'Zone status')
add('pos', 9, A, 'Respondent position')
add('recontact', 10, A, 'Agreed to be re-contacted')
add('own', 11, A, 'Investment type by ownership')
add('sector', 13, A, 'Main sector')
add('year', 7, A, 'Year operations began in Bangladesh', 'num', unit='year')
add('perm', 15, A, 'Permanent workers at present', 'num', unit='workers')
add('temp', 16, A, 'Temporary workers at present', 'num', unit='workers')
add('exp_dir', 17, A, 'Share of sales exported directly (last FY)', ordinal=True)
add('exp_ind', 18, A, 'Share of sales as deemed or indirect exports', ordinal=True)
add('assets', 19, A, 'Value of fixed assets, excluding land and building', ordinal=True)
add('turnover', 20, A, 'Annual turnover, FY2025-26', ordinal=True)
add('imp', 21, B, 'Imported goods or equipment in the last 12 months')
add('port', 22, B, 'Main port or station of entry for imports')
add('imp_share', 23, B, 'Share of input and equipment spending that is imported', ordinal=True, suffix='%')
add('scheme', 40, B, 'Use of bonded warehouse, drawback, exemption or zone schemes')
add('duty_cost', 41, B, 'Import duties and taxes as a share of production cost', ordinal=True)
add('imp_value', 42, B, 'Annual value of imports, last FY', ordinal=True)
add('consign', 43, B, 'Import consignments (bills of entry), last 12 months', ordinal=True)
add('burden', 44, B, 'Border-tax burden compared with three years ago',
    order=['Much lower', 'Lower', 'Same', 'Higher', 'Much higher'])
add('duty_paid', 46, B, 'Total duties and taxes paid at import, last FY', ordinal=True)
add('trade_app', 59, B, 'Traded or applied for a customs, tariff or investment process in the last 3 years')
C1 = [(61, 'Tariff on imported raw materials'), (62, 'Tariff on intermediate inputs'),
      (63, 'Tariff on capital machinery'), (64, 'Input duty as high as duty on competing finished goods'),
      (65, 'Supplementary Duty (SD) burden'), (66, 'Regulatory Duty (RD) burden'),
      (67, 'Delay adjusting VAT, AT and AIT paid at import'), (68, 'HS classification uncertainty or dispute'),
      (69, 'Customs clearance delay'), (70, 'Frequent SRO or policy changes'),
      (71, 'Complex bonded, drawback or exemption procedures'), (72, 'Minimum import or tariff value above price paid')]
for i, (c, l) in enumerate(C1):
    add('c1_%d' % (i + 1), c, C, l, 'sev', order=SEV, grid='c1')
add('mech', 73, C, 'Main way the largest problem affects the firm')
add('clear_time', 74, C, 'Routine customs clearance time at the main port', ordinal=True, suffix=' days')
add('disputes', 75, C, 'HS classification or valuation disputes, last 12 months (number)', ordinal=True)
add('info', 76, C, 'Ease of getting reliable advance information on tariffs')
add('bonded_eff', 77, C, 'Effectiveness of bonded access',
    order=['very ineffective', 'ineffective', 'neutral', 'effective', 'Very effective', 'not applicable (N/A)'])
add('drawback_time', 78, C, 'Typical duty drawback or refund time', ordinal=True, suffix=' months')
add('obstacle', 79, C, 'Main obstacle to using duty-free import schemes')
add('unrefunded', 80, C, 'Duty or tax paid but not yet refunded', ordinal=True)
add('guarantee', 81, C, 'Amount held as bank guarantee or bond with customs', ordinal=True)
E1 = [(83, 'Land and site access'), (84, 'Power reliability and cost'), (85, 'Transport, logistics and ports'),
      (86, 'Finance, trade finance and foreign exchange'), (87, 'Regulatory approvals and licences'),
      (88, 'Coordination between government agencies'), (89, 'Shortage of skilled labour')]
for i, (c, l) in enumerate(E1):
    add('e1_%d' % (i + 1), c, E, l, 'sev', order=SEV, grid='e1')
add('invest3', 90, F, 'Total investment FY2023-24 to FY2025-26, excluding land', ordinal=True)
F1 = [(92, 'Delayed an investment'), (93, 'Deferred an investment'), (94, 'Reduced the scale of a planned investment'),
      (95, 'Cancelled an investment'), (96, 'Used a less-preferred sourcing option'), (97, 'Lost export orders or competitiveness')]
add('f1', None, F, 'Effects of tariff or customs issues in the last 3 years', 'multi',
    cols=F1, none_label='None of these')
G1 = [(99, 'Cut tariffs on selected inputs'), (100, 'Simplify tariff bands, reduce escalation'),
      (101, 'Better HS guidance and advance rulings'), (102, 'Valuation transparency and dispute resolution'),
      (103, 'Faster clearance through digital reform'), (104, 'Simpler bonded, drawback and refund procedures'),
      (105, 'Advance notice and consultation on SRO changes'), (106, 'NBR, MoC, BTTC and IB coordination'),
      (107, 'Adjustment support for own sector')]
add('g1', None, G, 'Three most urgent reforms', 'multi', cols=G1, none_label='None selected')
add('confidence', 108, G, 'Confidence that tariff and customs rules stay the same over two years',
    order=['not at all confident', 'less confident', 'moderately confident', 'confident', 'very confident', 'DK'])
add('consult', 109, G, 'Effectiveness of government and business consultation on tariff changes',
    order=['very ineffective', 'ineffective', 'moderately effective', 'effective', 'very effective', 'no experience'])
add('overall', 110, G, 'Overall tariff and customs environment for investment',
    order=['Very constraining', 'Constraining', 'neutral', 'Supportive', 'Very supportive'])

LAST = re.compile(r"^(dk|don't know|dk / refused|varies|not applicable \(n/a\)|no experience)$", re.I)
def code_of(raw):
    m = re.match(r'^\s*(\d+)\s*=', str(raw)) if raw is not None else None
    return int(m.group(1)) if m else None

def label_for(q, raw):
    s = clean(raw)
    if s is None: return None
    if q['type'] == 'sev': s = s.lower()
    suf = q.get('suffix')
    if suf and re.search(r'\d$', s) and not LAST.match(s): s = s + suf
    return s

# ---------- geography ----------
geo = json.load(open(GEO))
def pip(x, y, ring):
    ins, j = False, len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i]; xj, yj = ring[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi: ins = not ins
        j = i
    return ins
def which_div(lon, lat):
    for f in geo['features']:
        g = f['geometry']; polys = g['coordinates'] if g['type'] == 'MultiPolygon' else [g['coordinates']]
        for p in polys:
            if pip(lon, lat, p[0]): return f['properties']['name']
    return None

DHAKA_PLACES = re.compile(r'dhanmondi|motijheel|gulistan|chankharpool|paltan|polton|^dh\d*$', re.I)
UPAZILA = {'mirsharai': 'Chattogram'}

# ---------- records ----------
recs, issues = [], []
opt_codes = {q['key']: {} for q in Q}
name_seen = {}
def issue(rec, text, kind='bad', typ=''):
    issues.append(dict(id=rec['id'], firm=rec['firm'], issue=text, kind=kind, type=typ, enum=rec['enum'], date=rec['date']))

for r in R:
    start, end = r[0], r[1]
    rec = dict(id=str(r[131]), idx=int(r[139]) if r[139] else None,
               firm=re.sub(r'\s+', ' ', str(r[4] or '')).strip(),
               date=start.strftime('%Y-%m-%d') if start else None,
               start=start.strftime('%Y-%m-%d %H:%M') if start else None,
               mins=round((end - start).total_seconds() / 60, 1) if start and end else None,
               enum=(r[136] or '').strip() or 'Not recorded',
               version=str(r[137] or ''), product=clean(r[14]) or '',
               a={})
    for q in Q:
        if q['type'] == 'multi':
            picks = [lab for c, lab in q['cols'] if str(r[c]).strip() == '1']
            rec['a'][q['key']] = picks if picks else [q['none_label']]
        elif q['type'] == 'num':
            rec['a'][q['key']] = num(r[q['col']])
        else:
            raw = r[q['col']]
            lab = label_for(q, raw)
            rec['a'][q['key']] = lab
            if lab is not None and code_of(raw) is not None:
                opt_codes[q['key']][lab] = code_of(raw)
    # district as entered vs cleaned
    d_raw = re.sub(r'\s+', ' ', str(r[6] or '')).strip()
    d = d_raw
    if DHAKA_PLACES.search(d_raw): d = 'Dhaka'
    elif d_raw.lower() in UPAZILA: d = UPAZILA[d_raw.lower()]
    rec['dist'] = d; rec['dist_raw'] = d_raw
    # group
    own = rec['a']['own'] or ''
    rec['grp'] = 'Domestic' if re.match(r'domestic|state', own, re.I) else 'Foreign or JV'
    rec['inzone'] = (rec['a']['zone'] or '').lower() != 'outside zones'
    rec['a']['age'] = (NOW.year - rec['a']['year']) if rec['a']['year'] else None
    rec['a']['workers'] = (rec['a']['perm'] or 0) + (rec['a']['temp'] or 0) if rec['a']['perm'] is not None else None
    # products
    imps = []
    for pc, vc, hc in [(29, 30, 31), (33, 34, 35), (37, 38, 39)]:
        if r[pc] in (None, ''): continue
        v, prob = money(r[vc])
        hs = re.sub(r'\D', '', str(r[hc] or ''))
        imps.append(dict(name=clean(r[pc]), value=v, raw=str(r[vc]).strip() if r[vc] is not None else '', hs=hs))
        if prob: issue(rec, 'Import value for "%s": %s (%s)' % (clean(r[pc]), prob.lower(), str(r[vc]).strip()), 'warn', 'Product value format')
        if hs and len(hs) < 4: issue(rec, 'HS code %s for "%s" is shorter than 4 digits' % (hs, clean(r[pc])), 'warn', 'HS code')
        if not hs: issue(rec, 'No HS code for imported "%s"' % clean(r[pc]), 'warn', 'HS code')
    exps = []
    for pc, hc, vc in [(52, 53, 54), (56, 57, 58)]:
        if r[pc] in (None, ''): continue
        v, prob = money(r[vc])
        hs = re.sub(r'\D', '', str(r[hc] or ''))
        exps.append(dict(name=clean(r[pc]), value=v, raw=str(r[vc]).strip() if r[vc] is not None else '', hs=hs))
        if prob: issue(rec, 'Export value for "%s": %s (%s)' % (clean(r[pc]), prob.lower(), str(r[vc]).strip()), 'warn', 'Product value format')
        if hs and len(hs) < 4: issue(rec, 'HS code %s for "%s" is shorter than 4 digits' % (hs, clean(r[pc])), 'warn', 'HS code')
    rec['imports'], rec['exports'] = imps, exps
    for i in range(3): rec['a']['imp_v%d' % (i + 1)] = imps[i]['value'] if i < len(imps) else None
    for i in range(2): rec['a']['exp_v%d' % (i + 1)] = exps[i]['value'] if i < len(exps) else None
    rec['suggest'] = clean(r[111]) or ''
    lat, lon, prec = num(r[113]), num(r[114]), num(r[116])
    rec['gps'] = [lon, lat] if lat and lon else None
    rec['gps_prec'] = prec

    # ---- data checks ----
    if rec['mins'] is not None and rec['mins'] < 10:
        issue(rec, 'Interview took %s minutes (under 10)' % rec['mins'], 'bad', 'Very short interview')
    if rec['mins'] is not None and rec['mins'] > 240:
        issue(rec, 'Form was open for %s hours' % round(rec['mins'] / 60, 1), 'warn', 'Form left open')
    if d != d_raw:
        issue(rec, 'District entered as "%s" (counted as %s)' % (d_raw, d), 'warn', 'District entry')
    if rec['gps']:
        gd = which_div(lon, lat)
        if prec == 0: issue(rec, 'GPS precision is 0, so the location was probably typed in', 'warn', 'GPS')
        if gd is None: issue(rec, 'GPS point falls outside Bangladesh', 'bad', 'GPS')
        elif rec['a']['div'] and gd != rec['a']['div']:
            issue(rec, 'GPS point is in %s but the form says %s' % (gd, rec['a']['div']), 'bad', 'GPS')
    else:
        issue(rec, 'No GPS location', 'bad', 'GPS')
    if rec['grp'] == 'Foreign or JV' and r[12] in (None, ''):
        issue(rec, 'Foreign ownership share is blank for a foreign or joint-venture firm', 'warn', 'Ownership share blank')
    if rec['a']['imp'] == 'no' and imps:
        issue(rec, 'Says no imports in 12 months but lists imported products', 'bad', 'Contradiction')
    if rec['a']['exp_dir'] == '0%' and rec['a']['exp_ind'] == '0%' and exps:
        issue(rec, 'Reports 0% exports but lists export products', 'warn', 'Contradiction')
    s = rec['suggest']
    if s.lower() in ('yes', 'no'):
        issue(rec, 'Answered "%s" but named no tariff line, SRO or procedure' % s, 'warn', 'Open answer')
    elif s and (len(s) < 5 or re.match(r'^(open record|hh+n+)', s, re.I)):
        issue(rec, 'Open answer looks like a placeholder: "%s"' % s, 'warn', 'Open answer')
    if rec['a']['perm'] is not None and rec['a']['perm'] == 0:
        issue(rec, 'Zero permanent workers', 'warn', 'Workers')
    key = re.sub(r'[^a-z]', '', rec['firm'].lower())
    if key in name_seen:
        issue(rec, 'Same firm name as record %s, possible duplicate' % name_seen[key], 'bad', 'Possible duplicate')
    else:
        name_seen[key] = rec['id']
    recs.append(rec)

# form-version label conflicts
for q in Q:
    if q.get('ordinal'):
        by_code = {}
        for lab, c in opt_codes[q['key']].items(): by_code.setdefault(c, set()).add(lab)
        clash = {c: sorted(v) for c, v in by_code.items() if len(v) > 1}
        if clash:
            for c, labs in clash.items():
                issues.append(dict(id='', firm='All records', enum='', date='', kind='warn', type='Form version wording',
                    issue='"%s": answer code %d has different wording across form versions (%s)' % (q['label'], c, ' / '.join(labs))))

# option order
def order_opts(q):
    vals = Counter()
    for rc in recs:
        v = rc['a'][q['key']]
        for x in (v if isinstance(v, list) else [v]):
            if x is not None: vals[x] += 1
    if q['type'] == 'multi':
        labs = [l for _, l in q['cols']] + [q['none_label']]
        return labs
    if q.get('order'):
        o = list(q['order']); low = [x.lower() for x in o]
        extra = [k for k in vals if k.lower() not in low]
        return o + sorted(extra, key=lambda k: -vals[k])
    if q.get('ordinal'):
        codes = opt_codes[q['key']]
        def k(lab):
            if lab in codes: return (1, codes[lab], lab)
            if lab.lower() == 'none': return (0, 0, lab)
            if LAST.match(lab): return (3, 0, lab)
            return (2, -vals[lab], lab)
        return sorted(vals, key=k)
    return sorted(vals, key=lambda k: (bool(LAST.match(k)), -vals[k], k))

schema = []
for q in Q:
    item = dict(key=q['key'], section=q['section'], label=q['label'], type=q['type'])
    if q['type'] != 'num': item['options'] = order_opts(q)
    if q.get('unit'): item['unit'] = q['unit']
    if q.get('grid'): item['grid'] = q['grid']
    schema.append(item)
NUMS = [('year', A, 'Year operations began in Bangladesh', 'year'), ('age', A, 'Years operating in Bangladesh', 'years'),
        ('perm', A, 'Permanent workers at present', 'workers'), ('temp', A, 'Temporary workers at present', 'workers'),
        ('workers', A, 'All workers (permanent + temporary)', 'workers'),
        ('imp_v1', B, 'Import value, top imported product 1', 'Tk'), ('imp_v2', B, 'Import value, top imported product 2', 'Tk'),
        ('imp_v3', B, 'Import value, top imported product 3', 'Tk'),
        ('exp_v1', B, 'Export value, main export product 1', 'Tk'), ('exp_v2', B, 'Export value, main export product 2', 'Tk')]
schema = [s for s in schema if s['type'] != 'num'] + [dict(key=k, section=s, label=l, type='num', unit=u) for k, s, l, u in NUMS]

versions = Counter(rc['version'] for rc in recs)
last_sub = max(str(r[133]) for r in R if r[133])
out = dict(title='Investor Survey 2026', updated=max(rc['start'] for rc in recs),
           last_submission_utc=last_sub, exported=(sys.argv[4] if len(sys.argv) > 4 else NOW.strftime('%Y-%m-%d %H:%M')),
           schema=schema, records=recs, issues=issues,
           versions=[dict(v=v, n=n, first=min(rc['date'] for rc in recs if rc['version'] == v),
                          last=max(rc['date'] for rc in recs if rc['version'] == v)) for v, n in versions.most_common()])
json.dump(out, open(OUT, 'w'), ensure_ascii=False, separators=(',', ':'))
print('records', len(recs), 'issues', len(issues))
for i in issues: print(' -', i['id'], i['issue'])
print(Counter(rc['grp'] for rc in recs))
