"""
Pack loaders: feed Kareem's engine (moontech_ref.py, unchanged) from the MONTHLY DATA PACK
instead of the original workbook layouts.

Pack files (data/pack/, git-ignored, re-uploaded monthly):
  net_sales_by_client.xlsx     Net Sales & Orders by Client  -> per-client GMV (Net Revenue AED), orders
  orders_by_creator.xlsx       Orders by Creator             -> net orders per Contact ID per month
  ads_by_creator.xlsx          Ads by Creator                -> closed ads per month
  revenue_by_merchant.xlsx     Revenue by Merchant (model)   -> commission revenue per merchant (AED); its Adjustment row is ignored
  2026_model_pnl.csv           2026 Model PnL (published)    -> P&L lines (AED), Sales -> Net Profit (+ transfer pricing)

Sources per domain:
  2024-25 (k 0..23)   untouched: pipeline/history_base.json (snapshot of the Aug-26 build)
  2026    (k 24..31)  from the pack
  ROAS ladder         history_base roas_brands (pack has no ladder export yet)       -> to-do #17
  CAC spend           history_base (pack CAC tab is 2024-25 only; CAC is live in the dashboard)
  creators            2026 per-creator from the pack; status/history from history_base -> to-do #17

Rule #1: pure-ROAS brands found in Revenue by Merchant are discounted (ROAS revenue comes from the ladder);
migrated/overlap clients (DKNY, Izil, ...) stay in commission and the engine splits them at their Since month.
"""
import os, re, csv, json, datetime as dt
import openpyxl
import moontech_ref as R
from moontech_ref import kidx, NM

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PACK = os.path.join(ROOT, 'data', 'pack')
NF = 32                                  # Jan-24 .. Aug-26
K26 = {m: kidx(2026, m) for m in range(1, 9)}   # 2026-01..08 -> 24..31

def available(): return os.path.isdir(PACK) and os.path.exists(os.path.join(PACK, 'revenue_by_merchant.xlsx'))

_H = None
def hist():
    global _H
    if _H is None: _H = json.load(open(os.path.join(HERE, 'history_base.json')))
    return _H

def _rows(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True); ws = wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]; wb.close(); return rows
def _g(r, i): return r[i] if i < len(r) else None
def _ym(x): return (x.year, x.month) if isinstance(x, (dt.datetime, dt.date)) else None
def _key(s): return ''.join(c for c in str(s or '').lower() if c.isalnum())
def _num(v):
    if isinstance(v, (int, float)): return float(v)
    if v is None: return None
    s = str(v).replace('AED', '').replace(',', '').replace('$', '').strip()
    if s in ('', '-', '—'): return None
    neg = s.startswith('(') and s.endswith(')')
    try: f = float(s.strip('()'))
    except ValueError: return None
    return -f if neg else f

# ----------------------------------------------------------------------------- ladder (live ROAS ladder snapshot: data/pack/ladder.json; fallback history_base)
def _ladder_rows():
    """ladder.json is produced from the live receivables + tracker sheets by the dashboard's own ladder logic
    (pulled through the browser at pack-update time): rows = [brand, level, since, churn label, status, dropped, grid by month]."""
    p = os.path.join(PACK, 'ladder.json')
    if os.path.exists(p):
        L = json.load(open(p)); out = []
        for brand, level, since, churn, status, dropped, grid in L['rows']:
            monthly = [0.0] * NM
            for lab, v in zip(L['months'], grid):
                if v in ('', None): continue
                y, m = map(int, lab.split('-')); k = kidx(y, m)
                if 0 <= k < NM: monthly[k] += float(v)
            out.append({'name': brand, 'since': since, 'level': ('L%d' % level) if level else 'CPL', 'churn_month': _churn_ym(churn) if dropped else None,
                        'monthly': monthly, 'other_pricing': not level})
        return out, L.get('today')
    return [dict(b, other_pricing=bool(b.get('other_pricing'))) for b in hist()['roas_brands']], None

def _churn_ym(s):
    s = str(s or '').strip()
    if not s: return None
    if re.match(r'^\d{4}-\d{2}', s): return s[:7]
    mon = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    m = re.match(r'^([A-Za-z]{3,9})[\s\-/,]+(\d{2,4})$', s)
    if not m: return None
    mi = mon.index(m.group(1)[:3].lower()) + 1; y = int(m.group(2)); y = y + 2000 if y < 100 else y
    return '%d-%02d' % (y, mi)

def cut_k():
    """Model update date (control.json latest_month) as a month index: nothing after it enters the model, even when a live
    sheet already has later months."""
    ctl = json.load(open(os.path.join(HERE, 'control.json'))); y, m = map(int, ctl['latest_month'].split('-')); return kidx(y, m)

def load_ladder(kit, cfg):
    lvl_re = re.compile(cfg['roas_status']['other_pricing_levels_not_matching'])
    include = set(cfg['roas_status'].get('include_as_roas', []))
    rows, _ = _ladder_rows(); lad = {}; CUT = cut_k()
    for b in rows:
        if not b.get('since'): continue
        y, m = map(int, str(b['since'])[:7].split('-'))
        if kidx(y, m) > CUT: continue                                  # signed after the update date: not in the model yet
        pay = [0.0] * NM
        for k, v in enumerate(b['monthly'][:NM]): pay[k] = float(v or 0) if k <= CUT else 0.0
        ch = b.get('churn_month'); churn = None
        if ch:
            try: cy, cm = map(int, str(ch)[:7].split('-')); churn = kidx(cy, cm)
            except Exception: churn = None
            if churn is not None and churn > CUT: churn = None          # churn after the update date is not yet known
        level = b['level']
        lm = level_monthly(b['name'])
        if lm is not None and lm[CUT] is not None and str(level).startswith('L'): level = 'L%d' % lm[CUT]   # level reached by the update date
        other = (bool(b.get('other_pricing')) or not lvl_re.match(str(level))) and b['name'] not in include
        lad[b['name']] = {'since': kidx(y, m), 'pay': pay, 'churn': churn, 'level': level, 'other_pricing': other}
    return lad

def _roas_brand_names():
    rows, _ = _ladder_rows(); return [b['name'] for b in rows]

def _migrated_names():
    return {b['name'] for b in hist()['roas_brands'] if b.get('migrated')}        # DKNY, Izil: commission clients that moved to ROAS
def _pure_roas_keys():
    mig = _migrated_names(); return {_key(n) for n in _roas_brand_names() if n not in mig}
def _strip_tags(name):
    s = re.sub(r'\s*-\s*MT$|\s+MT$', '', str(name).strip(), flags=re.I)      # "Joli - MT", "Goldenscent MT"
    s = re.sub(r'^Kitopi\s*-\s*', '', s, flags=re.I)                           # "Kitopi - Rightbite MT"
    return s
def _is_pure_roas(name, pure):
    import difflib
    k = _key(_strip_tags(name))
    if not k: return False
    for rk in pure:
        if k == rk: return True
        if (k in rk or rk in k) and min(len(k), len(rk)) >= 4: return True
        if difflib.SequenceMatcher(None, k, rk).ratio() >= 0.8: return True        # Naysa Beauty ~ Nysaa Beauty
    # first word of the ROAS brand name (Lauras ~ Laurus Aroma)
    mig = _migrated_names()
    for n in _roas_brand_names():
        if n in mig: continue
        w = _key(str(n).split()[0])
        if len(w) >= 5 and len(k) >= 5 and difflib.SequenceMatcher(None, k, w).ratio() >= 0.8: return True
    return False

# ----------------------------------------------------------------------------- merchant revenue (AED per merchant per month)
def load_merchant_revenue(kit, master, lad, cfg, source='system'):
    fx = cfg['fx']['aed_per_usd']; H = hist()
    rev = {}; adj = [0.0] * 32; review = []
    # 2024-25 from the baked build (USD -> AED), untouched
    for m in H['merchants']:
        rev[m['name']] = [float(m['monthly'][k] or 0) * fx if k < 24 else 0.0 for k in range(32)]
    for k in range(24):
        adj[k] = (float(H['rev_total'][k] or 0) - sum(float(m['monthly'][k] or 0) for m in H['merchants'])) * fx
    # 2026 from the pack
    rows = _rows(os.path.join(PACK, 'revenue_by_merchant.xlsx'), 'Revenue by Merchant (model)')
    hdr = rows[0]
    cols = {}
    for i in range(2, len(hdr)):
        ym = _ym(_g(hdr, i))
        if ym and ym[0] == 2026 and ym[1] in K26: cols[i] = K26[ym[1]]
    alias = {}
    for m, v in master.items():
        alias[m] = m
        for a in v.get('aliases', []): alias[a] = m
    for n in list(rev): alias.setdefault(n, n)
    pure = _pure_roas_keys()
    for r in rows[1:]:
        n = _g(r, 0)
        if not n: continue
        n = str(n).strip()
        if n.lower() in ('total', 'pnl total') or 'total' in n.lower(): continue
        if n == 'Adjustment': continue          # ignored: the Adjustment is computed (P&L Sales - tracked), not read
        if _is_pure_roas(n, pure):            # rule #1: ROAS revenue comes from the ladder
            continue
        if n in alias: m = alias[n]
        else:
            t, cand, sc = R.match_name(n, [x for x in master if master[x].get('source') == 'Datos'], cfg)
            if t in ('exact', 'contains'): m = cand
            else:
                t2, cand2, sc2 = R.match_name(n, list(lad), cfg)
                m = n
                review.append({'extract_name': n, 'merchant_match': (t, cand, sc), 'ladder_match': (t2, cand2, sc2),
                               'action': 'Proposed: new merchant' + (' linked to ' + cand2 if t2 != 'none' else ' (commission)')})
                master.setdefault(n, {'aliases': [], 'source': 'Datos', 'roas': cand2 if t2 != 'none' else None,
                                      'size_class': 'Unassigned', 'status': 'Proposed'})
            alias[n] = m
        rev.setdefault(m, [0.0] * 32)
        for i, k in cols.items():
            v = _num(_g(r, i))
            if v is not None: rev[m][k] += v
    return rev, adj, review

# ----------------------------------------------------------------------------- P&L (AED, k 0..31)
_PL_LINES = ['sales', 'creator_payouts', 'materials', 'other_cos', 'freelancers', 'bd_salaries', 'digital_marketing',
             'marketing_events', 'account_mgmt', 'c_suite', 'ops_salaries', 'advisory', 'development', 'system_expenses',
             'ga', 'da', 'finance_cost', 'tax', 'transfer_pricing', 'pl_gross_profit', 'pl_ebitda', 'pl_total_cos']
# ordered template: (label, line) — repeated labels (Moon/Infpact/Cost plus) are matched sequentially
_PL_TPL = [('sales', 'sales'), ('moon', None), ('infpact', None), ('cost plus', None),
           ('cost of sales', 'pl_total_cos'), ('moon', 'creator_payouts'), ('infpact', 'creator_payouts'), ('cost plus', 'creator_payouts'),
           ('materials', 'materials'), ('shipping and delivery expense', 'other_cos'), ('sales commission', 'other_cos'),
           ('freelancers (prospectors, watchers)', 'freelancers'),
           ('gross profit', 'pl_gross_profit'),
           ('business developmenet salaries', 'bd_salaries'), ('business development salaries', 'bd_salaries'),
           ('digital marketing', 'digital_marketing'), ('marketing events', 'marketing_events'), ('account management salaries', 'account_mgmt'),
           ('c-suite', 'c_suite'), ('operations salaries', 'ops_salaries'), ('advisory fees', 'advisory'),
           ('development', 'development'), ('system expenses', 'system_expenses'), ('g&a', 'ga'),
           ('ebitda', 'pl_ebitda'), ('depreciation & amortization', 'da'), ('finance cost', 'finance_cost'),
           ('corporate tax', 'tax'), ('transfer pricing expected cost', 'transfer_pricing')]
_MON = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']

def load_pl(kit, cfg):
    fx = cfg['fx']['aed_per_usd']; H = hist()
    pl = {l: [0.0] * 32 for l in _PL_LINES}
    # 2024-25 from the baked build (USD -> AED). Opex is stored grouped there; carry each group on one line so
    # group totals, EBITDA and the checks are unchanged (the dashboard shows groups only).
    hp = H['pl']
    direct = {'sales': 'sales', 'creator_payouts': 'creator_payouts', 'materials': 'materials', 'freelancers': 'freelancers',
              'other_cos': 'other_cos', 'development': 'development', 'da': 'da', 'finance_cost': 'finance_cost', 'tax': 'tax',
              'transfer_pricing': 'transfer_pricing'}
    grouped = {'opex_marketing': 'account_mgmt', 'opex_staff': 'ops_salaries', 'opex_technology': 'system_expenses', 'opex_ga': 'ga'}
    for k in range(24):
        for src, line in direct.items(): pl[line][k] = float(hp[src][k] or 0) * fx
        for src, line in grouped.items(): pl[line][k] = float(hp[src][k] or 0) * fx
    # 2026 from the pack P&L CSV (label-driven, sequential for repeated labels)
    rows = list(csv.reader(open(os.path.join(PACK, '2026_model_pnl.csv'), encoding='utf-8-sig')))
    norm = lambda s: str(s or '').strip().lower().replace(' ', ' ')
    hdr_i = next(i for i, r in enumerate(rows) if sum(1 for c in r if norm(c) in _MON) >= 3)
    mcols = {j: _MON.index(norm(c)) + 1 for j, c in enumerate(rows[hdr_i]) if norm(c) in _MON}
    pos = hdr_i + 1
    for label, line in _PL_TPL:
        found = next((i for i in range(pos, len(rows)) if rows[i] and norm(rows[i][0]) == label), -1)
        if found < 0: continue
        pos = found + 1
        if line is None: continue
        for j, mnum in mcols.items():
            if mnum in K26:
                v = _num(_g(rows[found], j))
                if v is not None: pl[line][K26[mnum]] += v
    return pl

# ----------------------------------------------------------------------------- creators (net orders per Contact ID, 2026 from pack)
def load_creators(kit, cfg):
    orders = {}; unatt = [0.0] * 32; flags = []
    valid = set(cfg['creator_id_checks']['valid_lengths'])
    rows = _rows(os.path.join(PACK, 'orders_by_creator.xlsx'), 'Export')
    h = rows[0]; ix = {str(h[i]).strip(): i for i in range(len(h)) if h[i] is not None}
    ci, mi, ni = ix['Contact ID'], ix['Month'], ix['Net Orders']
    for r in rows[1:]:
        ym = _ym(_g(r, mi))
        if not ym or ym[0] != 2026 or ym[1] not in K26: continue
        k = K26[ym[1]]; v = _num(_g(r, ni))
        if v is None: continue
        i = _g(r, ci)
        i = None if i in (None, '', 0, '0') else str(i).strip()
        if i is None: unatt[k] += v; continue
        if re.fullmatch(r'\d+\.0', i): i = i[:-2]
        if not i.isdigit() or len(i) not in valid: flags.append(('rejected_format', i, '2026-%02d' % ym[1])); continue
        orders.setdefault(i, [0.0] * 32)
        orders[i][k] += v                      # BI net orders are fractional (partial returns); summed raw they tie to the client export
    return orders, unatt, flags

# ----------------------------------------------------------------------------- CAC spend (USD per month) — from baked history
def load_cac(kit, cfg):
    H = hist()
    def pad(a): a = [(float(v) if isinstance(v, (int, float)) else None) for v in a]; return (a + [None] * NM)[:NM]
    return pad(H['cac_merchant_spend']), pad(H['cac_creator_spend'])

# ----------------------------------------------------------------------------- GMV (USD) + closed ads per month (NF)
def load_gmv_ads(kit):
    fx = hist()['fx']; H = hist()
    gmv = [(float(v) if isinstance(v, (int, float)) else None) for v in H['gmv']][:NF]
    ads = [(float(v) if isinstance(v, (int, float)) else None) for v in H['ads']][:NF]
    gmv += [None] * (NF - len(gmv)); ads += [None] * (NF - len(ads))
    # 2026 GMV = Net Revenue (AED) summed over clients, -> USD
    rows = _rows(os.path.join(PACK, 'net_sales_by_client.xlsx'), 'Export'); mrow, hrow = rows[0], rows[1]
    cols = [(i, _ym(_g(mrow, i)), _g(hrow, i)) for i in range(1, len(hrow))]
    g26 = {k: 0.0 for k in K26.values()}
    for r in rows[2:]:
        n = _g(r, 0)
        if not n or str(n).strip().lower() == 'total': continue
        for i, ym, met in cols:
            if ym and ym[0] == 2026 and ym[1] in K26 and met == 'Net Revenue (AED)':
                v = _num(_g(r, i))
                if v is not None: g26[K26[ym[1]]] += v / fx
    for k, v in g26.items(): gmv[k] = v
    # 2026 closed ads
    rows = _rows(os.path.join(PACK, 'ads_by_creator.xlsx'), 'Export'); h = rows[0]; ix = {str(h[i]).strip(): i for i in range(len(h)) if h[i] is not None}
    a26 = {k: 0.0 for k in K26.values()}
    for r in rows[1:]:
        ym = _ym(_g(r, ix['Month']))
        if ym and ym[0] == 2026 and ym[1] in K26:
            v = _num(_g(r, ix['Closed Ads']))
            if v is not None: a26[K26[ym[1]]] += v
    for k, v in a26.items(): ads[k] = v
    return gmv, ads

# ----------------------------------------------------------------------------- creators post-patch (status/history from baked build)
PACK_CREATOR_FIELDS = ('orders_total', 'orders_unatt')

def patch_creator_series(cr):
    """Pack has no 2024-25 per-creator history, so creator status metrics (active/new/churned/major...) come from
    the baked build; only the order totals for 2026 (orders_total / orders_unatt) are taken from the pack."""
    hc = hist()['creators']
    keep = {f: list(cr[f]) for f in PACK_CREATOR_FIELDS}
    for key, arr in hc.items(): cr[key] = list(arr)[:len(cr['active'])]
    for f in PACK_CREATOR_FIELDS:
        for k in range(24, 32): cr[f][k] = keep[f][k]

def patch_creator_periods(P):
    """Period creator blocks from the baked build, except the pack-driven order totals (already computed from the
    patched series by build.py)."""
    for pk, pc in hist()['periods_creators'].items():
        if pk in P and pc is not None:
            mine = P[pk].get('creators') or {}
            merged = dict(pc)
            for f in PACK_CREATOR_FIELDS:
                if f in mine: merged[f] = mine[f]
            P[pk]['creators'] = merged
    # creator order NRR needs per-creator history across both windows -> carried from the baked build
    for pk, v in hist().get('periods_retention_orders_12m', {}).items():
        if pk in P and isinstance(P[pk].get('retention'), dict) and P[pk]['retention'].get('orders_12m') is None:
            P[pk]['retention']['orders_12m'] = v

def install():
    R.load_ladder = load_ladder
    R.load_merchant_revenue = load_merchant_revenue
    R.load_pl = load_pl
    R.load_creators = load_creators
    R.load_cac = load_cac

# ----------------------------------------------------------------------------- segmentation / control extras (replace the interim bakes in index.html)
def _rbm_2026():
    """Revenue by Merchant export rows -> (name, {k: AED}) for 2026 Jan-Aug."""
    rows = _rows(os.path.join(PACK, 'revenue_by_merchant.xlsx'), 'Revenue by Merchant (model)'); hdr = rows[0]
    cols = {}
    for i in range(2, len(hdr)):
        ym = _ym(_g(hdr, i))
        if ym and ym[0] == 2026 and ym[1] in K26: cols[i] = K26[ym[1]]
    out = []
    for r in rows[1:]:
        n = _g(r, 0)
        if not n or 'total' in str(n).lower(): continue
        vals = {}
        for i, k in cols.items():
            v = _num(_g(r, i))
            if v is not None: vals[k] = v
        out.append((str(n).strip(), vals))
    return out

def comm_ttm(cfg):
    """Per commission client (export-name key): commission USD over 2026 Jan-Aug, active months, GMV USD (Net Revenue)."""
    fx = cfg['fx']['aed_per_usd']
    rows = _rows(os.path.join(PACK, 'net_sales_by_client.xlsx'), 'Export'); mrow, hrow = rows[0], rows[1]
    cols = [(i, _ym(_g(mrow, i)), _g(hrow, i)) for i in range(1, len(hrow))]
    gmv = {}
    for r in rows[2:]:
        n = _g(r, 0)
        if not n or str(n).strip().lower() == 'total': continue
        k = _key(n)
        for i, ym, met in cols:
            if ym and ym[0] == 2026 and ym[1] in K26 and met == 'Net Revenue (AED)':
                v = _num(_g(r, i))
                if v: gmv[k] = gmv.get(k, 0) + v / fx
    out = {}
    for n, vals in _rbm_2026():
        if n == 'Adjustment': continue
        k = _key(n); c = sum(v for v in vals.values() if v > 0) / fx; nm = sum(1 for v in vals.values() if v > 0)
        if not nm: continue
        g = gmv.get(k)
        if g is None:
            for gk in gmv:
                if (k in gk or gk in k) and min(len(k), len(gk)) >= 5: g = gmv[gk]; break
        out[k] = {'c': round(c), 'n': nm, **({'g': round(g)} if g is not None else {})}
    return out

def adj_indicator(E):
    """Invoicing adjustment, computed per 2026 month (USD) and identical to the Adjustment shown on every other page:
       P&L Sales (incl. ROAS) - ROAS ladder billing - commission tracked (named merchants, as the model counts them)
       = Adjustment (commission revenue in the P&L that merchant tracking does not yet show).
       Memo: ROAS brands still invoiced through merchant tracking (discounted from commission; part of the Adjustment).
       The export's own 'Adjustment' row is ignored everywhere."""
    fx = E.fx; pure = _pure_roas_keys()
    ks = list(K26.values()); spotted = {k: 0.0 for k in ks}
    for n, vals in _rbm_2026():
        if n == 'Adjustment' or not _is_pure_roas(n, pure): continue
        for k, v in vals.items():
            if k in spotted: spotted[k] += v / fx
    pnl = [E.pl['sales'][k] / fx for k in ks]; billing = [E.roas[k] + E.other[k] for k in ks]; tracked = [E.named[k] for k in ks]
    adj = [E.plug[k] for k in ks]
    return {'m': ['%02d' % m for m in K26], 'pnl': [round(x) for x in pnl], 'billing': [round(x) for x in billing],
            'tracked': [round(x) for x in tracked], 'adj': [round(x) for x in adj],
            'ind': [round(pnl[i] - billing[i]) for i in range(len(ks))], 'spotted': [round(spotted[k]) for k in ks]}

# ----------------------------------------------------------------------------- ROAS level history (milestone by payment month)
def level_monthly(brand):
    """Level reached by month k (max paid milestone with payment month <= k); None before the first payment.
    Source: data/pack/ladder_milestones.json (paid invoices: [milestone, payment YYYY-MM]) pulled with the ladder."""
    p = os.path.join(PACK, 'ladder_milestones.json')
    if not os.path.exists(p): return None
    MS = json.load(open(p)); inv = MS.get(brand)
    if inv is None: return None
    out = [None] * NM; cur = None
    byk = {}
    for ms, ymv in inv:
        y, m = map(int, ymv.split('-')); k = kidx(y, m)
        if 0 <= k < NM: byk[k] = max(byk.get(k, 0), int(ms or 0))
    for k in range(NM):
        if k in byk: cur = max(cur or 0, byk[k])
        out[k] = cur
    return out
