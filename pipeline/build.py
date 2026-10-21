"""
MoonTech Monthly Model dashboard - production build.
Level 0 + Level 1 come from Kareem's reference engine (moontech_ref.py, unchanged).
This file adds: ROAS S/M/L tiers, Ounass buffer scenarios, per-page datasets -> data.json.

Run:  python pipeline/build.py            (from the repo root)
"""
import os, sys, json, csv, itertools
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import moontech_ref as R
from moontech_ref import Engine, ACT, kidx, mo, NM
import openpyxl

NF = 32                       # full months with Datos + P&L: Jan-24 .. Aug-26
LAUNCH = kidx(2025, 11)       # first ROAS signing
LEG_SEGS = ['Enterprise', 'Growth', 'SME']
TIERS = ['Small', 'Medium', 'Large']
LVL_GROUPS = ['L1', 'L2+']
PL_OPEX = ['bd_salaries', 'digital_marketing', 'marketing_events', 'account_mgmt', 'c_suite', 'ops_salaries',
           'advisory', 'system_expenses', 'ga']            # development handled separately
OPEX_GROUPS = {'marketing': ['bd_salaries', 'digital_marketing', 'marketing_events', 'account_mgmt'],
               'staff': ['c_suite', 'ops_salaries', 'advisory'], 'technology': ['system_expenses'], 'ga': ['ga']}

def r2(x, n=2):
    if x is None: return None
    if isinstance(x, float) and (x != x): return None
    return round(x, n)

# Simpler approach: subclass load so the add-back happens inside Engine.__init__.
_orig_load = R.load_merchant_revenue
def _patched_load(kit, master, lad, cfg, source='system'):
    rev, adj, review = _orig_load(kit, master, lad, cfg, source)
    ctl = _patched_load.control; ab = _patched_load.addback
    if ab:
        m = ctl['ounass']['merchant']
        for k in ab: rev[m][k] += ctl['ounass']['buffer_aed']
    return rev, adj, review
_patched_load.control = None; _patched_load.addback = set()
R.load_merchant_revenue = _patched_load
R.load_config = lambda kit: __import__('yaml').safe_load(open(os.path.join(HERE, 'config.yaml')))
_orig_master = R.load_master
R.load_master = lambda kit: _orig_master(HERE)

# ----------------------------------------------------------------------------- helpers on an Engine
def month_keys(): return [mo(k) for k in range(NF)]

def load_gmv_ads(kit):
    wb = openpyxl.load_workbook(os.path.join(kit, 'data', 'MoonTech_Monthly_Model_6.xlsx'), data_only=True, read_only=True)
    ws = wb['Monthly Actuals']
    rows = {r: next(ws.iter_rows(min_row=r, max_row=r, max_col=37, values_only=True)) for r in (16, 20)}
    ads = [rows[16][5 + k] for k in range(NF)]; gmv = [rows[20][5 + k] for k in range(NF)]
    f = lambda v: float(v) if isinstance(v, (int, float)) else None
    return [f(v) for v in gmv], [f(v) for v in ads]

def first_true(seq):
    return next((i for i, v in enumerate(seq) if v), None)

def cohort_window_status(status_seq, k, win=12):
    """Revenue-split label at month k under the 12-month cohort window."""
    s = status_seq[k]
    if s not in ACT: return 'Not active'
    # find start of the current active run
    start = k
    while start > 0 and status_seq[start - 1] in ACT: start -= 1
    kind = status_seq[start]        # 'New' or 'Reactivated'
    if start == 0: return 'Existing'   # pre-2024 base: true start unknown
    if kind in ('New', 'Reactivated') and k - start < win: return kind
    return 'Existing'

def twelve_month(mc): return None if mc is None else 1 - (1 - mc) ** 12
def lifetime(mc, cap): return None if mc is None else (cap if mc <= 0 else min(cap, 1 / mc))

# ----------------------------------------------------------------------------- dataset
def build_dataset(E, ctl):
    fx = E.fx; cfg = E.cfg; cap = cfg['ltv']['lifetime_cap_months']
    gmv, ads = load_gmv_ads(E.kit)
    L = NF - 1                                       # latest full month index (Aug-26)
    months = month_keys()
    leg_m = list(E.rev)                              # merchants
    brands = [b for b, v in E.lad.items() if not v['other_pricing']]
    other_b = [b for b, v in E.lad.items() if v['other_pricing']]
    lvl = {b: ('L1' if E.lad[b]['level'] == 'L1' else 'L2+') for b in brands}

    # ---- merchant CAC since launch at each month (for the Small/Medium line)
    spend_to = []; new_to = []
    for k in range(NM):
        s = sum((E.cac_m[j] or 0) for j in range(LAUNCH, k + 1)) if k >= LAUNCH else 0
        n = sum(1 for b in brands if LAUNCH <= E.lad[b]['since'] <= k)
        spend_to.append(s); new_to.append(n)
    cac_sl = [(spend_to[k] / new_to[k]) if (k >= LAUNCH and new_to[k] and all(E.cac_m[j] is not None for j in range(LAUNCH, k + 1))) else None for k in range(NM)]
    # TTM CAC after 12 months of history
    def cac_at(k):
        v = _cac_at(k)
        if v is None:
            for j in range(k - 1, LAUNCH - 1, -1):
                v = _cac_at(j)
                if v is not None: break
        return v
    def _cac_at(k):
        if k - LAUNCH + 1 >= 12:
            ks = range(k - 11, k + 1)
            if any(E.cac_m[j] is None for j in ks): return None
            n = sum(1 for b in brands if k - 11 <= E.lad[b]['since'] <= k)
            return sum(E.cac_m[j] for j in ks) / n if n else None
        return cac_sl[k]

    # ---- legacy benchmark for Large: ARPA of the legacy segment with most active legacy accounts
    bench = [None] * NM; bench_seg = [None] * NM
    for k in range(2, NF):
        counts = {s: len(E.legacy_active(k, [s])) for s in LEG_SEGS}
        seg = max(counts, key=counts.get)
        try: bench[k] = E.seg_arpa(k, [seg])[0]; bench_seg[k] = seg
        except ZeroDivisionError: pass
    for k in range(NF, NM): bench[k] = bench[NF - 1]; bench_seg[k] = bench_seg[NF - 1]

    # ---- ROAS tiers per brand per month
    margin_r = cfg['ltv']['roas_margin']
    tier = {}
    for b in brands:
        v = E.lad[b]; t = ['-'] * NM; prev = None
        for k in range(v['since'], NM):
            if E.rst[b][k] not in ACT: t[k] = 'Churned' if E.rst[b][k] == 'Churned' else '-'; prev = None; continue
            paid = sum(v['pay'][v['since']:k + 1]); ten = k - v['since'] + 1
            ann = paid * 12 / ten if ten < 12 else sum(v['pay'][k - 11:k + 1])
            c = cac_at(k)
            if bench[k] and ann >= bench[k]: tt = 'Large'
            elif c is not None and paid * margin_r >= c: tt = 'Medium'
            else: tt = 'Small'
            t[k] = tt
        tier[b] = t
    tier_moves = {k: [] for k in range(NM)}
    for b in brands:
        for k in range(1, NM):
            a, c = tier[b][k - 1], tier[b][k]
            if a in TIERS and c in TIERS and a != c: tier_moves[k].append({'brand': b, 'from': a, 'to': c})

    # ---- per-month series
    def leg_rev_usd(m, k):          # named legacy revenue (Datos before Since)
        return E.rev[m][k] / fx if (m not in E.since or k < E.since[m]) else 0.0
    def roas_group(k, g):           # brands in a group at k
        if g == 'ROAS': return [b for b in brands if E.rst[b][k] in ACT]
        if g in LVL_GROUPS: return [b for b in brands if E.rst[b][k] in ACT and lvl[b] == g]
        return [b for b in brands if E.rst[b][k] in ACT and tier[b][k] == g]
    def roas_group_members(g):      # static membership for churn windows
        if g == 'ROAS': return brands
        if g in LVL_GROUPS: return [b for b in brands if lvl[b] == g]
        return None

    S = {}                            # monthly series
    S['months'] = months
    S['incomplete'] = [False] * NF
    S['rev_total'] = E.total[:NF]; S['rev_named'] = E.named[:NF]; S['rev_roas'] = E.roas[:NF]
    S['rev_other'] = E.other[:NF]; S['rev_plug'] = E.plug[:NF]; S['rev_legacy'] = E.legacy[:NF]
    S['rev_roas_ext'] = E.roas[:NF + 1]        # Sep-26 ROAS (incomplete month)
    S['rev_leg_seg'] = {s: [sum(leg_rev_usd(m, k) for m in leg_m if E.seg[m][k] == s) for k in range(NF)] for s in LEG_SEGS + ['Unassigned']}
    S['rev_roas_tier'] = {t: [sum(E.lad[b]['pay'][k] for b in brands if tier[b][k] == t) for k in range(NF)] for t in TIERS}
    S['rev_roas_lvl'] = {g: [sum(E.lad[b]['pay'][k] for b in brands if lvl[b] == g) for k in range(NF)] for g in LVL_GROUPS}
    st_lab = ['Existing', 'New', 'Reactivated', 'Not active']
    S['rev_leg_status'] = {s: [0.0] * NF for s in st_lab}
    for m in leg_m:
        for k in range(NF):
            v = leg_rev_usd(m, k)
            if v: S['rev_leg_status'][cohort_window_status(E.leg[m], k)][k] += v
    # counts
    S['act_leg'] = [len(E.legacy_active(k)) for k in range(NF)]
    S['act_leg_seg'] = {s: [len(E.legacy_active(k, [s])) for k in range(NF)] for s in LEG_SEGS}
    S['act_roas'] = [len(roas_group(k, 'ROAS')) for k in range(NM)]
    S['act_roas_tier'] = {t: [len(roas_group(k, t)) for k in range(NM)] for t in TIERS}
    S['act_roas_lvl'] = {g: [len(roas_group(k, g)) for k in range(NM)] for g in LVL_GROUPS}
    S['act_other'] = [sum(1 for b in other_b if E.rst[b][k] in ACT) for k in range(NM)]
    S['act_total'] = [S['act_leg'][k] + S['act_roas'][k] + S['act_other'][k] for k in range(NF)]
    # movement
    mv = {x: [0] * NF for x in ['leg_new', 'leg_react', 'leg_churn', 'leg_migrated']}
    for m in leg_m:
        for k in range(NF):
            s = E.leg[m][k]
            if s == 'New': mv['leg_new'][k] += 1
            elif s == 'Reactivated': mv['leg_react'][k] += 1
            elif s == 'Churned': mv['leg_churn'][k] += 1
            if m in E.since and E.since[m] == k and k > 0 and E.leg[m][k - 1] in ACT: mv['leg_migrated'][k] += 1
    for x in ['roas_signed', 'roas_migrated', 'roas_churn', 'roas_react']: mv[x] = [0] * NM
    migrated_brands = {b for m, b in E.link.items() if E.since[m] > 0 and E.leg[m][E.since[m] - 1] in ACT}
    for b in brands:
        for k in range(NM):
            s = E.rst[b][k]
            if s == 'New': mv['roas_signed'][k] += 1; mv['roas_migrated'][k] += (b in migrated_brands)
            elif s == 'Churned': mv['roas_churn'][k] += 1
            elif s == 'Reactivated': mv['roas_react'][k] += 1
    S['mv'] = mv
    S['tier_moves'] = {mo(k): tier_moves[k] for k in range(NF + 1) if tier_moves[k]}
    S['at_risk_leg'] = []; S['at_risk_roas'] = []
    for k in range(NF):
        Q = cfg['merchant_status']['quiet_month_threshold_usd'] * fx; lo, hi = cfg['churn']['at_risk_quiet_months']
        ar = []
        for m in E.legacy_active(k):
            q = 0
            for j in range(k, -1, -1):
                if E.rev[m][j] < Q: q += 1
                else: break
            if lo <= q <= hi: ar.append({'name': m, 'months': q})
        S['at_risk_leg'].append(ar)
    for k in range(NM):
        lo, hi = cfg['churn']['at_risk_quiet_months']; ar = []
        for b in roas_group(k, 'ROAS'):
            q = 0
            for j in range(k, E.lad[b]['since'] - 1, -1):
                if E.lad[b]['pay'][j] <= 0: q += 1
                else: break
            if lo <= q <= hi: ar.append({'name': b, 'months': q})
        S['at_risk_roas'].append(ar)
    # ARR + ARPA
    arr_r = [E.roas_arr(k) if k >= 2 else None for k in range(NM)]
    arr_l = [E.legacy_runrate(k) if k >= 2 else None for k in range(NF)]
    S['arr_roas'] = arr_r; S['arr_legacy'] = arr_l
    S['arr_blended'] = [(arr_r[k] + arr_l[k]) if k >= 2 else None for k in range(NF)]
    def leg_named_arr(k, segs):
        Rr = N = 0.0
        for j in range(k - 2, k + 1):
            act = E.legacy_active(j, segs)
            Rr += sum(E.rev[m][j] for m in act) / fx / E.w[j % 12] / 3; N += len(act) / 3
        return Rr, N
    def roas_grp_arr(k, g):
        Rr = N = 0.0
        for j in range(k - 2, k + 1):
            mem = roas_group(j, g); N += len(mem) / 3
            Rr += sum(E.lad[b]['pay'][j] for b in mem) * cfg['arr']['roas_multiplier']
        return Rr, N
    arpa = {g: [None] * NM for g in ['blended', 'blended_ex_ent', 'Enterprise', 'Growth', 'SME', 'ROAS'] + TIERS + LVL_GROUPS}
    arpa_n = {g: [None] * NM for g in arpa}
    for k in range(2, NM):
        for g in ['ROAS'] + TIERS + LVL_GROUPS:
            rr, n = roas_grp_arr(k, g); arpa[g][k] = rr / n if n else None; arpa_n[g][k] = n
        if k < NF:
            tot_r = tot_n = 0.0
            for s in LEG_SEGS:
                rr, n = leg_named_arr(k, [s]); arpa[s][k] = rr / n if n else None; arpa_n[s][k] = n
                tot_r += rr; tot_n += n
            rr, n = roas_grp_arr(k, 'ROAS')
            arpa['blended'][k] = (tot_r + rr) / (tot_n + n) if tot_n + n else None; arpa_n['blended'][k] = tot_n + n
            er, en = leg_named_arr(k, ['Enterprise'])
            d = tot_n + n - en
            arpa['blended_ex_ent'][k] = (tot_r + rr - er) / d if d else None; arpa_n['blended_ex_ent'][k] = d
    S['arpa'] = arpa; S['arpa_n'] = arpa_n
    S['benchmark_large'] = bench; S['benchmark_seg'] = bench_seg; S['cac_sl'] = cac_sl
    S['ent_share'] = [(S['rev_leg_seg']['Enterprise'][k] / S['rev_named'][k]) if S['rev_named'][k] else None for k in range(NF)]
    # P&L (USD)
    P = {}
    for l in ['sales', 'creator_payouts', 'materials', 'freelancers', 'other_cos', 'development', 'da', 'finance_cost', 'tax', 'transfer_pricing']:
        P[l] = [E.pl[l][k] / fx for k in range(NF)]
    P['sales'] = list(E.total[:NF])          # P&L Sales where loaded; Datos total + adjustment for 2024-25
    for g, ls in OPEX_GROUPS.items(): P['opex_' + g] = [sum(E.pl[l][k] for l in ls) / fx for k in range(NF)]
    P['opex_ex_dev'] = [sum(P['opex_' + g][k] for g in OPEX_GROUPS) for k in range(NF)]
    P['cos_total'] = [P['creator_payouts'][k] + P['materials'][k] + P['freelancers'][k] + P['other_cos'][k] for k in range(NF)]
    S['pl'] = P
    S['gmv'] = gmv; S['ads'] = ads
    S['cac_merchant_spend'] = E.cac_m[:NF]; S['cac_creator_spend'] = E.cac_c[:NF]
    # creators
    cr = {x: [0] * NF for x in ['active', 'emerging', 'new', 'churned', 'react', 'major', 'orders_total', 'orders_active', 'orders_unatt', 'orders_major']}
    first_order = {i: first_true([v >= 1 for v in x]) for i, x in E.orders.items()}
    cr['recruited'] = [0] * NF; cr['first_active'] = [0] * NF
    first_active = {}
    for i, x in E.orders.items():
        st = E.cst[i]
        fa = first_true([s in ACT for s in st]); first_active[i] = fa
        if fa is not None: cr['first_active'][fa] += 1
        if first_order[i] is not None: cr['recruited'][first_order[i]] += 1
        for k in range(NF):
            s = st[k]
            if s in ACT:
                cr['active'][k] += 1; cr['orders_active'][k] += x[k]
                if E.major[i][k]: cr['major'][k] += 1; cr['orders_major'][k] += x[k]
            if s == 'Emerging': cr['emerging'][k] += 1
            if s == 'New': cr['new'][k] += 1
            if s == 'Reactivated': cr['react'][k] += 1
            if s == 'Churned': cr['churned'][k] += 1
            cr['orders_total'][k] += x[k]
    for k in range(NF): cr['orders_unatt'][k] = E.unatt[k]; cr['orders_total'][k] += E.unatt[k]
    S['creators'] = cr

    # ---- period metrics
    def nrr(series, p0, p1):
        s = e = keep = 0.0; n = 0; lost = 0.0
        for v in series.values():
            a = sum(v[j] for j in p0); b = sum(v[j] for j in p1)
            if a > 0:
                n += 1; s += a; e += b; keep += min(a, b)
                if b <= 0: lost += a
        if not s: return None
        return {'nrr': e / s, 'grr': keep / s, 'lost': lost / s, 'contraction': 1 - keep / s - lost / s, 'expansion': e / s - keep / s, 'cohort': n}
    def leg_series(segs=None, at=None):
        return {m: [leg_rev_usd(m, k) for k in range(NF)] for m in leg_m if not segs or E.seg[m][at] in segs}
    def roas_series(mem=None):
        return {b: E.lad[b]['pay'] for b in (mem or brands)}
    def churn_block(events, avg_active, nmonths):
        if not avg_active or not nmonths: return None
        mc = events / avg_active / nmonths
        return {'events': events, 'avg_active': r2(avg_active, 1), 'monthly': mc, 'period_rate': events / avg_active,
                'twelve_month': twelve_month(mc), 'lifetime': lifetime(mc, cap), 'small_sample': events < cfg['churn']['small_sample_flag_below_events']}
    def leg_churn(ks, segs=None):
        ev = len({m for m in leg_m for k in ks if E.leg[m][k] == 'Churned' and (not segs or E.seg[m][k] in segs)})
        avg = sum(len(E.legacy_active(k, segs)) for k in ks) / len(ks)
        return churn_block(ev, avg, len(ks))
    def roas_churn(ks, g='ROAS'):
        ks = [k for k in ks if k >= LAUNCH]
        if not ks: return None
        if g in TIERS:
            ev = len({b for b in brands for k in ks if E.rst[b][k] == 'Churned' and tier[b][k - 1] == g})
        else:
            mem = roas_group_members(g)
            ev = len({b for b in mem for k in ks if E.rst[b][k] == 'Churned'})
        avg = sum(len(roas_group(k, g)) for k in ks) / len(ks)
        return churn_block(ev, avg, len(ks))
    def conc(ks):
        acc = {}
        for m in leg_m:
            acc[m] = sum(leg_rev_usd(m, k) for k in ks)
            if m in E.link: acc[m] += sum(E.lad[E.link[m]]['pay'][k] for k in ks if k >= E.since[m])
        for b in E.lad:
            if b not in E.link.values(): acc[b] = sum(E.lad[b]['pay'][k] for k in ks)
        T = sum(E.total[k] for k in ks)
        top = sorted(acc.items(), key=lambda x: -x[1])
        if not T: return None
        return {'top1': top[0][1] / T, 'top1_name': top[0][0], 'top5': sum(v for _, v in top[:5]) / T, 'top10': sum(v for _, v in top[:10]) / T,
                'hhi': sum((v / T * 100) ** 2 for _, v in top), 'top': [{'name': n, 'share': v / T, 'rev': v} for n, v in top[:10]]}
    margin_ttm_cache = {}
    def legacy_margin(k):
        if k in margin_ttm_cache: return margin_ttm_cache[k]
        ks = range(max(0, k - 11), k + 1)
        rev = sum(E.total[j] for j in ks); cp = sum(P['creator_payouts'][j] for j in ks)
        margin_ttm_cache[k] = 1 - cp / rev if rev else None
        return margin_ttm_cache[k]

    def period(ks, label, kind):
        ks = list(ks); k1 = ks[-1]; n = len(ks)
        out = {'label': label, 'kind': kind, 'start': mo(ks[0]), 'end': mo(k1), 'months': n}
        sm = lambda arr: sum((arr[k] or 0) for k in ks)
        out['rev'] = {'total': sm(E.total), 'named': sm(E.named), 'roas': sm(E.roas), 'other': sm(E.other), 'plug': sm(E.plug), 'legacy': sm(E.legacy),
                      'leg_seg': {s: sm(S['rev_leg_seg'][s]) for s in LEG_SEGS + ['Unassigned']},
                      'roas_tier': {t: sm(S['rev_roas_tier'][t]) for t in TIERS}, 'roas_lvl': {g: sm(S['rev_roas_lvl'][g]) for g in LVL_GROUPS},
                      'leg_status': {s: sm(S['rev_leg_status'][s]) for s in st_lab}}
        # growth YoY (complete prior-year window only)
        if ks[0] - 12 >= 0:
            pk = [k - 12 for k in ks]
            py = lambda arr: sum((arr[k] or 0) for k in pk)
            out['yoy'] = {'total': sm(E.total) / py(E.total) - 1 if py(E.total) else None,
                          'legacy': sm(E.legacy) / py(E.legacy) - 1 if py(E.legacy) else None,
                          'roas': sm(E.roas) / py(E.roas) - 1 if py(E.roas) else None}
            top = ctl['ounass']['merchant']
            ex = lambda arr_ks: sum(E.total[k] - E.rev[top][k] / fx for k in arr_ks)
            out['yoy']['total_ex_top'] = ex(ks) / ex(pk) - 1 if ex(pk) else None
            out['yoy']['gmv'] = sm(gmv) / py(gmv) - 1 if py(gmv) else None
            out['yoy']['orders'] = sm(cr['orders_total']) / py(cr['orders_total']) - 1 if py(cr['orders_total']) else None
        # counts
        out['active_end'] = {'legacy': S['act_leg'][k1], 'leg_seg': {s: S['act_leg_seg'][s][k1] for s in LEG_SEGS}, 'roas': S['act_roas'][k1],
                             'roas_tier': {t: S['act_roas_tier'][t][k1] for t in TIERS}, 'roas_lvl': {g: S['act_roas_lvl'][g][k1] for g in LVL_GROUPS},
                             'other': S['act_other'][k1], 'total': S['act_total'][k1]}
        out['active_avg'] = {'legacy': sum(S['act_leg'][k] for k in ks) / n, 'roas': sum(S['act_roas'][k] for k in ks) / n}
        out['events'] = {x: sum(mv[x][k] for k in ks) for x in mv}
        # churn
        out['churn'] = {'legacy': leg_churn(ks), 'Enterprise': leg_churn(ks, ['Enterprise']), 'GrowthSME': leg_churn(ks, ['Growth', 'SME', 'Unassigned']),
                        'ROAS': roas_churn(ks), 'L1': roas_churn(ks, 'L1'), 'L2+': roas_churn(ks, 'L2+')}
        for t in TIERS: out['churn'][t] = roas_churn(ks, t)
        # retention
        ret = {}
        if n == 12 and ks[0] - 12 >= 0:
            p0 = range(ks[0] - 12, ks[0])
            ret['legacy_12m'] = nrr(leg_series(), p0, ks)
            ret['ent_12m'] = nrr(leg_series(['Enterprise'], ks[0] - 1), p0, ks)
            ret['growthsme_12m'] = nrr(leg_series(['Growth', 'SME', 'Unassigned'], ks[0] - 1), p0, ks)
            ret['orders_12m'] = nrr({i: x for i, x in E.orders.items()}, p0, ks)
        if kind == 'quarter' and ks[0] - 3 >= LAUNCH and ks[0] + 2 <= NF:
            fq = list(range(ks[0], ks[0] + 3))
            ret['roas_qoq'] = nrr(roas_series(), range(ks[0] - 3, ks[0]), fq)
            ret['roas_qoq']['partial_legacy'] = len(ks) < 3
        if n == 3 and ks[0] - 12 >= 0:
            ret['legacy_3m_yoy'] = nrr(leg_series(), [k - 12 for k in ks], ks)
        out['retention'] = ret
        # CAC
        cks = [k for k in ks if k >= LAUNCH]
        if cks:
            complete = all(E.cac_m[k] is not None for k in cks)
            spend = sum((E.cac_m[k] or 0) for k in cks); newb = sum(1 for b in brands if cks[0] <= E.lad[b]['since'] <= cks[-1])
            out['cac'] = {'spend': spend, 'new_brands': newb, 'cac': spend / newb if (newb and complete) else None, 'complete': complete}
        else: out['cac'] = None
        ccomplete = all(E.cac_c[k] is not None for k in ks)
        cspend = sum((E.cac_c[k] or 0) for k in ks); rec = sum(cr['recruited'][k] for k in ks); fa = sum(cr['first_active'][k] for k in ks)
        out['creator_cac'] = {'spend': cspend, 'recruited': rec, 'new_active': fa, 'per_recruited': cspend / rec if (rec and ccomplete) else None,
                              'per_active': cspend / fa if (fa and ccomplete) else None, 'complete': ccomplete}
        # ARR / ARPA at end
        out['arr'] = {'roas': arr_r[k1], 'legacy': arr_l[k1] if k1 < NF else None, 'blended': S['arr_blended'][k1] if k1 < NF else None}
        out['arpa'] = {g: arpa[g][k1] for g in arpa}; out['arpa_n'] = {g: arpa_n[g][k1] for g in arpa}
        # LTV
        ltv = {}
        cac_v = out['cac']['cac'] if out['cac'] else None
        for g in ['ROAS', 'L1', 'L2+'] + TIERS:
            ch = out['churn'].get(g); a = arpa[g][k1]
            if ch and a:
                lt = ch['lifetime']; v = a / 12 * margin_r * lt
                ltv[g] = {'arpa': a, 'margin': margin_r, 'churn_basis': 'logo', 'monthly_churn': ch['monthly'], 'lifetime': lt, 'ltv': v,
                          'ltv_cac': v / cac_v if cac_v else None, 'payback': cac_v / (a / 12 * margin_r) if cac_v else None,
                          'small_sample': ch['small_sample'], 'survivorship': g == 'L2+'}
        lm = legacy_margin(k1)
        for g, key, segs in [('Enterprise', 'ent_12m', ['Enterprise']), ('GrowthSME', 'growthsme_12m', ['Growth', 'SME', 'Unassigned'])]:
            rr = ret.get(key); a = arpa['Enterprise'][k1] if g == 'Enterprise' else None
            if g == 'GrowthSME' and k1 >= 2:
                r_, n_ = leg_named_arr(k1, ['Growth', 'SME', 'Unassigned']); a = r_ / n_ if n_ else None
            if rr and a and lm:
                mc = (1 - rr['grr']) / 12; lt = lifetime(mc, cap)
                ltv[g] = {'arpa': a, 'margin': lm, 'churn_basis': 'revenue (1 - GRR)', 'monthly_churn': mc, 'lifetime': lt, 'ltv': a / 12 * lm * lt}
        # company-level LTV (weighted by avg active accounts)
        parts = []
        if 'Enterprise' in ltv: parts.append((ltv['Enterprise']['ltv'], arpa_n['Enterprise'][k1] or 0))
        if 'GrowthSME' in ltv: parts.append((ltv['GrowthSME']['ltv'], (arpa_n['Growth'][k1] or 0) + (arpa_n['SME'][k1] or 0)))
        if 'ROAS' in ltv: parts.append((ltv['ROAS']['ltv'], arpa_n['ROAS'][k1] or 0))
        if parts and sum(w for _, w in parts): ltv['company'] = {'ltv': sum(v * w for v, w in parts) / sum(w for _, w in parts), 'churn_basis': 'weighted'}
        out['ltv'] = ltv; out['legacy_margin_ttm'] = lm
        out['concentration'] = conc(ks)
        # P&L sums
        pl = {l: sm(P[l]) for l in P}
        out['pl'] = pl
        out['gmv'] = sm(gmv) if all(gmv[k] is not None for k in ks) else None
        out['ads'] = sm(ads) if all(ads[k] is not None for k in ks) else None
        out['creators'] = {'active_end': cr['active'][k1], 'emerging_end': cr['emerging'][k1], 'major_end': cr['major'][k1],
                           'new': sum(cr['new'][k] for k in ks), 'churned': sum(cr['churned'][k] for k in ks), 'react': sum(cr['react'][k] for k in ks),
                           'avg_active': sum(cr['active'][k] for k in ks) / n, 'orders_total': sm(cr['orders_total']), 'orders_active': sm(cr['orders_active']),
                           'orders_major': sm(cr['orders_major']), 'orders_unatt': sm(cr['orders_unatt']),
                           'recruited': rec, 'first_active': fa}
        oc = out['creators']
        oc['orders_per_active'] = oc['orders_active'] / oc['avg_active'] if oc['avg_active'] else None
        oc['churn'] = churn_block(oc['churned'], oc['avg_active'], n)
        # graduation (cohort of first orders in period -> active within 3 months)
        coh = [i for i, f in first_order.items() if f is not None and f in ks and f >= 3]
        grad3 = sum(1 for i in coh if first_active[i] is not None and first_active[i] - first_order[i] <= 3)
        oc['graduation'] = {'cohort': len(coh), 'within3': grad3 / len(coh) if coh else None}
        out['ounass_flag'] = any(mo(k) in ctl['ounass']['provisional_months'] for k in ks)
        return out

    periods = {}
    for k in range(NF): periods['M:' + mo(k)] = period([k], mo(k), 'month')
    for y in (2024, 2025, 2026):
        for q in range(1, 5):
            ks = [kidx(y, 3 * (q - 1) + 1 + i) for i in range(3) if kidx(y, 3 * (q - 1) + 1 + i) < NF]
            if ks:
                p = period(ks, 'Q%d-%s' % (q, str(y)[2:]), 'quarter'); p['complete'] = len(ks) == 3
                periods['Q:%d-Q%d' % (y, q)] = p
    periods['FY:2024'] = period(range(0, 12), 'FY2024', 'year')
    periods['FY:2025'] = period(range(12, 24), 'FY2025', 'year')
    periods['YTD:2026'] = period(range(24, NF), '2026 YTD', 'ytd')
    periods['YTD:2025'] = period(range(12, 12 + NF - 24), '2025 YTD (Jan-Aug)', 'ytd')
    for k in range(11, NF): periods['TTM:' + mo(k)] = period(range(k - 11, k + 1), 'TTM ' + mo(k), 'ttm')
    periods['SL'] = period(range(LAUNCH, NF), 'Since launch (Nov-25)', 'since_launch')

    # ---- annualised view (10.19)
    elapsed = NF - 24; share = sum(E.w[:elapsed])
    ytd = periods['YTD:2026']
    leg_ann = ytd['rev']['legacy'] / share
    roas_ann = ytd['rev']['roas'] + ytd['rev']['other'] + (arr_r[L] or 0) * (12 - elapsed) / 12
    rev_ann = leg_ann + roas_ann
    pct = lambda l: ytd['pl'][l] / ytd['rev']['total'] if ytd['rev']['total'] else 0
    ann = {'legacy': leg_ann, 'roas': roas_ann, 'total': rev_ann, 'share_elapsed': share, 'months_elapsed': elapsed,
           'creator_payouts': rev_ann * pct('creator_payouts'), 'materials': rev_ann * pct('materials'), 'freelancers': rev_ann * pct('freelancers'),
           'other_cos': rev_ann * pct('other_cos'), 'transfer_pricing': rev_ann * pct('transfer_pricing'),
           'opex_ex_dev': ytd['pl']['opex_ex_dev'] * 12 / elapsed, 'development': ytd['pl']['development'] * 12 / elapsed,
           'da': ytd['pl']['da'] * 12 / elapsed, 'finance_cost': ytd['pl']['finance_cost'] * 12 / elapsed, 'tax': ytd['pl']['tax'] * 12 / elapsed}
    for g in OPEX_GROUPS: ann['opex_' + g] = ytd['pl']['opex_' + g] * 12 / elapsed
    # orders / gmv annualised by own seasonal share (history 2024-25)
    def own_share(arr):
        t = sum((arr[k] or 0) for k in range(24));
        return sum(((arr[i] or 0) + (arr[12 + i] or 0)) / t for i in range(elapsed)) if t else None
    os_ = own_share(cr['orders_total']); gs_ = own_share(gmv)
    ann['orders'] = ytd['creators']['orders_total'] / os_ if os_ else None
    ann['gmv'] = ytd['gmv'] / gs_ if (gs_ and ytd['gmv']) else None
    ann['ads'] = (ytd['ads'] or 0) * 12 / elapsed if ytd['ads'] else None
    ann['acq_spend'] = (ytd['cac']['spend'] if ytd['cac'] else 0) * 12 / elapsed

    # ---- cohorts (10.7) ROAS logo retention by signing quarter
    def qlabel(k): return 'Q%d-%s' % ((k % 12) // 3 + 1, str(2024 + k // 12)[2:])
    coh = {}
    for b in brands:
        s = E.lad[b]['since']; q = qlabel(s); coh.setdefault(q, {'brands': [], 'since': []})
        coh[q]['brands'].append(b); coh[q]['since'].append(s)
    roas_cohorts = []
    for q in sorted(coh, key=lambda x: min(coh[x]['since'])):
        bs = coh[q]['brands']; cells = []
        for n_ in range(0, 12):
            vals = []
            for b in bs:
                k = E.lad[b]['since'] + n_
                if k <= L + 1: vals.append(1 if E.rst[b][k] in ACT else 0)
            if len(vals) == len(bs): cells.append(sum(vals) / len(bs))
            else: break
        roas_cohorts.append({'cohort': q, 'size': len(bs), 'cells': cells})
    leg_cohorts = []; lc = {}
    for m in leg_m:
        f = first_true([s == 'New' for s in E.leg[m]])
        if f is None: continue
        key = 'Pre-2024 base' if f == 0 else qlabel(f)
        lc.setdefault(key, []).append((m, f))
    for key in sorted(lc, key=lambda x: (x != 'Pre-2024 base', min(f for _, f in lc[x]))):
        ms = lc[key]; cells = []
        for n_ in range(0, 25):
            vals = [1 if E.leg[m][f + n_] in ACT else 0 for m, f in ms if f + n_ < NF]
            if len(vals) == len(ms): cells.append(sum(vals) / len(ms))
            else: break
        leg_cohorts.append({'cohort': key, 'size': len(ms), 'cells': cells})

    # ---- entity tables
    merchants = []
    for m in leg_m:
        last_act = max([k for k in range(NF) if E.leg[m][k] in ACT], default=None)
        merchants.append({'name': m, 'segment': E.seg[m][L], 'status': E.leg[m][L] if E.leg[m][L] != '-' else ('Migrated to ROAS' if m in E.since and E.since[m] <= L else 'Inactive'),
                          'roas_link': E.link.get(m), 'since': mo(E.since[m]) if m in E.since else None,
                          'rev_ttm': sum(leg_rev_usd(m, k) for k in range(L - 11, L + 1)), 'rev_2026': sum(leg_rev_usd(m, k) for k in range(24, NF)),
                          'rev_2025': sum(leg_rev_usd(m, k) for k in range(12, 24)), 'rev_2024': sum(leg_rev_usd(m, k) for k in range(0, 12)),
                          'last_active': mo(last_act) if last_act is not None else None, 'master_status': E.master.get(m, {}).get('status'),
                          'monthly': [r2(leg_rev_usd(m, k), 0) for k in range(NF)], 'seg_monthly': E.seg[m][:NF]})
    roas_tbl = []
    for b in brands + other_b:
        v = E.lad[b]
        roas_tbl.append({'name': b, 'since': mo(v['since']), 'level': v['level'], 'lvl_group': lvl.get(b, 'Other'), 'other_pricing': v['other_pricing'],
                         'status': E.rst[b][L + 1], 'churn_month': mo(v['churn']) if v['churn'] is not None else None,
                         'paid_total': sum(v['pay']), 'paid_to_latest': sum(v['pay'][:L + 2]), 'tier': tier[b][L + 1] if b in tier else None,
                         'tier_monthly': tier[b][:NM] if b in tier else None, 'migrated': b in migrated_brands,
                         'monthly': v['pay'][:NM], 'status_monthly': E.rst[b][:NM]})
    return {'series': S, 'periods': periods, 'annualised': ann, 'cohorts': {'roas': roas_cohorts, 'legacy': leg_cohorts},
            'merchants': merchants, 'roas_brands': roas_tbl, 'weights': E.w,
            'review': {'new_names': E.review, 'creator_id_flags': E.id_flags, 'r2_churned': [b for b in brands if E.lad[b]['churn'] is None and any(E.rst[b][k] == 'Churned' for k in range(NM))],
                       'low_confidence': [m for m, v in E.master.items() if v.get('status') != 'Confirmed']}}

# ----------------------------------------------------------------------------- scenarios + checks
def run_checks(E, kit):
    res = E.compute_checks(); out = []
    for c in csv.DictReader(open(os.path.join(HERE, 'check_figures.csv'))):
        got = res.get(c['id']); exp = float(c['expected_spec_ingestion'])
        out.append({'id': c['id'], 'metric': c['metric'], 'period': c['period'], 'expected': exp, 'got': r2(got, 3), 'pass': got is not None and abs(got - exp) <= float(c['tolerance']) + 1e-9})
    return out

def make_engine(kit, ctl, addback):
    _patched_load.control = ctl; _patched_load.addback = set(addback)
    E = Engine(kit, 'system'); E.kit = kit
    return E

def main():
    kit = ROOT
    # Engine expects kit/data/*, kit/config.yaml, kit/merchant_master_seed.csv
    ctl = json.load(open(os.path.join(HERE, 'control.json')))
    prov = ctl['ounass']['provisional_months']
    prov_k = [kidx(int(m[:4]), int(m[5:])) for m in prov]
    scen = {}
    base_E = None
    for r in range(len(prov_k) + 1):
        for combo in itertools.combinations(sorted(prov_k), r):
            key = '+'.join(mo(k) for k in combo) or 'none'
            E = make_engine(kit, ctl, combo)
            if r == 0: base_E = E
            scen[key] = build_dataset(E, ctl)
            print('scenario', key, 'built', file=sys.stderr)
    checks = run_checks(base_E, kit)
    npass = sum(1 for c in checks if c['pass'])
    print('check figures: %d/%d pass' % (npass, len(checks)), file=sys.stderr)
    base = scen['none']
    # store overrides only where a section differs
    out = {'meta': {'built': __import__('datetime').date.today().isoformat(), 'latest_month': mo(NF - 1), 'fx': base_E.fx, 'spec_version': 'v0.25',
                    'config': base_E.cfg, 'control': ctl, 'checks': checks, 'checks_pass': npass, 'ladder_file': 'moontech-roas-ladder-2026-10-05.csv'},
           'base': base, 'scenarios': {}}
    for key, ds in scen.items():
        if key == 'none': continue
        diff = {}
        for sec in ds:
            if json.dumps(ds[sec], sort_keys=True, default=str) != json.dumps(base[sec], sort_keys=True, default=str): diff[sec] = ds[sec]
        out['scenarios'][key] = diff
    def clean(o):
        if isinstance(o, float): return None if o != o else round(o, 6)
        if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
        if isinstance(o, (list, tuple)): return [clean(v) for v in o]
        return o
    with open(os.path.join(ROOT, 'data.json'), 'w') as f: json.dump(clean(out), f, separators=(',', ':'))
    print('wrote data.json (%.1f MB)' % (os.path.getsize(os.path.join(ROOT, 'data.json')) / 1e6), file=sys.stderr)

if __name__ == '__main__':
    main()
