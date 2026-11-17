#!/usr/bin/env python3
"""
Contract check: does a modified index.html still fit the data the pipeline produces?

Usage (from the repo root, after `python3 pipeline/build.py`):
    python3 pipeline/contract_check.py --mine path/to/your/index.html [--ref index.html] [--data data.json]

What it reports (markdown, to stdout):
  1. Build facts: update date, spec version, check figures.
  2. Data-contract references in YOUR index.html that the build does not produce (fields renamed or invented).
  3. Function-level diff between YOUR index.html and the reference: which top-level JS functions changed,
     were removed or added — and which of those sit on the input -> output path (the "contract" list), so
     presentation-only changes (CSS, markup, labels) are separated from changes that alter how data is read
     or computed.
  4. Live-source URLs (ROAS ladder, CAC sheet, 2026 P&L, mapping) that differ from the reference.
A change on the contract path is not necessarily wrong — it is what to review before deploying.
"""
import argparse, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)

# Top-level functions that read data.json or compute metrics shown on the pages. A change here changes numbers.
CONTRACT = {
    'mergeData', 'computeInvestor', 'arpaAt', 'mtTierOf', 'mtCacFor', 'mtLtvCac', 'mtPayback', 'mtRoasFactor', 'mtRoasFactorIsStd',
    'mtCutKey', 'mtWaitCut', 'mtLiveReport', 'addbackActive', 'parsePeriodKey', 'periodKeyFromPicker', 'setPeriod', 'customPeriod',
    'ensurePeriod', 'buildPeriodPicker', 'buildPeriodPickerFor', 'renderKpis', 'renderInvestor', 'renderFinancials', 'renderPerformance',
    'renderSegmentation', 'renderControl', 'renderLadder', 'renderCac', 'perfCatalogue', 'assignAxes', 'annualisedPL', 'pPL',
}
PRESENTATION_HINT = {'barChart', 'lineChart', 'card', 'split2', 'esc', 'applyTheme', 'applyZoom'}

def read(p):
    with open(p, encoding='utf-8') as f: return f.read()

def script_text(html):
    return '\n'.join(re.findall(r'<script[^>]*>(.*?)</script>', html, flags=re.S))

def top_level_functions(js):
    """{name: normalised body} for `function name(` / `async function name(` declarations at column 0,
    plus anonymous IIFE blocks keyed by their first distinctive line."""
    lines = js.split('\n'); out = {}; cur = None; buf = []
    def flush():
        if cur: out[cur] = re.sub(r'\s+', ' ', '\n'.join(buf)).strip()
    starts = re.compile(r'^(async\s+)?function\s+(\w+)\s*\(')
    iife = re.compile(r'^\((async\s+)?function\s*\(\)\s*\{')
    for i, ln in enumerate(lines):
        m = starts.match(ln)
        if m: flush(); cur = m.group(2); buf = [ln]; continue
        if iife.match(ln):
            flush(); nxt = next((l.strip() for l in lines[i + 1:i + 6] if l.strip() and not l.strip().startswith('//')), '')
            cur = 'IIFE: ' + nxt[:60]; buf = [ln]; continue
        if cur and ln and not ln[0].isspace() and not ln.startswith('}') and not ln.startswith('//') and not ln.startswith('/*') and not ln.startswith(' *'):
            flush(); cur = None; buf = []
        if cur: buf.append(ln)
    flush(); return out

def live_sources(html):
    return sorted(set(re.findall(r"https://docs\.google\.com/spreadsheets/d/e/[A-Za-z0-9_-]+/pub\?[^'\"]+", html)))

def data_fields(data):
    """Every field path the build produces, flattened to dotted keys without indices."""
    keys = set()
    def walk(o, p):
        if isinstance(o, dict):
            for k, v in o.items():
                kk = (p + '.' if p else '') + str(k); keys.add(kk); walk(v, kk)
        elif isinstance(o, list) and o and isinstance(o[0], dict):
            walk(o[0], p)
    walk(data, '')
    return keys

def references(js):
    """Field references the page makes into the dataset: S.x, S.creators.x, S.pl.x, D.x, M.x, sel.x / latest.x / p.x (period fields),
    .rev.x, .active_end.x, .events.x, .retention.x, .arr.x, .cac.x, .creators.x on period objects."""
    refs = {}
    def add(kind, name, ctx):
        refs.setdefault((kind, name), ctx)
    for m in re.finditer(r'\bS\.(creators|pl|mv)\.(\w+)', js): add('series.' + m.group(1), m.group(2), m.group(0))
    for m in re.finditer(r'\bS\.(\w+)', js):
        if m.group(1) not in ('months',): add('series', m.group(1), m.group(0))
    for m in re.finditer(r'\bD\.(\w+)', js): add('base', m.group(1), m.group(0))
    for m in re.finditer(r'\bM\.(\w+)', js): add('meta', m.group(1), m.group(0))
    for m in re.finditer(r'\b(?:sel|latest|ttm|sl|cur|prev|p|pa|pb|yb)\.(rev|active_end|events|retention|arr|cac|creators|creator_cac|pl|arpa|arpa_n|ltv|churn|gmv|ads|yoy|concentration|start|end|label|months|roas_tier|leg_seg)\b(?:\.(\w+))?', js):
        add('period', m.group(1) + ('.' + m.group(2) if m.group(2) else ''), m.group(0))
    return refs

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--mine', required=True); ap.add_argument('--ref', default=os.path.join(ROOT, 'index.html')); ap.add_argument('--data', default=os.path.join(ROOT, 'data.json'))
    a = ap.parse_args()
    mine = read(a.mine); ref = read(a.ref); data = json.load(open(a.data, encoding='utf-8'))
    meta = data.get('meta', {}); checks = meta.get('checks', [])
    print('# Contract check — your index.html vs the reference build\n')
    print('## 1. Build facts')
    print('- Update date (control.json latest_month): **%s**' % meta.get('latest_month'))
    print('- Spec version: **%s** + post-handover changes (see "Changes post handover")' % meta.get('spec_version'))
    print('- Check figures: **%d/%d pass**%s' % (sum(1 for c in checks if c.get('pass')), len(checks), '' if all(c.get('pass') for c in checks) else ' — failing: ' + ', '.join(c['id'] for c in checks if not c.get('pass'))))
    print('- 2024-25 lock: %s' % next((h['title'] for h in meta.get('health', []) if h.get('id') == 'static_2024_25'), 'n/a'))
    print()

    # 2. references
    fields = data_fields(data); base = data.get('base', {}); periods = base.get('periods', {})
    pfields = set()
    for pk, pv in periods.items():
        for k in data_fields(pv): pfields.add(k)
    mine_js = script_text(mine); ref_js = script_text(ref)
    my_refs = references(mine_js); ref_refs = references(ref_js)
    missing = []
    for (kind, name), ctx in sorted(my_refs.items()):
        if kind == 'series': ok = ('base.series.' + name) in fields
        elif kind.startswith('series.'): ok = ('base.series.' + kind.split('.')[1] + '.' + name) in fields
        elif kind == 'base': ok = ('base.' + name) in fields
        elif kind == 'meta': ok = ('meta.' + name) in fields
        elif kind == 'period': ok = name in pfields
        else: ok = True
        if not ok: missing.append((kind, name, ctx, (kind, name) in ref_refs))
    print('## 2. Data fields your page reads that the build does not produce')
    if not missing: print('None — every field your page reads exists in data.json.')
    else:
        print('| Kind | Field | As written | Also in reference? |'); print('| --- | --- | --- | --- |')
        for kind, name, ctx, inref in missing: print('| %s | `%s` | `%s` | %s |' % (kind, name, ctx, 'yes (false alarm, dynamic key)' if inref else '**no — introduced after handover**'))
    print()

    # 3. function diff
    mf = top_level_functions(mine_js); rf = top_level_functions(ref_js)
    changed = sorted(n for n in mf if n in rf and mf[n] != rf[n]); removed = sorted(n for n in rf if n not in mf); added = sorted(n for n in mf if n not in rf)
    print('## 3. JavaScript functions that differ from the reference')
    print('Reference has %d top-level functions, yours %d. Changed: %d, removed: %d, added: %d.\n' % (len(rf), len(mf), len(changed), len(removed), len(added)))
    def tag(n): return '**contract — review before deploy**' if (n in CONTRACT or n.startswith('IIFE')) else ('presentation' if n in PRESENTATION_HINT else 'other')
    if changed or removed or added:
        print('| Function | Status | Path |'); print('| --- | --- | --- |')
        for n in changed: print('| `%s` | changed | %s |' % (n, tag(n)))
        for n in removed: print('| `%s` | removed | %s |' % (n, tag(n)))
        for n in added: print('| `%s` | added | %s |' % (n, tag(n)))
    else: print('No function differs — your changes are outside the script (markup or CSS), or you are comparing identical files.')
    # 3b. new number literals inside changed contract functions (a hard-coded figure where data should be read)
    lits = []
    for n in changed:
        if not (n in CONTRACT or n.startswith('IIFE')): continue
        before = set(re.findall(r'(?<![\w.])\d{4,}(?:\.\d+)?(?![\w])', rf[n])); after = set(re.findall(r'(?<![\w.])\d{4,}(?:\.\d+)?(?![\w])', mf[n]))
        new_lits = sorted(after - before, key=lambda x: -float(x))
        if new_lits: lits.append((n, new_lits[:8]))
    if lits:
        print('\nNew number literals (≥ 4 digits) inside changed contract functions — check none is a hard-coded result:')
        for n, L in lits: print('- `%s`: %s' % (n, ', '.join(L)))
    style_changed = re.findall(r'<style[^>]*>(.*?)</style>', mine, flags=re.S) != re.findall(r'<style[^>]*>(.*?)</style>', ref, flags=re.S)
    print('\nCSS (`<style>`) %s.' % ('differs — presentation only, fine' if style_changed else 'identical'))
    print()

    # 4. live sources
    ms = live_sources(mine); rs = live_sources(ref)
    print('## 4. Live Google Sheet sources')
    if ms == rs: print('Identical to the reference (%d published-sheet links).' % len(rs))
    else:
        for u in rs:
            if u not in ms: print('- **Missing in yours:** `%s`' % u)
        for u in ms:
            if u not in rs: print('- **Added in yours (not in reference):** `%s`' % u)
    print('\n## How to read this')
    print('Section 2 lists real breaks: the page asks for data the pipeline does not make. Section 3 "contract" rows are the functions that decide '
          'what a number means; a change there must be deliberate and should be checked against the locked definitions. "presentation" and CSS changes do not affect numbers.')

if __name__ == '__main__':
    main()
