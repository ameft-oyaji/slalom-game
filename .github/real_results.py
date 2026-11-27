# Reads the real World Cup podiums of the season from Wikipedia's season page and writes real-results.json
# for the game's calendar ("実際の結果"). Run by .github/workflows/real-results.yml every few hours.
#   python real_results.py [out.json] [--wikitext saved.txt]
import json, re, sys, urllib.request, urllib.parse, datetime

PAGE = '2026–27 FIS Alpine Ski World Cup'
MONTHS = {m: i + 1 for i, m in enumerate(['January', 'February', 'March', 'April', 'May', 'June', 'July',
                                          'August', 'September', 'October', 'November', 'December'])}
DISCS = ('DH', 'SG', 'GS', 'SL', 'AC', 'PAR', 'PSL', 'PGS', 'TC')


def fetch_wikitext():
    q = urllib.parse.urlencode({'action': 'parse', 'page': PAGE, 'prop': 'wikitext', 'format': 'json', 'formatversion': 2})
    req = urllib.request.Request('https://en.wikipedia.org/w/api.php?' + q,
                                 headers={'User-Agent': 'slalom-game-results/1.0 (https://github.com/ameft-oyaji/slalom-game)'})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)['parse']['wikitext']


def short_name(full):
    full = re.sub(r'\s*\(.*?\)\s*', ' ', full).strip()
    parts = full.split(' ')
    if len(parts) < 2:
        return full
    return parts[0][0] + '. ' + ' '.join(parts[1:])


def podium_cell(cell):
    """every racer in a podium cell (ties put two or more): [[name, NAT], ...]"""
    out = []
    for m in re.finditer(r'\{\{flagicon\|([A-Z]{3})[^}]*\}\}(?:&nbsp;|\s)*\[\[([^\]|]+)(?:\|([^\]]+))?\]\]', cell):
        out.append([short_name(m.group(3) or m.group(2)), m.group(1)])
    return out


def calendar(section):
    i = section.find('=== Calendar')
    if i < 0:
        return {}
    j = section.find('\n|}', i)
    table = section[i:j]
    races = {}
    for row in re.split(r'\n\|-', table)[1:]:
        lines = [l for l in row.split('\n') if l.startswith('|') and not l.startswith('|+')]
        date = None
        for l in lines:
            m = re.search(r'(\d{1,2}) (' + '|'.join(MONTHS) + r') (\d{4})', l)
            if m:
                date = datetime.date(int(m.group(3)), MONTHS[m.group(2)], int(m.group(1))).isoformat()
                break
        if not date:
            continue
        k = next((n for n, l in enumerate(lines) if re.search(r'\|\s*(' + '|'.join(DISCS) + r')\b\s*(<sub>|$)', l)), None)
        if k is None:
            continue
        d = re.search(r'\|\s*(' + '|'.join(DISCS) + r')\b', lines[k]).group(1)
        pod = [podium_cell(c) for c in lines[k + 1:k + 4]]
        if not pod or not pod[0]:
            continue                                   # not raced yet (or cancelled)
        races.setdefault(date, []).append({'d': d, 'p': pod})
    return races


def parse(w):
    men = re.search(r'^== Men\s*==\s*$', w, re.M)
    women = re.search(r'^== (Women|Ladies)\s*==\s*$', w, re.M)
    end = re.search(r'^== Nations Cup', w, re.M)
    if not (men and women):
        raise SystemExit('page layout changed: no Men / Women sections')
    return {'men': calendar(w[men.end():women.start()]),
            'women': calendar(w[women.end():end.start() if end else len(w)])}


def main():
    args = sys.argv[1:]
    out = 'real-results.json'
    w = None
    if '--wikitext' in args:
        n = args.index('--wikitext')
        w = open(args[n + 1], encoding='utf-8').read()
        del args[n:n + 2]
    if args:
        out = args[0]
    res = parse(w if w is not None else fetch_wikitext())
    if not res['men'] and not res['women']:
        raise SystemExit('no results found: page layout changed?')
    try:
        old = json.load(open(out, encoding='utf-8'))
    except Exception:
        old = {}
    if old.get('men') == res['men'] and old.get('women') == res['women']:
        print('no change')
        return
    res = {'source': 'https://en.wikipedia.org/wiki/' + PAGE.replace(' ', '_'),
           'updated': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%MZ'), **res}
    json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print('written', sum(len(v) for v in res['men'].values()), 'men /', sum(len(v) for v in res['women'].values()), 'women races')


if __name__ == '__main__':
    main()
