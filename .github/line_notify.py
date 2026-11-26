# Sends the game's big updates to the LINE Official Account's friends (Messaging API broadcast).
# Runs on GitHub Actions after a push. "Big" = the middle number of the version went up (4.4.x -> 4.5.0)
# or the first one did (4.x -> 5.0.0); small fixes (4.4.0 -> 4.4.1) are not sent, to save the monthly message quota.
# The channel access token is the repository secret LINE_TOKEN; without it nothing is sent.
import json, os, re, subprocess, sys, urllib.request

SITE = 'https://ameft-oyaji.github.io/slalom-game/'
VER_RE = re.compile(r"const APP_VERSION = '(\d+)\.(\d+)\.(\d+)'")
ENTRY_RE = re.compile(r"^  \{ v: '(\d+\.\d+\.\d+)', d: '([^']*)', items: \[(.*)\] \},\s*$", re.M)
STR_RE = re.compile(r"'((?:[^'\\]|\\.)*)'")


def version_of(src):
    m = VER_RE.search(src or '')
    return tuple(int(x) for x in m.groups()) if m else None


def entries(src):
    out = []
    for v, d, items in ENTRY_RE.findall(src):
        out.append((tuple(int(x) for x in v.split('.')), d, [s.replace("\\'", "'") for s in STR_RE.findall(items)]))
    return out


def previous_index():
    try:
        return subprocess.run(['git', 'show', 'HEAD~1:index.html'], capture_output=True, text=True, encoding='utf-8', check=True).stdout
    except Exception:
        return ''


def main():
    force = os.environ.get('FORCE') == 'true'
    token = os.environ.get('LINE_TOKEN', '').strip()
    src = open('index.html', encoding='utf-8').read()
    new, old = version_of(src), version_of(previous_index())
    if not new:
        print('no version found'); return
    big = old is None or new[:2] > old[:2]
    if not (big or force):
        print(f'small update {old} -> {new}: not sent'); return
    # every change since the last big version the friends were told about (old's x.y.0), newest first
    since = (old[0], old[1], 0) if old and not force else (new[0], new[1], 0)
    rows = [e for e in entries(src) if e[0] > since or e[0] == new] if not force else [e for e in entries(src) if e[0][:2] == new[:2]]
    lines = [f'スラローム・ワールドカップ Ver. {".".join(map(str, new))} に更新しました']
    for v, d, items in rows[:6]:
        lines.append('')
        lines.append(f'■ Ver. {".".join(map(str, v))}（{d}）')
        lines += ['・' + t for t in items]
    lines += ['', '▶ 遊ぶ：' + SITE, '▶ メニュー案内：' + SITE + 'menu-guide.html']
    text = '\n'.join(lines)
    if len(text) > 4900:
        text = text[:4850] + '…\n\n▶ 遊ぶ：' + SITE
    print(text)
    if not token:
        print('LINE_TOKEN is not set: nothing sent'); return
    req = urllib.request.Request('https://api.line.me/v2/bot/message/broadcast',
                                 data=json.dumps({'messages': [{'type': 'text', 'text': text}]}).encode('utf-8'),
                                 headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token}, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            print('sent:', r.status)
    except urllib.error.HTTPError as e:
        print('LINE error:', e.code, e.read().decode('utf-8', 'replace')); sys.exit(1)


if __name__ == '__main__':
    main()
