#!/usr/bin/env python3
"""Download each entry's Google Maps main photo into photos/<id>.jpg and resize to 480px wide (macOS sips).
usage: python3 build/fetch.py   (run after build.py; re-run build.py afterwards so data.json picks up the photo paths)"""
import json, os, subprocess, sys, urllib.request
from concurrent.futures import ThreadPoolExecutor
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
d = json.load(open(os.path.join(ROOT, 'data.json')))
os.makedirs(os.path.join(ROOT, 'photos'), exist_ok=True)
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'

def one(e):
    url = e['google'].get('photoUrl')
    out = os.path.join(ROOT, 'photos', e['id'] + '.jpg')
    if not url or os.path.exists(out):
        return e['id'], 'skip'
    try:
        # ask Google for a 640px-wide variant
        u = url.split('=')[0] + '=w640-h480-k-no' if '=' in url else url
        req = urllib.request.Request(u, headers={'User-Agent': UA})
        data = urllib.request.urlopen(req, timeout=30).read()
        tmp = out + '.tmp'
        open(tmp, 'wb').write(data)
        subprocess.run(['sips', '-s', 'format', 'jpeg', '-s', 'formatOptions', '72', '--resampleWidth', '480', tmp, '--out', out],
                       check=True, capture_output=True)
        os.remove(tmp)
        return e['id'], 'ok'
    except Exception as ex:
        return e['id'], 'ERR ' + str(ex)[:80]

with ThreadPoolExecutor(6) as ex:
    res = list(ex.map(one, d['entries']))
ok = sum(1 for _, r in res if r == 'ok'); err = [(i, r) for i, r in res if r.startswith('ERR')]
print('downloaded', ok, 'errors', len(err))
for i, r in err[:20]:
    print(' ', i, r)

# ---- extra images (build/images.json: placeId -> [urls]) -> photos/<id>-N.jpg
def extra():
    try:
        imgs = json.load(open(os.path.join(ROOT, 'build', 'images.json')))
    except FileNotFoundError:
        return
    jobs = []
    for e in d['entries']:
        pid = e['google'].get('placeId')
        for n, url in enumerate((imgs.get(pid) or [])[:5], 1):
            out = os.path.join(ROOT, 'photos', '%s-%d.jpg' % (e['id'], n))
            if not os.path.exists(out):
                jobs.append((url, out))
    def one(j):
        url, out = j
        try:
            u = url.split('=')[0] + '=w800-h600-k-no' if '=' in url else url
            data = urllib.request.urlopen(urllib.request.Request(u, headers={'User-Agent': UA}), timeout=30).read()
            tmp = out + '.tmp'; open(tmp, 'wb').write(data)
            subprocess.run(['sips', '-s', 'format', 'jpeg', '-s', 'formatOptions', '62', '--resampleWidth', '640', tmp, '--out', out], check=True, capture_output=True)
            os.remove(tmp); return 'ok'
        except Exception as ex:
            return 'ERR ' + str(ex)[:60]
    with ThreadPoolExecutor(8) as ex:
        res = list(ex.map(one, jobs))
    print('extra images:', sum(1 for r in res if r == 'ok'), 'ok,', sum(1 for r in res if r != 'ok'), 'failed of', len(jobs))

if __name__ == '__main__' or True:
    extra()
