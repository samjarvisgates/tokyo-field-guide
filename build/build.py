#!/usr/bin/env python3
"""tokyo_workspots_1.xlsx  ->  data.json   (stdlib only)

Reads the `priority` and `phase2_signals` tabs plus the hyperlinks on the
priority tab, joins coordinates/addresses from build/apify.json (Google Maps
scrape output keyed by place ID) and writes ../data.json for index.html.

usage:  python3 build/build.py [path/to/tokyo_workspots_1.xlsx]
"""
import json, os, re, sys, zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
XLSX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, '..', 'tokyo-working', 'tokyo_workspots_1.xlsx')
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# ---------------------------------------------------------------- xlsx reader
def read_xlsx(path):
    z = zipfile.ZipFile(path)
    ss = ET.fromstring(z.read('xl/sharedStrings.xml'))
    strings = [''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])) for si in ss.findall('m:si', NS)]
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}

    def colidx(ref):
        n = 0
        for ch in re.match(r'[A-Z]+', ref).group(0):
            n = n * 26 + ord(ch) - 64
        return n - 1

    sheets, links = {}, {}
    for sh in wb.find('m:sheets', NS):
        name = sh.get('name')
        target = relmap[sh.get(RNS + 'id')]
        path = 'xl/' + target
        root = ET.fromstring(z.read(path))
        rows = {}
        for row in root.iter('{%s}row' % NS['m']):
            r = int(row.get('r'))
            vals = {}
            for c in row.findall('m:c', NS):
                v = c.find('m:v', NS)
                t = c.get('t')
                if v is None:
                    isel = c.find('m:is', NS)
                    val = ''.join(x.text or '' for x in isel.iter('{%s}t' % NS['m'])) if isel is not None else None
                elif t == 's':
                    val = strings[int(v.text)]
                elif t == 'b':
                    val = bool(int(v.text))
                elif v.text is None:
                    val = None
                else:
                    try:
                        val = float(v.text)
                        val = int(val) if val == int(val) else val
                    except ValueError:
                        val = v.text
                vals[colidx(c.get('r'))] = val
            if vals:
                rows[r] = vals
        sheets[name] = rows
        # hyperlinks
        relpath = 'xl/worksheets/_rels/' + os.path.basename(target) + '.rels'
        if relpath in z.namelist():
            rr = ET.fromstring(z.read(relpath))
            rm = {x.get('Id'): x.get('Target') for x in rr}
            lk = {}
            for h in root.iter('{%s}hyperlink' % NS['m']):
                ref = h.get('ref'); rid = h.get(RNS + 'id')
                if rid in rm:
                    m = re.match(r'([A-Z]+)(\d+)', ref)
                    lk.setdefault(int(m.group(2)), {})[colidx(m.group(1))] = rm[rid]
            links[name] = lk
    return sheets, links

def table(rows):
    """rows dict -> (header list, list of (rownum, dict by header))"""
    hdr_row = rows[min(rows)]
    hdr = [hdr_row.get(i) for i in range(max(hdr_row) + 1)]
    out = []
    for r in sorted(rows):
        if r == min(rows):
            continue
        d = {}
        for i, h in enumerate(hdr):
            if h is not None:
                d[h] = rows[r].get(i)
        out.append((r, d))
    return hdr, out

# ---------------------------------------------------------------- sections
INK = {'V': '#C8452B', 'C': '#2F5FA8', 'M': '#D8951F', 'P': '#2E6B52', 'U': '#8A4A72'}
SECTIONS = [
    ('V', 'SHIBUYA & HARAJUKU', ['渋谷', '表参道/原宿', '恵比寿/代官山']),
    ('C', 'GINZA & THE CENTRE', ['銀座/日比谷', '東京駅/丸の内/日本橋', '赤坂/虎ノ門/新橋', '六本木/麻布/広尾', '田町/浜松町/芝']),
    ('M', 'KANDA, SHINJUKU & THE WEST', ['神保町/御茶ノ水', '秋葉原/神田', '新宿', '高田馬場/早稲田/神楽坂', '池袋/大塚/巣鴨', '中野', '高円寺/荻窪', '吉祥寺/三鷹', '立川/三多摩', '町田']),
    ('P', 'MEGURO & SETAGAYA', ['中目黒', '品川/五反田/目黒', '自由が丘/学芸大学/二子玉川', '下北沢']),
    ('U', 'YANAKA & THE EAST', ['上野/御徒町', '浅草/蔵前/浅草橋', '清澄白河/門前仲町', '錦糸町/押上/両国', '日暮里/谷根千', '北千住/足立', '赤羽/王子']),
]
SECTION_OF = {g: k for k, _, gs in SECTIONS for g in gs}
AREA_EN = {
    '渋谷': 'SHIBUYA', '新宿': 'SHINJUKU', '池袋/大塚/巣鴨': 'IKEBUKURO · SUGAMO', '神保町/御茶ノ水': 'JIMBOCHO',
    '中目黒': 'NAKAMEGURO', '清澄白河/門前仲町': 'KIYOSUMI', '東京駅/丸の内/日本橋': 'MARUNOUCHI · NIHONBASHI',
    '吉祥寺/三鷹': 'KICHIJOJI', '恵比寿/代官山': 'EBISU · DAIKANYAMA', '表参道/原宿': 'OMOTESANDO · HARAJUKU',
    '六本木/麻布/広尾': 'ROPPONGI · AZABU', '銀座/日比谷': 'GINZA · HIBIYA', '品川/五反田/目黒': 'SHINAGAWA · MEGURO',
    '上野/御徒町': 'UENO', '赤坂/虎ノ門/新橋': 'AKASAKA · SHIMBASHI', '浅草/蔵前/浅草橋': 'ASAKUSA · KURAMAE',
    '下北沢': 'SHIMOKITAZAWA', '高円寺/荻窪': 'KOENJI · OGIKUBO', '自由が丘/学芸大学/二子玉川': 'JIYUGAOKA · FUTAKO',
    '中野': 'NAKANO', '北千住/足立': 'KITA-SENJU', '秋葉原/神田': 'AKIHABARA · KANDA',
    '高田馬場/早稲田/神楽坂': 'TAKADANOBABA · KAGURAZAKA', '田町/浜松町/芝': 'TAMACHI · SHIBA', '立川/三多摩': 'TACHIKAWA',
    '錦糸町/押上/両国': 'KINSHICHO · OSHIAGE', '町田': 'MACHIDA', '日暮里/谷根千': 'NIPPORI · YANAKA', '赤羽/王子': 'AKABANE · OJI',
}

# ---------------------------------------------------------------- hours
DAYS = ['月', '火', '水', '木', '金', '土', '日']
RANGE_RE = re.compile(r'(\d{1,2})\s*[:時]\s*(\d{2})\s*分?\s*[-–～〜~]\s*(\d{1,2})\s*[:時]\s*(\d{2})\s*分?')

def _day_set(tok):
    """'月-金', '月〜日', '土,日', '土日', '毎日', '月' -> list of day indexes"""
    tok = tok.replace('曜日', '').replace('曜', '').replace('祝', '')
    if '毎日' in tok or tok in ('月-日', '月〜日', '月～日'):
        return list(range(7))
    m = re.match(r'^([月火水木金土日])\s*[-〜～~]\s*([月火水木金土日])$', tok)
    if m:
        a, b = DAYS.index(m.group(1)), DAYS.index(m.group(2))
        return list(range(a, b + 1)) if a <= b else list(range(a, 7)) + list(range(0, b + 1))
    days = []
    for part in re.split(r'[,、・]', tok):
        part = part.strip()
        m = re.match(r'^([月火水木金土日])\s*[-〜～~]\s*([月火水木金土日])$', part)
        if m:
            a, b = DAYS.index(m.group(1)), DAYS.index(m.group(2))
            days += list(range(a, b + 1)) if a <= b else list(range(a, 7)) + list(range(0, b + 1))
        else:
            days += [DAYS.index(c) for c in part if c in DAYS]
    return days

def parse_hours(raw):
    """-> list of 7 (Mon..Sun): list of [openMin, closeMin] (close > 1440 = past midnight), [] = closed, None = unknown"""
    if not raw:
        return None
    out = [None] * 7
    for piece in re.split(r'[;\n]', raw):
        piece = re.sub(r'（.*?）|\(.*?\)', '', piece).strip()
        if not piece:
            continue
        m = re.match(r'^((?:毎日|[月火水木金土日祝](?:曜日?)?)(?:\s*[-〜～~,、・]\s*[月火水木金土日祝](?:曜日?)?)*(?:[月火水木金土日祝]+)?)\s*(.*)$', piece)
        if not m:
            continue
        days = _day_set(m.group(1)); rest = m.group(2).replace('翌', '')
        if not days:
            continue
        if '24' in rest and '時間' in rest:
            val = [[0, 1440]]
        elif '定休' in rest or '休業' in rest or 'closed' in rest.lower() or rest.strip() in ('休', '休み', '閉店'):
            val = []
        else:
            val = []
            for a, b, c, e in RANGE_RE.findall(rest):
                o = int(a) * 60 + int(b); cl = int(c) * 60 + int(e)
                if cl <= o:
                    cl += 1440
                val.append([o, cl])
            if not val:
                val = None
        for d in days:
            out[d] = val
    if all(v is None for v in out):
        return None
    return out

# ---------------------------------------------------------------- helpers
def s(v):
    return (str(v).strip() if v is not None else '')

def split_lines(v):
    return [x.strip() for x in s(v).split('\n') if x.strip()]

def price_band(v):
    v = s(v).replace('￥', '¥').replace('〜', '～')
    if not v or not v.startswith('¥'):
        return None
    m = re.match(r'¥([\d,]+)～([\d,]+)', v)
    if not m:
        return None
    lo, hi = m.group(1), m.group(2)
    if lo.replace(',', '') == '1':
        return 'UNDER ¥' + hi
    return '¥' + lo + '–' + hi

def quotes(v):
    out = []
    for line in split_lines(v):
        m = re.match(r'^(\d{4}-\d{2}):\s*(.*)$', line)
        if m:
            out.append({'date': m.group(1), 'text': m.group(2)})
        elif line:
            out.append({'date': '', 'text': line})
    return out

def tags(v):
    t = [x.strip().upper() for x in re.split(r'[,、;]', s(v)) if x.strip()]
    return t[:6]

# ---------------------------------------------------------------- main
def main():
    sheets, links = read_xlsx(XLSX)
    _, prio = table(sheets['priority'])
    _, sig = table(sheets['phase2_signals'])
    plinks = links.get('priority', {})
    sig_by_name = {d['Name (JA)']: d for _, d in sig if d.get('Name (JA)')}
    try:
        apify = json.load(open(os.path.join(HERE, 'apify.json')))
    except FileNotFoundError:
        apify = []
    geo = {}
    for it in apify:
        if isinstance(it, list):  # compact [placeId, lat, lng, address]
            geo[it[0]] = {'lat': it[1], 'lng': it[2], 'address': it[3]}
            continue
        pid = it.get('placeId'); loc = it.get('location') or {}
        if pid and loc.get('lat'):
            geo[pid] = {'lat': loc['lat'], 'lng': loc['lng'], 'address': it.get('address')}
    photos_dir = os.path.join(ROOT, 'photos')
    try:
        extra = json.load(open(os.path.join(HERE, 'geocode_extra.json')))
    except FileNotFoundError:
        extra = {}

    entries, stats = [], {'hours_parsed': 0, 'hours_partial': 0, 'hours_none': 0, 'coords': 0, 'photos': 0, 'tabelog': 0}
    for rownum, d in prio:
        if not d.get('#') or not d.get('Name (EN)'):
            continue
        num = int(d['#']); eid = '%03d' % num
        area = s(d['Area group'])
        sec = SECTION_OF.get(area)
        if not sec:
            print('!! no section for area', area, file=sys.stderr); sec = 'M'
        sg = sig_by_name.get(d['Name (JA)'], {})
        hours_raw = s(d.get('Hours')) or s(sg.get('Google hours'))
        hours = parse_hours(hours_raw)
        if hours is None:
            stats['hours_none'] += 1
        elif any(v is None for v in hours):
            stats['hours_partial'] += 1
        else:
            stats['hours_parsed'] += 1
        lk = plinks.get(rownum, {})
        map_url, photo_url, tabelog_url = lk.get(31), lk.get(32), lk.get(33)
        pid = None
        if map_url:
            m = re.search(r'place_id:([\w-]+)', map_url)
            pid = m.group(1) if m else None
        g = geo.get(pid) if pid else None
        approx = False
        if not g and eid in extra:
            g = {'lat': extra[eid]['lat'], 'lng': extra[eid]['lng'], 'address': None}; approx = True
        if g:
            stats['coords'] += 1
        photo = None
        if os.path.exists(os.path.join(photos_dir, eid + '.jpg')):
            photo = 'photos/' + eid + '.jpg'; stats['photos'] += 1
        gallery = ['photos/%s-%d.jpg' % (eid, n) for n in range(1, 7) if os.path.exists(os.path.join(photos_dir, '%s-%d.jpg' % (eid, n)))]
        tier = s(d['Tier'])
        tab_rating = sg.get('Tabelog rating')
        tabelog_url = tabelog_url or (s(sg.get('Tabelog URL')) or None)
        if tabelog_url:
            stats['tabelog'] += 1
        e = {
            'id': eid,
            'nameEn': s(d['Name (EN)']), 'nameJa': s(d['Name (JA)']),
            'areaGroup': area, 'areaEn': AREA_EN.get(area, area), 'station': s(d['Area / station']),
            'section': sec, 'ink': INK[sec],
            'tier': tier[:1], 'tierLabel': tier, 'session': s(d['Work session']),
            'visitOnly': tier.startswith('X'), 'notCafe': s(d.get('Not really a café')) == 'yes',
            'confidence': s(d.get('Confidence')), 'status': s(d.get('Phase-2 status')),
            'vibe': d['Vibe'], 'uniq': d['Uniqueness'], 'food': d['Food/drink'], 'work': d['Work-fit'],
            'sheetScore': d.get('Ranked score (work-fit gate applied)') or d.get('Vibe score') or 0,
            'why': {'vibe': s(d['Why: vibe']), 'uniq': s(d['Why: uniqueness']), 'food': s(d['Why: food/drink']), 'work': s(d['Why: work-fit'])},
            'pitch': s(d['Pitch']), 'tags': tags(d['Vibe tags']), 'order': split_lines(d['What to order']),
            'best': s(d['Best time & seat']), 'workNotes': s(d['Work notes']), 'headsUp': s(d['Heads-up']),
            'quotes': quotes(d['Evidence (verbatim review quotes)']),
            'summaryEn': s(sg.get('Review summary (EN)')),
            'signals': {'outlets': s(sg.get('Outlets')), 'wifi': s(sg.get('Wi-Fi')), 'laptop': s(sg.get('Laptop work')),
                        'crowding': s(sg.get('Crowding')), 'timeLimit': s(sg.get('Time limits (from reviews)'))},
            'google': {'rating': d.get('Google rating'), 'reviews': d.get('Google reviews'),
                       'priceBand': price_band(d.get('Price')), 'placeId': pid, 'mapUrl': map_url, 'photoUrl': photo_url},
            'tabelog': {'url': tabelog_url, 'rating': tab_rating if isinstance(tab_rating, (int, float)) else None},
            'hours': hours, 'hoursRaw': hours_raw or None,
            'lat': g['lat'] if g else None, 'lng': g['lng'] if g else None,
            'address': (g['address'] if g else None), 'approx': approx,
            'photo': photo, 'gallery': gallery,
        }
        entries.append(e)

    ids = [e['id'] for e in entries]
    assert len(ids) == len(set(ids)), 'duplicate ids'
    out = {
        'generated': 'from tokyo_workspots_1.xlsx (priority + phase2_signals)',
        'sections': [{'key': k, 'ink': INK[k], 'name': n, 'count': sum(1 for e in entries if e['section'] == k)} for k, n, _ in SECTIONS],
        'entries': entries,
    }
    with open(os.path.join(ROOT, 'data.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print('entries:', len(entries))
    print('sections:', [(x['name'], x['count']) for x in out['sections']])
    print('stats:', stats)
    missing = [e['nameEn'] for e in entries if e['lat'] is None]
    print('no coords (%d):' % len(missing), missing)

if __name__ == '__main__':
    main()
