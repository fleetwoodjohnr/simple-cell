import sys, re, json, html
src = open(sys.argv[1], encoding='utf-8').read()
rows = re.findall(r'<tr[^>]*>(.*?)</tr>', src, re.S)
txt = lambda s: re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', ' ', s))).strip()
out = []
for r in rows:
    cells = re.findall(r'<td[^>]*>(.*?)</td>', r, re.S)
    if len(cells) < 2: continue
    links = re.findall(r'<a [^>]*>(.*?)</a>', cells[0], re.S)
    rest = txt(re.sub(r'<a [^>]*>.*?</a>', '', cells[0], flags=re.S))
    ver = re.search(r'\((\d{4}|Office 365[^)]*|Microsoft 365[^)]*|[^)]*)\)', rest)
    desc = txt(cells[1])
    m = re.match(r'^([A-Za-z ,]+):\s*(.*)$', desc)
    cat, d = (m.group(1), m.group(2)) if m else ('', desc)
    for l in links:
        for name in re.split(r'[,\s]+', txt(l)):
            if re.match(r'^[A-Z][A-Z0-9_.]*$', name):
                out.append({'name': name, 'category': cat, 'description': d, 'since': ver.group(1) if ver else None})
json.dump(out, sys.stdout, indent=0)
