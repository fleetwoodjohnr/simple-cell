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
        for raw in re.split(r'[,\s]+', txt(l)):
            # The page has the odd typo ("BETA.INVn"): keep the upper-case function name.
            m = re.match(r'^[A-Z][A-Z0-9_.]*', raw)
            name = m.group(0).rstrip('.') if m else ''
            if name and len(name) >= 1 and (name == raw or raw[len(name):].islower()):
                out.append({'name': name, 'category': cat, 'description': d, 'since': ver.group(1) if ver else None})
json.dump(out, sys.stdout, indent=0)
