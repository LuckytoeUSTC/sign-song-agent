from pathlib import Path
from urllib.parse import unquote, quote, urlsplit
import hashlib
import html
import json
import re
import shutil
import markdown

import os
SRC = Path(__file__).resolve().parents[2]
BASE = SRC.parent
OUT = Path(os.environ.get('NO_AI_OUTPUT', str(BASE / 'NO-AI-ZONE'))).resolve()
BUILD = Path(os.environ.get('NO_AI_BUILD', str(BASE / '构建记录/NO-AI-ZONE'))).resolve()
BUILD.mkdir(parents=True, exist_ok=True)
raw = (SRC / 'scripts/query_page.html').read_text(encoding='utf-8-sig')
data = json.loads(re.search(r'<script id="data" type="application/json">(.*?)</script>', raw, re.S).group(1))

def target(p):
    rel = p.relative_to(SRC)
    return OUT / (rel.with_suffix('.html') if p.suffix == '.md' else rel)

def excluded(p):
    return any(x in {'LLM-demo', '工作台', 'scripts', 'dictionaries'} for x in p.parts) or p.name in {
        'AGENTS.md', 'TOOLING.md', '数据库使用方法.md', '转写规则提炼待审.md', '本次讲义格式要求.md'}

pending = [SRC / '资料' / unquote(x['link']) for x in data['entries'] + data['rules'] if x.get('link')]
selected = set()
removed = []
missing = []
link_pattern = re.compile(r'(!?\[[^\]]*\])\((<[^>]*>|[^)]+)\)')

def content(p):
    text = p.read_text(encoding='utf-8-sig')
    if p.name == '公共转写规则.md':
        first = text.index('## R01 ')
        text = '# 公共转写规则\n\n这些经验供查阅和比较，具体表达结合语境与老师试打决定。\n\n' + text[first:]
    return text

def linked(p, value):
    value = html.unescape(value.strip('<>'))
    url = urlsplit(value)
    if url.scheme or url.netloc or not url.path:
        return None
    q = (p.parent / unquote(url.path)).resolve()
    if not q.is_relative_to((SRC / '资料').resolve()):
        return None
    return q

while pending:
    p = pending.pop().resolve()
    if p in selected or excluded(p):
        continue
    if p.is_dir():
        pending.extend(x for x in p.iterdir() if x.is_file() and x.suffix in {'.md', '.pdf', '.docx', '.doc'})
        continue
    if not p.exists():
        missing.append(str(p))
        continue
    selected.add(p)
    if p.suffix == '.md':
        for match in link_pattern.finditer(content(p)):
            q = linked(p, match.group(2))
            if q and not excluded(q):
                pending.append(q)

OUT.mkdir(exist_ok=True)
css = '''body{font-family:"Microsoft YaHei",sans-serif;color:#26372e;background:#f7f8f5;max-width:960px;margin:32px auto;padding:0 24px;line-height:1.8}h1,h2{line-height:1.4}a{color:#35654b}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:8px}img{max-width:100%}pre{white-space:pre-wrap}blockquote{border-left:4px solid #b9cdbd;padding-left:16px}'''
script = '''document.addEventListener('click',async e=>{let a=e.target.closest('a[data-source]');if(!a||!window.pywebview?.api)return;e.preventDefault();let r=await window.pywebview.api.open_source(a.dataset.source);if(!r.ok)alert(r.error);});'''
manifest = []
for p in sorted(selected):
    dst = target(p)
    dst.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix != '.md':
        shutil.copy2(p, dst)
    else:
        def rewrite(m):
            label, value = m.groups()
            q = linked(p, value)
            if q is None:
                return m.group()
            if q.is_dir():
                options = [x for x in selected if x.parent == q]
                return label.strip('[]') + '：' + '、'.join(
                    f'[{x.stem}](<{quote(str(target(x).relative_to(OUT / "资料")).replace(chr(92), "/"), safe="/")}>)'
                    for x in sorted(options))
            if q not in selected:
                removed.append({'page': str(p.relative_to(SRC)), 'link': str(q.relative_to(SRC)) if q.is_relative_to(SRC) else str(q)})
                return label.lstrip('!').strip('[]') + '（本查阅包未收录）'
            return f'{label}(<{quote(str(target(q).relative_to(OUT / "资料")).replace(chr(92), "/"), safe="/")}>)'
        text = link_pattern.sub(rewrite, content(p))
        body = markdown.markdown(text, extensions=['tables', 'fenced_code', 'sane_lists'])
        def anchor(m):
            href = html.unescape(m.group(1))
            if not urlsplit(href).scheme and not href.startswith('#'):
                return 'href="' + html.escape(href, quote=True) + '" data-source="' + html.escape(href, quote=True) + '"'
            return m.group()
        body = re.sub(r'href="([^"]*)"', anchor, body)
        csp = "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src data: file:"
        dst.write_text(f'<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="{csp}"><title>{html.escape(p.stem)}</title><style>{css}</style><body>{body}<script>{script}</script></body></html>', encoding='utf-8')
    manifest.append({'source': str(p.relative_to(SRC)), 'file': str(dst.relative_to(OUT)), 'sha256': hashlib.sha256(dst.read_bytes()).hexdigest()})

for item in data['entries'] + data['rules']:
    if item.get('link'):
        p = (SRC / '资料' / unquote(item['link'])).resolve()
        item['link'] = quote(str(target(p).relative_to(OUT / '资料')).replace('\\', '/'), safe='/')
payload = json.dumps(data, ensure_ascii=False).replace('<', '\\u003c')
raw = re.sub(r'(<script id="data" type="application/json">).*?(</script>)', lambda m: m.group(1) + payload + m.group(2), raw, flags=re.S)
(BUILD / 'query_page.html').write_text(raw, encoding='utf-8')
for name in ['STOP-USING-AI.jpg', 'STOP-AI.jpg']:
    shutil.copy2(SRC / 'scripts/release/assets' / name, OUT / name)
shutil.copy2(SRC / 'LICENSE', OUT / '许可说明.txt')
report = {'entries': len(data['entries']), 'rules': len(data['rules']), 'resources': manifest, 'excluded_links': removed, 'missing_links': sorted(set(missing))}
(BUILD / '资料核查.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'entries': len(data['entries']), 'rules': len(data['rules']), 'resources': len(manifest), 'missing': sorted(set(missing)), 'excluded_links': len(removed)}, ensure_ascii=False))
