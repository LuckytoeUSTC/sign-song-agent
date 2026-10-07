"""Build a disposable reverse index from existing rule and case records."""
from pathlib import Path
import re
from build_case_catalog import ROOT, read_case
from search_cases import read_sheets


def collect(sheets=None):
    folder = ROOT / '资料/规则与例句数据库'
    if sheets is None:
        sheets = read_sheets(folder / '转写例句.xlsx')
    examples = sheets['转写例句']
    cases = [read_case(p) for p in (ROOT / '资料/讲义制作过程存档').glob('*/转写修改与辨析案例/*.md')
             if re.match(r'.+-M\d+ ', p.name)]
    def source_location(value):
        if value:
            try:
                return str(Path(value).relative_to(ROOT)).replace('\\', '/')
            except ValueError:
                return value
        return ''
    records = {}
    for row in examples:
        records[row['编号']] = (row['原句'], row['手语转写'], row.get('上下文', ''),
                              source_location(row.get('查看原件', '')), row.get('关联规则', ''))
    for row in cases:
        records[row['编号']] = (row['title'], row['核心辨析'], row['何时参考'],
                              source_location(str(row['_path'])), row.get('关联规则', ''))
    rules = {}
    paths = [folder / '公共转写规则.md', *(ROOT / '资料/讲义制作过程存档').glob('*/本曲转写规则.md')]
    for path in paths:
        for match in re.finditer(r'^## (R\d+|待审\d+|\w+-R\d+) ([^\n]+)\n(.*?)(?=^## |\Z)',
                                 path.read_text(encoding='utf-8-sig'), re.M | re.S):
            number, title, body = match.groups()
            rules[number] = (title.strip(), body.strip(), source_location(str(path)))
    associations = {}
    for number, record in records.items():
        for rule in re.split(r'[,，\s]+', record[4].strip()):
            if rule in rules:
                associations[(rule, number)] = '案例关联标注'
    for rule, (_, body, _) in rules.items():
        for number in re.findall(r'E\d{4}|\w+-M\d+', body):
            if number in records:
                associations.setdefault((rule, number), '规则正文提及')
    return records, rules, associations


def build():
    folder = ROOT / '资料/规则与例句数据库'
    sheets = read_sheets(folder / '转写例句.xlsx')
    records, rules, associations = collect(sheets)
    import json
    data = {'rules': {r: {'title': v[0], 'body': v[1], 'source': v[2]} for r, v in rules.items()},
            'links': [{'rule': r, 'case': c, 'basis': b} for (r, c), b in sorted(associations.items())]}
    (folder/'规则与例句索引.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    from build_query_page import build as build_page
    build_page((records, rules, associations), sheets)
    print(f'反向索引：{len(rules)}条转写规则／待审建议，{len(associations)}条案例关联。')


if __name__ == '__main__':
    build()

