"""Build a disposable reverse index from existing rule and case records."""
from pathlib import Path
import re
from build_case_catalog import ROOT, read_case, save_xlsx
from search_cases import read_sheets


def collect():
    folder = ROOT / '资料/规则与例句数据库'
    examples = read_sheets(folder / '转写例句.xlsx')['转写例句']
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
    records, rules, associations = collect()
    folder = ROOT / '资料/规则与例句数据库'
    headers = ['转写规则编号', '转写规则名称', '案例编号', '原句／问题', '转写／核心辨析',
               '上下文／何时参考', '关联依据', '案例／原件位置']
    rows = [[rule, rules[rule][0], number, *records[number][:3], basis, records[number][3]]
            for (rule, number), basis in sorted(associations.items())]
    import os
    links = [os.path.relpath(ROOT / row[7], folder).replace('\\', '/') if row[7] else '' for row in rows]
    for row in rows:
        row[7] = '查看'
    save_xlsx(dict(headers=headers[:-1]+['查看正文／原件'], rows=rows, links=links, linkcol=7,
                   widths=[18,38,16,52,65,60,24,20], sheet='规则与案例索引'), folder/'规则与例句索引.xlsx')
    from build_query_page import build as build_page
    build_page()
    print(f'反向索引：{len(rules)}条转写规则／待审建议，{len(rows)}条案例关联。')


if __name__ == '__main__':
    build()

