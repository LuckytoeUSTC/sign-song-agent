"""Build a standalone offline query page from the current teaching records."""
import json
import os
from pathlib import Path
from urllib.parse import quote
from build_rule_index import collect
from build_case_catalog import ROOT
from search_cases import read_sheets


def build():
    records, rules, associations = collect()
    folder = ROOT / '资料'
    sheets = read_sheets(folder / '规则与例句数据库/转写例句.xlsx')
    examples = {r['编号']: r for r in sheets['转写例句']}
    dialogues = {}
    for r in sheets.get('会话上下文', []):
        dialogues.setdefault(r['编号'], []).append(f"第{r['轮次']}轮：{r['原文（含转写）']}")
    def link(source):
        return quote(os.path.relpath(ROOT / source, folder).replace('\\', '/'), safe='/') if source else ''
    entries = []
    for number, (sentence, expression, context, source, _) in records.items():
        example = examples.get(number)
        location = example.get('原文位置', '') if example else ''
        category = '修改案例' if not example else '教材' if '教程' in example.get('来源文件', '') else '小班讲义' if '小班' in example.get('来源文件', '') else '优秀讲义'
        full = '' if example else (ROOT / source).read_text(encoding='utf-8-sig')
        if example:
            import re
            match = re.search(r'会话编号：(D\d+)', context)
            if match:
                full = '\n\n'.join(dialogues.get(match.group(1), []))
        entries.append(dict(id=number, title=sentence, expression=expression, context=context,
                            source=category, location=location, link=link(source), full=full,
                            rules=[dict(id=r, basis=b) for (r, c), b in associations.items() if c == number]))
    data = dict(entries=entries, rules=[dict(id=r, title=v[0], body=v[1], link=link(v[2])) for r, v in rules.items()])
    payload = json.dumps(data, ensure_ascii=False).replace('<', '\\u003c')
    template = Path(__file__).with_name('query_template.html').read_text(encoding='utf-8')
    target = folder / '规则与例句查询.html'
    target.write_text(template.replace('__QUERY_DATA__', payload), encoding='utf-8')
    print(f'查询页面：{len(entries)}条记录，{len(rules)}条转写规则。')


if __name__ == '__main__':
    build()
