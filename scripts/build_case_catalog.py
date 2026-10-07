"""Generate the revision catalogue from the short introductions in case Markdown."""
from pathlib import Path
import argparse
import json
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
HEADERS = ['编号', '原句／问题', '何时参考', '核心辨析', '关联规则', '手语歌', '查看正文']
WIDTHS = [14, 42, 58, 58, 24, 25, 12]
M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def read_case(path):
    text = path.read_text(encoding='utf-8')
    heading = re.search(r'^# (.+)$', text, re.M)
    fields = dict(re.findall(r'^\*\*([^*]+)\*\*：([^\n]*)$', text, re.M))
    return {'title': heading.group(1), **fields, '_path': path, '_text': text}


def catalogue():
    folder = ROOT / '资料/规则与例句数据库'
    cases = [read_case(p) for p in (ROOT / '资料/讲义制作过程存档').glob('*/转写修改与辨析案例/*.md') if re.match(r'.+-M\d+ ', p.name)]
    cases.sort(key=lambda c: c['编号'])
    groups = {}
    for case in cases:
        groups.setdefault(case['_path'].parent, []).append(case)
    for directory, entries in groups.items():
        lines = ['# 转写修改与辨析案例目录', '', '编号：歌曲简称＋M＋顺序号，例如太阳-M001；编号不随排序改变。', '', '| 编号 | 原句／问题 |', '|---|---|']
        lines += [f"| {c['编号']} | [{c['title']}](<{c['_path'].name}>) |" for c in entries]
        lines += ['', '先读案例的“何时参考”和“核心辨析”，再读完整修改经过。', '']
        (directory / '.content').write_text('\n'.join(lines), encoding='utf-8')
    numbers = [c['编号'] for c in cases]
    if len(numbers) != len(set(numbers)):
        raise ValueError('修改案例有重复编号，请为新案例使用未占用的编号。')
    rows = [[c['编号'], c['title'], c['何时参考'], c['核心辨析'], c.get('关联规则', ''), c['手语歌'], '查看'] for c in cases]
    import os
    links = [os.path.relpath(c['_path'], folder).replace('\\', '/') for c in cases]
    return dict(filename='转写修改与辨析目录.xlsx', sheet='转写修改与辨析目录', headers=HEADERS,
                rows=rows, links=links, linkcol=6, widths=WIDTHS)


def xml(tag, attrs=None, parent=None):
    return ET.SubElement(parent, '{'+M+'}'+tag, attrs or {}) if parent is not None else ET.Element('{'+M+'}'+tag, attrs or {})


def save_xlsx(data, output):
    """Portable exporter for this text-only index; no extra Python package needed."""
    headers = data['headers']; widths = data['widths']; linkcol = data.get('linkcol', 6)
    lastcol = chr(64 + len(headers))
    ET.register_namespace('', M)
    ET.register_namespace('r', R)
    sheet = xml('worksheet')
    views = xml('sheetViews', parent=sheet)
    view = xml('sheetView', {'workbookViewId': '0'}, views)
    xml('pane', {'xSplit': '1', 'ySplit': '1', 'topLeftCell': 'B2', 'activePane': 'bottomRight', 'state': 'frozen'}, view)
    cols = xml('cols', parent=sheet)
    for n, width in enumerate(widths, 1):
        xml('col', {'min': str(n), 'max': str(n), 'width': str(width), 'customWidth': '1'}, cols)
    records = xml('sheetData', parent=sheet)
    for n, values in enumerate([headers, *data['rows']], 1):
        lengths = [sum(max(1, (sum(2 if unicodedata.east_asian_width(c) in 'WF' else 1 for c in line)+w-1)//w) for line in str(v).splitlines()) for v,w in zip(values,widths)]
        row = xml('row', {'r': str(n), 'ht': str(32 if n==1 else max(42, max(lengths,default=1)*16+10)), 'customHeight': '1'}, records)
        for c, value in enumerate(values):
            cell = xml('c', {'r': chr(65+c)+str(n), 't': 'inlineStr', 's': '1' if n==1 else '2' if n%2==0 else '0'}, row)
            inline = xml('is', parent=cell)
            xml('t', parent=inline).text = str(value or '')
    hyperlinks = xml('hyperlinks', parent=sheet)
    relations = ET.Element('Relationships', {'xmlns': PKG})
    for n, target in enumerate(data['links'], 2):
        xml('hyperlink', {'ref': chr(65+linkcol)+str(n), '{'+R+'}id': 'link'+str(n)}, hyperlinks)
        ET.SubElement(relations, 'Relationship', {'Id':'link'+str(n),'Type':R+'/hyperlink','Target':target,'TargetMode':'External'})
    parts=xml('tableParts',{'count':'1'},sheet);xml('tablePart',{'{'+R+'}id':'table1'},parts)
    ET.SubElement(relations,'Relationship',{'Id':'table1','Type':R+'/table','Target':'../tables/table1.xml'})
    extent=f'A1:{lastcol}{len(data["rows"])+1}'
    table=xml('table',{'id':'1','name':'RevisionCases','displayName':'RevisionCases','ref':extent,'totalsRowShown':'0'})
    xml('autoFilter',{'ref':extent},table)
    columns=xml('tableColumns',{'count':str(len(headers))},table)
    for n,name in enumerate(headers,1):xml('tableColumn',{'id':str(n),'name':name},columns)
    xml('tableStyleInfo',{'name':'TableStyleLight1','showFirstColumn':'0','showLastColumn':'0','showRowStripes':'0','showColumnStripes':'0'},table)
    workbook = xml('workbook');sheets = xml('sheets', parent=workbook)
    xml('sheet', {'name':data['sheet'],'sheetId':'1','{'+R+'}id':'sheet1'}, sheets)
    styles=f'''<styleSheet xmlns="{M}"><fonts count="2"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="11"/><name val="等线"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    types='''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'''
    packages={'[Content_Types].xml':types,'_rels/.rels':f'<Relationships xmlns="{PKG}"><Relationship Id="root" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'xl/workbook.xml':ET.tostring(workbook,encoding='utf-8'),'xl/styles.xml':styles,
      'xl/_rels/workbook.xml.rels':f'<Relationships xmlns="{PKG}"><Relationship Id="sheet1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml':ET.tostring(sheet,encoding='utf-8'), 'xl/worksheets/_rels/sheet1.xml.rels':ET.tostring(relations,encoding='utf-8'), 'xl/tables/table1.xml':ET.tostring(table,encoding='utf-8')}
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
        for name,body in packages.items():archive.writestr(name,body)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out',type=Path,default=ROOT/'资料/规则与例句数据库/转写修改与辨析目录.xlsx')
    parser.add_argument('--json',type=Path,help='Export data for another spreadsheet renderer')
    args=parser.parse_args();data=catalogue()
    if args.json:args.json.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
    else:save_xlsx(data,args.out)
    print(f'目录包含{len(data["rows"])}个案例。正文及引导文字只在各案例Markdown中维护。')


if __name__=='__main__':main()
