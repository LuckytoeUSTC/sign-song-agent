"""Search canonical XLSX examples and Markdown revision cases."""
from pathlib import Path
import argparse,re,json,zipfile,xml.etree.ElementTree as E
P=Path(__file__).resolve().parents[1]
def main():
 ap=argparse.ArgumentParser();ap.add_argument('keyword');ap.add_argument('--source',default='');ap.add_argument('--limit',type=int,default=8);a=ap.parse_args();count=0;shown=0
 book=P/'资料/提取与检索/原句与手语转写.xlsx';ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
 with zipfile.ZipFile(book) as z:
  shared=[]
  rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/worksheets/_rels/sheet1.xml.rels'))}
  sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
  hyperlinks={h.get('ref'):rels[h.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] for h in sheet.findall('.//m:hyperlink',ns)}
  if 'xl/sharedStrings.xml' in z.namelist():shared=[''.join(si.itertext()) for si in E.fromstring(z.read('xl/sharedStrings.xml'))]
  for row in sheet.findall('.//m:sheetData/m:row',ns):
   number=int(row.get('r'))
   if number<2:continue
   values=['']*11
   for cell in row.findall('m:c',ns):
    letters=re.match(r'[A-Z]+',cell.get('r')).group();col=0
    for letter in letters:col=col*26+ord(letter)-64
    v=cell.find('m:v',ns);value=v.text if v is not None else ''
    if cell.get('t')=='s':value=shared[int(value)]
    elif cell.get('t')=='inlineStr':value=''.join(t.text or '' for t in cell.findall('.//m:t',ns))
    values[col-1]=value or ''
   values[10]=str((book.parent/hyperlinks.get('K'+str(number),values[10])).resolve())
   source_text=values[0]+' '+values[4]+' '+('教材' if '教程' in values[0] else '优秀讲义')
   if a.keyword.casefold() in '\n'.join(values).casefold() and (not a.source or a.source.casefold() in source_text.casefold()):
    count+=1
    if shown<a.limit:
     print(f'【{book.relative_to(P)} 第{number}行】\n'+ '\n'.join(f'{key}：{value}' for key,value in zip(['来源文件','作者','原句','手语转写','主题','上文','下文','完整段落','原文位置','说明','原件链接'],values))+'\n');shown+=1
 for path in [P/'资料/提取与检索/转写修改与辨析.md']:
  text=path.read_text(encoding='utf-8')
  for section in re.split(r'(?=^## )',text,flags=re.M):
   if a.keyword.casefold() in section.casefold() and (not a.source or a.source.casefold() in str(path).casefold()):
    count+=1
    if shown<a.limit:print(f'【{path.relative_to(P)}】\n{section.strip()}\n');shown+=1
 print(f'找到{count}处，显示{shown}处。可用--source筛选来源，--limit调整显示数量。')
if __name__=='__main__':main()
