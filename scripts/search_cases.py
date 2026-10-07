"""Search source examples, linked rules and human revision discussions."""
from pathlib import Path
import argparse,re,zipfile,xml.etree.ElementTree as E
from urllib.parse import unquote
P=Path(__file__).resolve().parents[1]
def read_sheets(path):
 import posixpath
 ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
 rid='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
 result={}
 with zipfile.ZipFile(path) as z:
  shared=[''.join(t.text or '' for t in item.findall('.//m:t',ns)) for item in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
  relationships={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
  for s in E.fromstring(z.read('xl/workbook.xml')).findall('m:sheets/m:sheet',ns):
   target=relationships[s.get(rid)]
   member=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
   xml=E.fromstring(z.read(member));rp=posixpath.dirname(member)+'/_rels/'+posixpath.basename(member)+'.rels'
   rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read(rp))} if rp in z.namelist() else {}
   links={h.get('ref'):rels.get(h.get(rid),'') for h in xml.findall('m:hyperlinks/m:hyperlink',ns)}
   headers={};records=[]
   for row in xml.findall('m:sheetData/m:row',ns):
    values={}
    for c in row.findall('m:c',ns):
     col=re.match('[A-Z]+',c.get('r')).group();v=c.findtext('m:v',default='',namespaces=ns)
     if c.get('t')=='s':v=shared[int(v)]
     elif c.get('t')=='inlineStr':v=''.join(t.text or '' for t in c.findall('.//m:t',ns))
     values[col]=v
    if row.get('r')=='1':headers=values;continue
    record={name:values.get(col,'') for col,name in headers.items()};record['_row']=int(row.get('r'))
    for col,name in headers.items():
     if col+row.get('r') in links:record[name]=str((path.parent/unquote(links[col+row.get('r')])).resolve())
    records.append(record)
   result[s.get('name')]=records
 return result

def linked(text,rule):
 return rule in re.split(r'[,，\s]+',text.strip())

def main():
 ap=argparse.ArgumentParser(description=__doc__)
 ap.add_argument('keyword',nargs='?',default='');ap.add_argument('--rule',default='');ap.add_argument('--source',default='');ap.add_argument('--limit',type=int,default=8);ap.add_argument('--context',action='store_true',help='Print the complete dialogue');ap.add_argument('--detail',action='store_true',help='Read the complete matching revision case')
 a=ap.parse_args()
 if not a.keyword and not a.rule:ap.error('provide a keyword or --rule')
 if a.rule and not re.fullmatch(r'R\d{2}|待审\d{2}|[\w]+-R\d{2}',a.rule):ap.error('use R04, 待审01, or a song rule such as 太阳-R01')
 if a.limit<1:ap.error('--limit must be positive')
 folder=P/'资料/规则与例句数据库';rules=folder/'公共转写规则.md';hits=[]
 if '-R' in a.rule:
  matches=[f for f in (P/'资料/讲义制作过程存档').glob('*/本曲转写规则.md') if re.search(r'^## '+re.escape(a.rule)+r' ',f.read_text(encoding='utf-8'),re.M) and (not a.source or a.source in str(f.parent))]
  if len(matches)!=1:ap.error('song rule not found or ambiguous; provide --source SONG_NAME')
  rules=matches[0]
 if a.rule and rules.exists():
  for section in re.split(r'(?=^## )',rules.read_text(encoding='utf-8'),flags=re.M):
   if section.startswith('## '+a.rule+' '):print(section.strip()+'\n')
 book=folder/'转写例句.xlsx'
 if book.exists():
  sheets=read_sheets(book);dialogues={}
  if a.rule:
   from build_rule_index import collect
   _, _, associations = collect(sheets)
   associated = {number for rule, number in associations if rule == a.rule}
  for row in sheets.get('会话上下文',[]):dialogues.setdefault(row.get('编号',row.get('会话编号','')),[]).append(f"第{row['轮次']}轮：{row['原文（含转写）']}")
  for row in sheets.get('转写例句',[]):
   source=row.get('来源文件','');category='教材' if '教程' in source else '小班' if '小班' in source else '手语歌讲义'
   if a.source and a.source.casefold() not in (source+' '+category).casefold():continue
   if a.rule and row.get('编号') not in associated:continue
   if a.keyword.casefold() not in '\n'.join(str(v) for v in row.values()).casefold():continue
   body='\n'.join(f'{k}：{v}' for k,v in row.items() if not k.startswith('_') and v)
   if a.context:
    match=re.search(r'会话编号：(D\d+)',row.get('上下文',''))
    if match:body+='\n完整会话：\n'+'\n'.join(dialogues.get(match.group(1),[]))
   hits.append(f'【{book.relative_to(P)} 第{row["_row"]}行】\n{body}')
 else:
  print('未找到转写例句.xlsx。')
  from build_rule_index import collect
  _, _, associations = collect({'转写例句': []})
  associated = {number for rule, number in associations if rule == a.rule}
 example_hits = hits
 hits = []
 from build_case_catalog import read_case
 for revisions in sorted((P/'资料/讲义制作过程存档').glob('*/转写修改与辨析案例/*.md')):
  if not re.match(r'.+-M\d+ ', revisions.name):continue
  case=read_case(revisions)
  if a.source and a.source not in ('修改','讲义制作过程存档') and a.source.casefold() not in case['手语歌'].casefold():continue
  if a.rule and case.get('编号') not in associated:continue
  introduction='\n'.join(str(case.get(k,'')) for k in ('编号','title','何时参考','核心辨析','关联规则','手语歌'))
  if a.keyword.casefold() not in (case['_text'] if a.detail else introduction).casefold():continue
  body=case['_text'] if a.detail else '\n'.join(f'{k}：{case.get(k,"")}' for k in ('编号','何时参考','核心辨析','关联规则','手语歌'))+f'\n正文：{revisions}\n用 --detail 阅读完整修改经过。'
  hits.append(f'【{revisions.relative_to(P)}】\n{body.strip()}')
 case_hits = hits
 def relevance(hit):
  return hit.casefold().count(a.keyword.casefold()) if a.keyword else 0
 example_hits.sort(key=relevance, reverse=True)
 case_hits.sort(key=relevance, reverse=True)
 hits = []
 for index in range(max(len(example_hits), len(case_hits))):
  for group in (example_hits, case_hits):
   if index < len(group):
    hits.append(group[index])
 for hit in hits[:a.limit]:print(hit+'\n')
 print(f'找到{len(hits)}处，显示{min(len(hits),a.limit)}处。')
 if a.rule:print('关联包含支持例、不同用法和边界案例；未关联条目仍可用关键词搜索。')

if __name__=='__main__':main()
