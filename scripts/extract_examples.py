"""Extract located examples to Markdown; source files are read-only.
Usage: python scripts/extract_examples.py source --out new.md
Generated text needs source review before semantic analysis or teaching use.
"""
from pathlib import Path
import argparse,re

def spans(text,opens='（(',closes='）)'):
 depth=0;start=0
 for i,ch in enumerate(text):
  if ch in opens:
   if not depth:start=i
   depth+=1
  elif ch in closes and depth:
   depth-=1
   if not depth:yield start,i+1

def features(text):
 flags=[]
 for label,pat in [('方向','[→←＜＞]'),('重复','\\+\\+'),('否定',r'不|没有|不能|无|不愿意'),('指示','指[（(]'),('动作说明','[（(]'),('疑问','[?？]')]:
  if re.search(pat,text):flags.append(label)
 return '、'.join(flags) or '—'

def records(path):
 if path.suffix.lower()=='.md':
  lines=path.read_text(encoding='utf-8').splitlines();chapter='';section='';turn=0
  for i,line in enumerate(lines):
   if line.startswith('## '):chapter=line[3:]
   if line.startswith('### '):section=line[4:];turn=0
   if not ('句子' in section or '会话' in section):continue
   parts=list(spans(line,'｛','｝'))
   if not parts:continue
   turn+=1;first=parts[0][0]
   yield dict(original=line[:first].lstrip('> ').strip(),transcriptions=[line[a+1:b-1] for a,b in parts],
     location=f'{chapter}；{section}；第{i+1}行'+(f'；会话第{turn}轮' if '会话' in section else ''),
     topic=chapter+' / '+section,before=lines[i-1] if i else '',after=lines[i+1] if i+1<len(lines) else '',
     note='原教材整理文本；尚未逐句人工审核。')
  return
 units=[]
 if path.suffix.lower()=='.docx':
  from docx import Document
  for i,p in enumerate(Document(path).paragraphs,1):
   text=''.join(('~~'+r.text+'~~' if r.font.strike and r.text else r.text) for r in p.runs)
   units.append((f'正文第{i}段',text))
 elif path.suffix.lower()=='.pdf':
  import pdfplumber
  with pdfplumber.open(path) as pdf:
   for n,page in enumerate(pdf.pages,1):
    lines=(page.extract_text() or '').splitlines();pending='';start=0
    for i,line in enumerate(lines,1):
     if pending:pending+=' '+line
     else:pending=line;start=i
     depth=sum(ch in '（(' for ch in pending)-sum(ch in '）)' for ch in pending)
     if depth>0:continue
     units.append((f'第{n}页；提取文本第{start}—{i}行',pending));pending=''
    if pending:units.append((f'第{n}页；提取文本第{start}行起',pending))
 else:raise ValueError('支持Markdown教材、DOCX和PDF')
 for i,(loc,text) in enumerate(units):
  pairs=[(a,b) for a,b in spans(text) if '/' in text[a:b]]
  if not pairs:continue
  # Multiple pairs retain the complete original paragraph: no invented lyric alignment.
  first=pairs[0][0];cursor=0;lyric=[]
  for a,b in pairs:lyric.append(text[cursor:a]);cursor=b
  lyric.append(text[cursor:]);original=' '.join(''.join(lyric).split()).strip(' /')
  if not original:continue
  yield dict(original=original,transcriptions=[text[a+1:b-1] for a,b in pairs],location=loc,topic=path.stem,
    before=next((t.strip() for _,t in reversed(units[:i]) if t.strip()),''),
    after=next((t.strip() for _,t in units[i+1:] if t.strip()),''),
    note=('同一段含多个转写，原句对应须看完整段落。' if len(pairs)>1 else '')+
    ('PDF提取不保留斜体与删除线，须看原页。' if path.suffix.lower()=='.pdf' else '斜体与图示须看原件。'),context=text)

def markdown(path,items,out):
 import os
 link=Path(os.path.relpath(path,out.parent)).as_posix()
 result=[f'# {path.stem}：转写例子','',f'原件：[查看资料](<{link}>)。以下为AI定位整理，不是逐句教学解释。','']
 for i,r in enumerate(items,1):
  result += [f'## 例{i}　{r["original"]}','',f'位置：{r["location"]}',f'主题：{r["topic"]}',f'表达线索：{features(" ".join(r["transcriptions"]))}','',f'原句：{r["original"]}','']
  for j,t in enumerate(r['transcriptions'],1):result += [f'转写{j}：{t}','']
  if r.get('context'):result += [f'完整段落：{r["context"]}','']
  result += [f'上文：{r["before"] or "—"}',f'下文：{r["after"] or "—"}',f'说明：{r["note"]}','']
 return '\n'.join(result)

def main():
 ap=argparse.ArgumentParser();ap.add_argument('source',type=Path);ap.add_argument('--out',type=Path,required=True);args=ap.parse_args()
 if args.out.exists():raise FileExistsError('不覆盖已有整理，请先比较或选择新路径')
 items=list(records(args.source));args.out.parent.mkdir(parents=True,exist_ok=True);args.out.write_text(markdown(args.source,items,args.out),encoding='utf-8')
 print(f'{len(items)}组，{sum(len(x["transcriptions"]) for x in items)}个转写方案；{args.out}')
if __name__=='__main__':main()
