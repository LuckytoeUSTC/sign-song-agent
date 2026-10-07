"""Read-only practical Word checks; prints findings, not a stored data report."""
from pathlib import Path
from collections import Counter
import argparse,re
from docx import Document
from docx.enum.text import WD_LINE_SPACING
def effective(r,p,attribute):
 direct=getattr(r.font,attribute)
 if direct is not None:return direct
 for style in [r.style,p.style]:
  while style:
   value=getattr(style.font,attribute)
   if value is not None:return value
   style=style.base_style
 return None
def audit(path,height=None):
 d=Document(path);warnings=[];fonts=Counter();sizes=Counter();transcriptions=[];images=0
 width=min(s.page_width-s.left_margin-s.right_margin for s in d.sections)
 paragraphs=list(d.paragraphs)+[p for t in d.tables for row in t.rows for c in row.cells for p in c.paragraphs]
 for i,p in enumerate(paragraphs,1):
  text=p.text
  if not text.startswith('【') and '（' in text and '/' in text:transcriptions.append(text.split('（',1)[1].rsplit('）',1)[0])
  flags=[]
  for r in p.runs:
   flags.extend([bool(effective(r,p,'italic'))]*len(r.text))
   if r.text:
    fonts[effective(r,p,'name') or '未显式指定']+=len(r.text)
    size=effective(r,p,'size');sizes[round(size.pt,1) if size else '未显式指定']+=len(r.text)
  for m in re.finditer(r'\+\+',text):
   if any(flags[m.start():m.end()]):warnings.append(f'段落{i}：++含斜体')
  if '**' in text or '\\n' in text:warnings.append(f'段落{i}：可能残留Markdown或转义字符')
  for dr in p._p.xpath('.//w:drawing'):
   images+=1;ext=dr.xpath('.//wp:extent')
   if dr.xpath('./wp:anchor'):warnings.append(f'段落{i}：浮动图片，检查位置')
   if not ext:continue
   h=int(ext[0].get('cy'));w=int(ext[0].get('cx'))
   if height and abs(h/914400-height)>.002:warnings.append(f'段落{i}：图高{h/914400:.3f}英寸，检查是否为照片等例外')
   if w>width:warnings.append(f'段落{i}：单图宽度超过正文宽度')
   if p.paragraph_format.line_spacing_rule==WD_LINE_SPACING.EXACTLY:
    line=p.paragraph_format.line_spacing
    if line and line<h:warnings.append(f'段落{i}：固定行距小于图高，可能截框')
 print(path,'\n转写',len(transcriptions),'处；图片',images,'幅\n字体分布',dict(fonts),'\n字号分布',dict(sizes))
 print('\n'.join(warnings) if warnings else '结构检查未发现上述问题；仍需逐页预览。')
 return transcriptions
def main():
 ap=argparse.ArgumentParser();ap.add_argument('files',nargs='+',type=Path);ap.add_argument('--height',type=float);a=ap.parse_args()
 values=[audit(p,a.height) for p in a.files]
 if len(values)==2:
  if values[0]==values[1]:print('两版正文转写文字一致。')
  else:
   print('两版正文转写有差异：')
   from itertools import zip_longest
   for i,(x,y) in enumerate(zip_longest(*values),1):
    if x!=y:print(i,repr(x),'<>',repr(y))
if __name__=='__main__':main()
