"""Build a new black-and-white handout from approved Markdown and selected pictures."""
from pathlib import Path
import argparse,re
from docx import Document
from docx.shared import Inches,Pt,RGBColor,Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def text_runs(p,text):
 italic=False;strike=False;index=0
 while index<len(text):
  if text.startswith('~~',index):strike=not strike;index+=2;continue
  if text[index]=='*':italic=not italic;index+=1;continue
  token='++' if text.startswith('++',index) else text[index]
  r=p.add_run(token);r.italic=italic and token!='++';r.font.strike=strike
  if token=='++':
   ics=OxmlElement('w:iCs');ics.set(qn('w:val'),'0');r._element.get_or_add_rPr().append(ics)
  r.font.name='Arial' if token=='~' else 'SimSun';r.font.size=Pt(11)
  r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'SimSun')
  index+=len(token)
 if italic or strike:raise ValueError('未闭合的斜体或删除线标记：'+text)

def main():
 ap=argparse.ArgumentParser();ap.add_argument('source',type=Path);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--height',type=float,default=.79);a=ap.parse_args()
 if a.out.exists():raise ValueError('输出文件已存在；请选择新文件名')
 if not .1<=a.height<=3:raise ValueError('图片高度应为0.1至3英寸')
 source=a.source.resolve();lines=source.read_text(encoding='utf-8').splitlines();d=Document()
 section=d.sections[0];section.page_width=Cm(21);section.page_height=Cm(29.7)
 section.left_margin=section.right_margin=Cm(2);section.top_margin=section.bottom_margin=Cm(1.8)
 for name in ['Normal','Title','Heading 1','Heading 2']:
  style=d.styles[name];style.font.name='SimSun';style.font.color.rgb=RGBColor(0,0,0)
  style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'SimSun')
 for style in d.styles:
  for border in style.element.xpath('.//w:pBdr'):border.getparent().remove(border)
 d.styles['Normal'].font.size=Pt(11)
 image=re.compile(r'!\[([^\]]*)\]\((.*?)\)')
 for i,line in enumerate(lines):
  if not line.strip():continue
  if line.startswith('# '):
   p=d.add_paragraph(line[2:],style='Title');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
   p.paragraph_format.space_after=Pt(10)
   for r in p.runs:r.font.size=Pt(16)
   continue
  p=d.add_paragraph();p.paragraph_format.space_after=Pt(5);p.paragraph_format.line_spacing=1.15
  p.paragraph_format.keep_together=True
  if i+1<len(lines) and '![' in lines[i+1]:p.paragraph_format.keep_with_next=True
  matches=list(image.finditer(line));cursor=0
  if matches:
   p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(7)
   p.paragraph_format.line_spacing=1.25
  for m in matches:
   text_runs(p,line[cursor:m.start()]);spec=m[2].strip();reference=spec.endswith('"参考"')
   if reference:spec=spec[:-len('"参考"')].strip()
   path=(source.parent/spec.strip('<>')).resolve()
   if not path.is_file():raise FileNotFoundError(f'图片不存在：{path}')
   shape=p.add_run().add_picture(str(path),height=Inches(a.height))
   shape._inline.docPr.set('descr',m[1])
   if reference:
    for effect in shape._inline.xpath('./wp:effectExtent'):
     effect.set('t','25400');effect.set('b','25400');effect.set('l','12700');effect.set('r','12700')
    props=shape._inline.xpath('.//pic:spPr')[0];ln=OxmlElement('a:ln');ln.set('w','12700')
    fill=OxmlElement('a:solidFill');color=OxmlElement('a:srgbClr');color.set('val','000000');fill.append(color);ln.append(fill)
    dash=OxmlElement('a:prstDash');dash.set('val','sysDot');ln.append(dash);props.append(ln)
   p.add_run(' ');cursor=m.end()
  text_runs(p,line[cursor:])
  if '![' in p.text:raise ValueError('图片Markdown未识别：'+line)
 a.out.parent.mkdir(parents=True,exist_ok=True);d.save(a.out)
 print('生成',a.out,'；请逐页预览后使用。')
if __name__=='__main__':main()
