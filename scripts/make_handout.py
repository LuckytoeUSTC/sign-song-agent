"""One handout entry point: legacy Markdown or explicitly enabled content JSON."""
from pathlib import Path
import argparse
import copy
import json
import math
import re

from docx import Document
from docx.image.image import Image
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Re-export takeover APIs while keeping the existing single command entry.
from handout_update import (TOOL_VERSION, UPDATE_BEGIN, UPDATE_END,
                            compatibility_check, update_record_path,
                            update_fingerprint, write_record, update_takeover)
DEFAULT_LAYOUT = dict(font_name='SimSun', font_size=11, title_size=16,
                      heading_size=13, image_height=.79, margin_horizontal=2,
                      margin_vertical=1.8, line_spacing=1.15)
IMAGE_PATTERN = re.compile(r'!\[([^\]]*)\]\((.*?)\)')


def text_runs(p, text, font_name='SimSun', font_size=11):
    """Keep the legacy marker interpretation and run boundaries."""
    italic = strike = False
    index = 0
    while index < len(text):
        if text.startswith('~~', index):
            strike = not strike
            index += 2
            continue
        if text[index] == '*':
            italic = not italic
            index += 1
            continue
        if text.startswith('++', index):
            token = '++'
        elif text[index] == '~':
            token = '~'
        else:
            end = index + 1
            while end < len(text) and text[end] not in '*~' and not text.startswith('++', end):
                end += 1
            token = text[index:end]
        run = p.add_run(token)
        run.italic = italic and token != '++'
        run.font.strike = strike
        if token == '++':
            ics = OxmlElement('w:iCs')
            ics.set(qn('w:val'), '0')
            run._element.get_or_add_rPr().append(ics)
        run.font.name = 'Arial' if token == '~' else font_name
        run.font.size = Pt(font_size)
        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font_name)
        index += len(token)
    if italic or strike:
        raise ValueError('未闭合的斜体或删除线标记：' + text)


def read_layout(path=None, height=None):
    settings = DEFAULT_LAYOUT.copy()
    settings['reference_frame_padding'] = False
    if path:
        supplied = json.loads(Path(path).read_text(encoding='utf-8'))
        unknown = set(supplied) - set(settings)
        if unknown:
            raise ValueError('未知版式参数：' + ', '.join(sorted(unknown)))
        settings.update(supplied)
    for key, value in settings.items():
        if key == 'reference_frame_padding':
            if not isinstance(value, bool):
                raise ValueError('reference_frame_padding须为true或false')
            continue
        if key != 'font_name' and (not isinstance(value, (int, float)) or value <= 0):
            raise ValueError('版式参数必须是正数：' + key)
    if height is not None:
        settings['image_height'] = height
    if not .1 <= settings['image_height'] <= 3:
        raise ValueError('图片高度应为0.1至3英寸')
    return settings


def read_content(source, input_format='markdown'):
    """Read data only; JSON is opt-in and never inferred from a filename."""
    source = Path(source).resolve()
    text = source.read_text(encoding='utf-8')
    if input_format == 'markdown':
        lines = text.splitlines()
        blocks = []
        for i, line in enumerate(lines):
            if not line.strip():
                continue
            heading = re.match(r'^#{1,3} ', line)
            if heading:
                level = len(line) - len(line.lstrip('#'))
                blocks.append(dict(heading=line[level + 1:], level=level))
                continue
            tokens = []
            cursor = 0
            for match in IMAGE_PATTERN.finditer(line):
                tokens.append(dict(text=line[cursor:match.start()]))
                spec = match[2].strip()
                reference = spec.endswith('"参考"')
                if reference:
                    spec = spec[:-len('"参考"')].strip()
                tokens.append(dict(path=spec.strip('<>'), alt=match[1], reference=reference))
                tokens.append(dict(text=' ', raw=True))
                cursor = match.end()
            tokens.append(dict(text=line[cursor:]))
            blocks.append(dict(tokens=tokens, keep_with_next=i + 1 < len(lines) and '![' in lines[i + 1]))
        return dict(input_format='markdown', base=source.parent, blocks=blocks)
    data = json.loads(text)
    if not isinstance(data, dict) or data.get('version') != 1 or not isinstance(data.get('blocks'), list):
        raise ValueError('制作JSON须含version: 1和blocks数组')
    # Resolve only previous IDs: this makes repeat cycles impossible.
    resolved = []
    by_id = {}
    for original in data['blocks']:
        if not isinstance(original, dict):
            raise ValueError('每个制作块须为对象')
        block = copy.deepcopy(original)
        if 'repeat_of' in block:
            target = block['repeat_of']
            if target not in by_id:
                raise ValueError('重复段须引用前面已定义的id：' + str(target))
            inherited = copy.deepcopy(by_id[target])
            for key in ('lyrics', 'transcription', 'images', 'heading', 'text', 'level'):
                if key in block:
                    raise ValueError('重复段不另存一份正文或图片；修改原块或取消repeat_of：' + key)
            inherited.pop('id', None)
            inherited.update(block)
            block = inherited
            if block.get('omit_images', False):
                block['images'] = []
        elif 'omit_images' in block:
            raise ValueError('omit_images仅用于repeat_of重复段')
        if 'id' in block:
            if not isinstance(block['id'], str) or not block['id'] or block['id'] in by_id:
                raise ValueError('制作块id须为唯一的非空字符串')
            by_id[block['id']] = copy.deepcopy(block)
        resolved.append(block)
    return dict(input_format='json', base=source.parent, blocks=resolved)


def check_image(item, base):
    """Check file/rectangle facts, not whether a hand action is natural."""
    if not isinstance(item, dict) or not isinstance(item.get('path'), str):
        raise ValueError('每张选图须有path；不能用词条名代替已选文件')
    path = (base / item['path']).resolve()
    if not path.is_file():
        raise FileNotFoundError(f'图片不存在：{path}')
    image = Image.from_file(str(path))
    crop = item.get('crop')
    if crop is not None:
        if not isinstance(crop, list) or len(crop) != 4 or any(
                not isinstance(v, (int, float)) or not math.isfinite(v) for v in crop):
            raise ValueError('crop须为原图像素坐标[left, top, right, bottom]')
        left, top, right, bottom = crop
        if not (0 <= left < right <= image.px_width and 0 <= top < bottom <= image.px_height):
            raise ValueError('裁图范围越界或为空：' + str(path))
    return path, image


def check_content(content):
    for block in content['blocks']:
        if content['input_format'] == 'markdown':
            for token in block.get('tokens', []):
                if 'path' in token:
                    check_image(token, content['base'])
            continue
        for key in ('lyrics', 'transcription', 'text', 'heading', 'note', 'record'):
            if key in block and not isinstance(block[key], str):
                raise ValueError(key + '须为文字')
        if 'heading' in block and any(k in block for k in ('lyrics', 'transcription', 'text', 'images', 'note')):
            raise ValueError('标题块不同时包含正文、图片或短注释')
        if 'text' in block and any(k in block for k in ('lyrics', 'transcription')):
            raise ValueError('正文text与歌词/转写二选一，避免并行数据')
        if 'heading' in block and block.get('level', 1) not in (1, 2, 3):
            raise ValueError('标题level须为1、2或3')
        if not any(key in block for key in ('heading', 'text', 'lyrics', 'transcription')):
            raise ValueError('制作块须有标题、正文或歌词转写')
        if not isinstance(block.get('images', []), list):
            raise ValueError('images须为数组')
        if 'omit_images' in block and not isinstance(block['omit_images'], bool):
            raise ValueError('omit_images须为true或false')
        for item in block.get('images', []):
            check_image(item, content['base'])
            for key in ('entry', 'sense', 'source', 'action', 'note', 'record', 'alt'):
                if key in item and not isinstance(item[key], str):
                    raise ValueError('图片' + key + '须为文字')
            if 'reference' in item and not isinstance(item['reference'], bool):
                raise ValueError('reference须为true或false')
    return content


def new_document(settings):
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(settings['margin_horizontal'])
    section.top_margin = section.bottom_margin = Cm(settings['margin_vertical'])
    for name in ['Normal', 'Title', 'Heading 1', 'Heading 2']:
        style = document.styles[name]
        style.font.name = settings['font_name']
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), settings['font_name'])
    for style in document.styles:
        for border in style.element.xpath('.//w:pBdr'):
            border.getparent().remove(border)
    document.styles['Normal'].font.size = Pt(settings['font_size'])
    return document


def add_heading(document, block, settings):
    level = block.get('level', 1)
    p = document.add_paragraph(block['heading'], style='Title' if level == 1 else 'Heading ' + str(level - 1))
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    for run in p.runs:
        run.font.size = Pt(settings['title_size'] if level == 1 else settings['heading_size'])
    return p


def add_paragraph(document, settings, images=False, keep_with_next=False):
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(7 if images else 5)
    p.paragraph_format.line_spacing = 1.25 if images else settings['line_spacing']
    p.paragraph_format.keep_together = True
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    if images:
        p.paragraph_format.space_before = Pt(4)
        if settings.get('reference_frame_padding', False):
            from docx.enum.text import WD_LINE_SPACING
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            p.paragraph_format.line_spacing = Pt(settings['image_height'] * 72 + 6)
    return p


def add_picture(p, item, base, settings):
    path, image = check_image(item, base)
    shape = p.add_run().add_picture(str(path), height=Inches(settings['image_height']))
    # Alt/entry metadata is never turned into a visible caption.
    shape._inline.docPr.set('descr', item.get('alt', item.get('entry', '')))
    if 'crop' in item:
        left, top, right, bottom = item['crop']
        shape.width = int(shape.height * (right - left) / (bottom - top))
        rect = OxmlElement('a:srcRect')
        for key, fraction in dict(l=left / image.px_width, t=top / image.px_height,
                                  r=1 - right / image.px_width, b=1 - bottom / image.px_height).items():
            rect.set(key, str(round(fraction * 100000)))
        fill = shape._inline.xpath('.//pic:blipFill')[0]
        stretch = fill.find(qn('a:stretch'))
        if stretch is None:
            fill.append(rect)
        else:
            fill.insert(list(fill).index(stretch), rect)
    if item.get('reference', False):
        for effect in shape._inline.xpath('./wp:effectExtent'):
            for key, value in dict(t='25400', b='25400', l='12700', r='12700').items():
                effect.set(key, value)
        props = shape._inline.xpath('.//pic:spPr')[0]
        line = OxmlElement('a:ln')
        line.set('w', '12700')
        if settings.get('reference_frame_padding', False):
            line.set('algn', 'in')
        fill = OxmlElement('a:solidFill')
        color = OxmlElement('a:srgbClr')
        color.set('val', '000000')
        fill.append(color)
        line.append(fill)
        dash = OxmlElement('a:prstDash')
        dash.set('val', 'sysDot')
        line.append(dash)
        props.append(line)
    return shape


def render_content(content, settings):
    document = new_document(settings)
    write = lambda p, value: text_runs(p, value, settings['font_name'], settings['font_size'])
    for block in content['blocks']:
        if 'heading' in block:
            add_heading(document, block, settings)
            continue
        if content['input_format'] == 'markdown':
            tokens = block['tokens']
            p = add_paragraph(document, settings, any('path' in t for t in tokens), block['keep_with_next'])
            for token in tokens:
                if 'path' in token:
                    add_picture(p, token, content['base'], settings)
                elif token.get('raw'):
                    p.add_run(token['text'])
                else:
                    write(p, token['text'])
            if '![' in p.text:
                raise ValueError('图片Markdown未识别：' + p.text)
            continue
        images = block.get('images', [])
        text = block.get('text', block.get('lyrics', ''))
        if 'transcription' in block:
            text += '（' + block['transcription'] + '）'
        p = add_paragraph(document, settings, keep_with_next=bool(images))
        write(p, text)
        if images:
            picture_row = add_paragraph(document, settings, images=True)
            for item in images:
                add_picture(picture_row, item, content['base'], settings)
                if item.get('note'):
                    write(picture_row, '〔' + item['note'] + '〕')
                write(picture_row, ' ')
        if block.get('note'):
            note_row = add_paragraph(document, settings)
            write(note_row, block['note'])
    return document


def make_handout(source, output, height=None, layout=None, input_format='markdown'):
    output = Path(output)
    if output.exists():
        raise ValueError('输出文件已存在；请选择新文件名')
    settings = read_layout(layout, height)
    content = check_content(read_content(source, input_format))
    document = render_content(content, settings)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return output


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--version', action='version', version=TOOL_VERSION)
    ap.add_argument('source', type=Path, nargs='?')
    ap.add_argument('--out', type=Path)
    ap.add_argument('--height', type=float)
    ap.add_argument('--layout', type=Path, help='JSON layout settings')
    ap.add_argument('--input-format', choices=('markdown', 'json'), default='markdown')
    ap.add_argument('--update-action', choices=('check', 'complete', 'fail'))
    ap.add_argument('--update-note', default='')
    args = ap.parse_args()
    if args.update_action:
        if args.source or args.out or args.layout or args.height is not None or args.input_format != 'markdown':
            ap.error('更新接管动作不能与讲义制作参数混用')
        print(json.dumps(update_takeover(args.update_action, note=args.update_note), ensure_ascii=False, indent=2))
    else:
        if not args.source or not args.out:
            ap.error('制作需要source与--out')
        make_handout(args.source, args.out, args.height, args.layout, args.input_format)
        print('生成', args.out, '；请逐页预览后使用。')


if __name__ == '__main__':
    main()
