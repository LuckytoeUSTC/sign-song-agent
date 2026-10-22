"""Portable host for the existing offline query interface."""
from pathlib import Path
from urllib.parse import unquote, urlsplit
import ctypes
import json
import os
import re
import sys
import subprocess

def find_root():
    if not getattr(sys, 'frozen', False):
        return Path(__file__).resolve().parents[2]
    here = Path(sys.executable).resolve().parent
    if sys.platform == 'darwin':
        for parent in here.parents:
            if parent.suffix == '.app':
                return parent.parent
    return here

ROOT = find_root()
PAGE = ROOT / 'scripts' / 'query_page.html'
if not PAGE.is_file() and getattr(sys, 'frozen', False):
    PAGE = Path(sys._MEIPASS) / 'query_page.html'



def open_document(path):
    if sys.platform == 'win32':
        os.startfile(str(path))
    else:
        env = os.environ.copy()
        if sys.platform.startswith('linux'):
            original = env.pop('LD_LIBRARY_PATH_ORIG', None)
            if original is None:
                env.pop('LD_LIBRARY_PATH', None)
            else:
                env['LD_LIBRARY_PATH'] = original
        subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(path)], env=env)


def reader_html(path):
    text = path.read_text(encoding='utf-8-sig')
    if path.suffix == '.html':
        return text
    import html as htmlmod
    import markdown
    from urllib.parse import quote
    def rewrite(match):
        label, value = match.groups()
        value = htmlmod.unescape(value.strip('<>'))
        parts = urlsplit(value)
        if parts.scheme or parts.netloc or not parts.path:
            return match.group()
        target = (path.parent / unquote(parts.path)).resolve()
        if not target.is_relative_to(ROOT / '资料') or not target.is_file():
            return label.strip('[]')
        link = quote(str(target.relative_to(ROOT / '资料')).replace('\\', '/'), safe='/')
        return f'{label}(<{link}>)'
    text = re.sub(r'(!?\[[^\]]*\])\((<[^>]*>|[^)]+)\)', rewrite, text)
    body = markdown.markdown(text, extensions=['tables', 'fenced_code', 'sane_lists'])
    body = re.sub(r'href="([^"]*)"', lambda m: m.group() + (' data-source="' + m.group(1) + '"' if not urlsplit(htmlmod.unescape(m.group(1))).scheme and not m.group(1).startswith('#') else ''), body)
    style = 'body{font-family:sans-serif;color:#26372e;background:#f7f8f5;max-width:960px;margin:32px auto;padding:0 24px;line-height:1.8}a{color:#35654b}td,th{border:1px solid #ccc;padding:8px}img{max-width:100%}'
    script = "document.addEventListener('click',async e=>{let a=e.target.closest('a[data-source]');if(!a||!window.pywebview?.api)return;e.preventDefault();let r=await window.pywebview.api.open_source(a.dataset.source);if(!r.ok)alert(r.error);});"
    return '<!doctype html><html lang="zh-CN"><meta charset="utf-8"><style>' + style + '</style><body>' + body + '<script>' + script + '</script></body></html>'


class FileApi:
    def __init__(self, html, opener=None):
        self.data = json.loads(re.search(r'<script id="data" type="application/json">(.*?)</script>', html, re.S).group(1))
        self._opener = opener
        self._allowed = {p.resolve() for p in (ROOT / '资料').rglob('*') if p.is_file() and p.suffix.lower() in {'.html', '.md', '.docx', '.doc', '.pdf'}}

    def reader_ready(self, length):
        if isinstance(length, int):
            self._reader_length = length
        return True

    def open_source(self, link):
        if not isinstance(link, str) or urlsplit(link).scheme or urlsplit(link).netloc:
            return {'ok': False, 'error': '只能打开随包资料。'}
        path = (ROOT / '资料' / unquote(urlsplit(link).path)).resolve()
        if path not in self._allowed or not path.is_relative_to(ROOT / '资料'):
            return {'ok': False, 'error': '原件不存在，或链接不在随包资料范围内。'}
        try:
            if self._opener:
                self._opener(str(path))
            elif path.suffix in {'.html', '.md'}:
                import webview
                webview.create_window(path.stem, html=reader_html(path), js_api=self,
                                      width=1000, height=780, min_size=(640, 480), text_select=True,
                                      background_color='#F7F8F5')
            else:
                # External document apps should not inherit PyInstaller DLL lookup paths.
                frozen = getattr(sys, 'frozen', False) and sys.platform == 'win32'
                if frozen:
                    ctypes.windll.kernel32.SetDllDirectoryW(None)
                try:
                    open_document(path)
                finally:
                    if frozen:
                        ctypes.windll.kernel32.SetDllDirectoryW(str(sys._MEIPASS))
            return {'ok': True}
        except Exception as error:
            return {'ok': False, 'error': '无法打开原件，请检查相应阅读软件：' + str(error)}


def check(html):
    opened = []
    api = FileApi(html, opened.append)
    for entry in api.data['entries'] + api.data['rules']:
        if entry.get('link'):
            result = api.open_source(entry['link'])
            if not result['ok']:
                raise RuntimeError(entry['link'] + ': ' + result['error'])
    for link in ['../使用说明.txt', 'https://example.com', '../规则与例句查询.exe', '不存在.pdf']:
        assert not api.open_source(link)['ok']
    return {'frozen': bool(getattr(sys, 'frozen', False)), 'root': str(ROOT),
            'entries': len(api.data['entries']), 'rules': len(api.data['rules']),
            'checked_links': len(opened), 'python': sys.version.split()[0]}


def main():
    if not PAGE.is_file():
        raise RuntimeError('缺少scripts\\query_page.html，请完整解压，并保留原目录。')
    html = PAGE.read_text(encoding='utf-8')
    if len(sys.argv) > 2 and sys.argv[1] == '--check':
        Path(sys.argv[2]).write_text(json.dumps(check(html), ensure_ascii=False, indent=2), encoding='utf-8')
        return
    import webview
    api = FileApi(html)
    # Qt setHtml has a 2 MB encoded-data limit; load the bundled local file instead.
    page_content = {'url': PAGE.as_uri()} if sys.platform.startswith('linux') else {'html': html}
    window = webview.create_window(('手语歌智能工坊 · 规则与例句查询' if (ROOT / 'AGENTS.md').exists() else 'NO-AI-ZONE · 规则与例句查询'), **page_content, js_api=api,
                                  width=1120, height=800, min_size=(640, 480), text_select=True,
                                  background_color='#F5F6F3')
    if len(sys.argv) > 2 and sys.argv[1] == '--smoke':
        report_path = Path(sys.argv[2])
        def smoke():
            import time
            report = check(html)
            try:
                report['interface'] = window.evaluate_js('''(()=>{
                    const initial=document.getElementById('count').textContent;
                    document.getElementById('keyword').value='猫';render();
                    const search=filtered.length;
                    const highlighted=document.querySelectorAll('mark').length;
                    document.getElementById('keyword').value='';
                    document.getElementById('rule').value='R01';render();
                    const rule=filtered.length;
                    const ruleBody=!!document.querySelector('#rulebody article');
                    document.getElementById('rule').value='';
                    document.getElementById('source').value='教材';render();
                    const source=filtered.length;
                    document.getElementById('source').value='';
                    document.getElementById('keyword').value='zzzzzz不存在';render();
                    const empty=!!document.querySelector('.empty');
                    document.getElementById('keyword').value='';render();
                    document.getElementById('more').click();
                    return {initial,search,highlighted,rule,ruleBody,source,empty,cards:document.querySelectorAll('.card').length};
                })()''')
                v = report['interface']
                assert v['initial'] == str(report['entries']) + ' 条'
                assert v['search'] > 0 and v['highlighted'] > 0 and v['rule'] > 0
                assert v['ruleBody'] and v['source'] > 0 and v['empty'] and v['cards'] == 60
                link = next(x['link'] for x in api.data['entries'] if x.get('link', '').endswith(('.html', '.md')))
                assert api.open_source(link)['ok']
                time.sleep(1)
                child = [w for w in webview.windows if w != window][-1]
                api._reader_length = 0
                child.run_js("(()=>{let tries=0;function probe(){if(window.pywebview?.api){window.pywebview.api.reader_ready(document.body.innerText.length)}else if(++tries<20){setTimeout(probe,200)}}probe()})()")
                for _ in range(50):
                    if api._reader_length > 100:
                        break
                    time.sleep(0.1)
                report['reader_chars'] = api._reader_length
                assert report['reader_chars'] > 100
                report['ok'] = True
                time.sleep(5)
            except Exception as error:
                report['ok'] = False
                report['error'] = repr(error)
            finally:
                report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
                for w in list(webview.windows):
                    w.destroy()
        window.events.loaded += smoke
    storage = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / '.cache'))) / 'NO-AI-ZONE' / 'WebView2'
    engine = 'edgechromium' if sys.platform == 'win32' else 'cocoa' if sys.platform == 'darwin' else 'qt'
    webview.start(gui=engine, private_mode=False, storage_path=str(storage))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        message = '无法启动查询程序：\n' + str(error) + '\n请确认已完整解压，并查看README中的系统要求。'
        if sys.platform == 'win32':
            ctypes.windll.user32.MessageBoxW(None, message, '规则与例句查询', 16)
        else:
            print(message, file=sys.stderr)
            try:
                if sys.platform == 'darwin':
                    subprocess.run(['osascript', '-e', 'display alert "规则与例句查询" message ' + json.dumps(message, ensure_ascii=False)], check=False)
            except OSError:
                pass
        sys.exit(1)
