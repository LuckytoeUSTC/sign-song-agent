"""Desktop host for the offline query page; no model or remote service."""
import json
import os
import sys
import subprocess
from pathlib import Path
import re
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parents[1]
PAGE = ROOT / 'scripts/query_page.html'


def open_file(path):
    if sys.platform == 'win32':
        os.startfile(path)
    else:
        subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', path])


class FileApi:
    def __init__(self, html, opener=None):
        data = json.loads(re.search(r'<script id="data" type="application/json">(.*?)</script>', html, re.S).group(1))
        self._allowed = {entry['link'] for entry in data['entries'] + data['rules'] if entry.get('link')}
        self._opener = opener or open_file

    def open_source(self, link):
        if not isinstance(link, str) or link not in self._allowed:
            return {'ok': False, 'error': '只能打开查询资料中的原件。'}
        path = (ROOT / '资料' / unquote(link)).resolve()
        if not path.is_relative_to(ROOT / '资料') or path.suffix.lower() not in {'.docx', '.doc', '.pdf', '.md'}:
            return {'ok': False, 'error': '原件路径不在资料范围内。'}
        if not path.is_file():
            return {'ok': False, 'error': '原件不存在，请更新查询资料。'}
        try:
            self._opener(str(path))
            return {'ok': True}
        except OSError as error:
            return {'ok': False, 'error': f'无法打开原件：{error}'}


def main():
    import webview
    if not PAGE.exists():
        from build_query_page import build
        build()
    html = PAGE.read_text(encoding='utf-8')
    page_content = {'url': PAGE.as_uri()} if sys.platform.startswith('linux') else {'html': html}
    webview.create_window('规则与例句查询', **page_content, js_api=FileApi(html),
                          width=1120, height=800, min_size=(640, 480), text_select=True,
                          background_color='#F5F6F3')
    storage = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / '.cache'))) / 'SignSongAgent' / 'WebView2'
    engine = 'edgechromium' if sys.platform == 'win32' else 'cocoa' if sys.platform == 'darwin' else 'qt'
    webview.start(gui=engine, private_mode=False, storage_path=str(storage))


if __name__ == '__main__':
    main()
