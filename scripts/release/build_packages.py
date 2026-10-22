"""Build native query applications and full/query-only release archives."""
from pathlib import Path
import hashlib
import importlib.metadata as metadata
import json
import os
import platform
import shutil
import subprocess
import sys
import tarfile
import zipfile

SRC = Path(__file__).resolve().parents[2]
BUILD = Path(os.environ.get('PACKAGE_BUILD', str(SRC.parent / '构建记录/三系统发布'))).resolve()
BUILD.mkdir(parents=True, exist_ok=True)
SYSTEM = {'win32': 'windows', 'darwin': 'macos'}.get(sys.platform, 'linux')
ARCH = 'arm64' if platform.machine().lower() in {'arm64', 'aarch64'} else 'x64'
SUFFIX = f'{SYSTEM}-{ARCH}'
STAGE = BUILD / SUFFIX
STAGE.mkdir(exist_ok=True)
QUERY = STAGE / 'NO-AI-ZONE'
FULL = STAGE / '手语歌智能工坊'
OUTPUT = BUILD / 'artifacts'
OUTPUT.mkdir(exist_ok=True)


def run(args, **kwargs):
    subprocess.run([str(x) for x in args], check=True, **kwargs)


env = os.environ.copy()
env.update(NO_AI_OUTPUT=str(QUERY), NO_AI_BUILD=str(STAGE))
run([sys.executable, SRC / 'scripts/release/prepare_query_data.py'], env=env)

# One executable on Windows/Linux; one application bundle in Finder on macOS.
cmd = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--windowed',
       '--onedir' if SYSTEM == 'macos' else '--onefile', '--name', 'Query',
       '--distpath', STAGE / 'dist', '--workpath', STAGE / 'work', '--specpath', STAGE,
       '--add-data', str(STAGE / 'query_page.html') + os.pathsep + '.',
       '--hidden-import', 'markdown.extensions.tables', '--hidden-import', 'markdown.extensions.fenced_code',
       '--hidden-import', 'markdown.extensions.sane_lists']
exclude = ['cef', 'android']
exclude += ['gtk', 'qt', 'cocoa'] if SYSTEM == 'windows' else ['gtk', 'qt', 'winforms'] if SYSTEM == 'macos' else ['gtk', 'cocoa', 'winforms']
for backend in exclude:
    cmd += ['--exclude-module', 'webview.platforms.' + backend]
cmd.append(SRC / 'scripts/release/portable_query.py')
run(cmd)

tracked = subprocess.check_output(['git', '-C', str(SRC), 'ls-files', '-z']).decode().split('\0')
for name in tracked:
    if not name or name.startswith(('scripts/release/', '.github/')):
        continue
    src = SRC / name
    if not src.is_file() or src.name == 'debug.log':
        continue
    dst = FULL / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)

if SYSTEM == 'macos':
    binary = STAGE / 'dist/Query.app'
    appname = '规则与例句查询.app'
else:
    binary = STAGE / 'dist' / ('Query.exe' if SYSTEM == 'windows' else 'Query')
    appname = '规则与例句查询.exe' if SYSTEM == 'windows' else '规则与例句查询'
for folder in [QUERY, FULL]:
    folder.mkdir(parents=True, exist_ok=True)
    if binary.is_dir():
        shutil.copytree(binary, folder / appname, dirs_exist_ok=True, symlinks=True)
    else:
        shutil.copy2(binary, folder / appname)

requirements = {
    'windows': 'Windows 10/11，x64。查询窗口需要WebView2；系统缺少时需补装。',
    'macos': 'macOS 14或更新版本，Apple Silicon（M系列，arm64）。本版未签名、公证，首次打开可能需要在系统隐私与安全性设置中确认来源。Intel Mac请勿选择此包。',
    'linux': 'Linux x64，基于Ubuntu 22.04构建与测试。需要桌面环境；其他发行版可能需要补齐Qt所需的系统图形库。'
}[SYSTEM]
launch = '双击“' + appname + '”' if SYSTEM != 'linux' else '运行“规则与例句查询”（若无执行权限，先在文件属性中允许作为程序执行）'
readme = f'''# NO-AI-ZONE · 规则与例句查询

这是给电脑里没有装智能体，或者反对AI的人类准备的规则与例句查询程序。两张STOP图片也在包里，欢迎围观。

## 怎样开始

完整解压后，{launch}。查询程序已包含Python运行环境，无需安装Python、智能体或模型账号，查询在本地离线进行。

本包：**{SUFFIX}**。{requirements}

Word、PDF原件使用电脑默认阅读软件打开；规则和辨析文字直接在程序阅读窗口打开。移动时请带上整个NO-AI-ZONE文件夹，保留资料目录。

## 有什么

- 搜索773条例句与修改辨析，按18条转写规则或资料来源筛选。
- 点击规则编号反查相关例子，查看完整会话与修改理由。
- 打开对应教材、小班讲义、歌曲讲义与收录案例原文。

这是原工坊的精简查阅版。它保留规则、例句和必要原件，备课中的选歌、翻译讨论、词典查词、自动贴图、Word制作及经验审核汇入等功能由手语歌智能工坊完整版提供。本包不包含这些功能，也不包含原工坊的全部资料。

历史案例保留原有出处与修改记录，纯AI试制稿未随包提供，相应链接注明“本查阅包未收录”。这些材料供比较与学习，实际表达仍需结合语境和老师试打。

## 选哪个版本

Windows电脑选`windows-x64`；M系列Mac选`macos-arm64`；Linux x64电脑选`linux-x64`。需要智能体辅助完整备课，请下载对应系统的`sign-song-agent-v3.1.1-full-*`包。已有v3.1.0项目先阅读覆盖更新说明。

讲义及学习材料沿用原授权范围，扩大分发前请确认制作者允许的用途。程序与运行组件许可见“许可说明.txt”。反馈：liqiding@mail.ustc.edu.cn
'''
(QUERY / 'README.md').write_text(readme, encoding='utf-8')
(FULL / '本系统使用说明.txt').write_text(f'手语歌智能工坊 v3.1.1 · {SUFFIX}\n\n{requirements}\n\n查询程序：{launch}，此功能无需安装Python。完整备课请在本目录接入智能体，先阅读README.md和AGENTS.md，由智能体准备Python及制作所需依赖。\n\n这个完整包用于新项目。已有项目先核对覆盖更新说明，保留自己的工作台和资料。\n', encoding='utf-8-sig')

license_text = [(SRC / 'LICENSE').read_text(encoding='utf-8')]
for package in ['pywebview', 'pythonnet', 'clr_loader', 'cffi', 'pycparser', 'bottle', 'proxy_tools', 'typing_extensions', 'setuptools', 'packaging', 'markdown', 'PySide6', 'PySide6_Essentials', 'PySide6_Addons', 'shiboken6', 'pyobjc-core', 'pyobjc-framework-Cocoa', 'pyobjc-framework-WebKit']:
    try:
        dist = metadata.distribution(package)
    except metadata.PackageNotFoundError:
        continue
    for entry in dist.files or []:
        if any(k in entry.name.lower() for k in ['license', 'licence', 'copying', 'notice']):
            file = Path(dist.locate_file(entry))
            if file.is_file():
                license_text.append('\n\n=== ' + package + ': ' + str(entry) + ' ===\n' + file.read_text(encoding='utf-8', errors='replace'))
for file in [Path(sys.base_prefix) / 'LICENSE.txt', Path(sys.base_prefix) / 'LICENSE']:
    if file.is_file():
        license_text.append('\n\n=== Python ===\n' + file.read_text(encoding='utf-8', errors='replace'))
        break
for folder in [QUERY, FULL]:
    (folder / '许可说明.txt').write_text('\n'.join(license_text), encoding='utf-8')

# Validate frozen binaries after relocating their resources into final package roots.
for folder, label in [(QUERY, 'query'), (FULL, 'full')]:
    exe = folder / appname
    if SYSTEM == 'macos':
        exe = exe / 'Contents/MacOS/Query'
    report = STAGE / (label + '-check.json')
    run([exe, '--check', report], timeout=60)
    assert json.loads(report.read_text(encoding='utf-8'))['frozen']
    report = STAGE / (label + '-smoke.json')
    run([exe, '--smoke', report], timeout=90)
    result = json.loads(report.read_text(encoding='utf-8'))
    if not result.get('ok'):
        raise RuntimeError(result)

for folder, name in [(QUERY, 'no-ai-zone-v3.1.1-' + SUFFIX), (FULL, 'sign-song-agent-v3.1.1-full-' + SUFFIX)]:
    for path in sorted((x for x in folder.rglob('*') if x.is_dir()), key=lambda p: len(p.parts), reverse=True):
        if not any(path.iterdir()):
            path.rmdir()
    if SYSTEM == 'linux':
        archive = OUTPUT / (name + '.tar.gz')
        with tarfile.open(archive, 'w:gz') as tar:
            tar.add(folder, arcname=folder.name)
    elif SYSTEM == 'macos':
        archive = OUTPUT / (name + '.zip')
        run(['ditto', '-c', '-k', '--sequesterRsrc', '--keepParent', folder, archive])
    else:
        archive = OUTPUT / (name + '.zip')
        with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for path in folder.rglob('*'):
                if path.is_file():
                    z.write(path, path.relative_to(folder.parent))
    (OUTPUT / (archive.name + '.sha256')).write_text(hashlib.sha256(archive.read_bytes()).hexdigest() + '  ' + archive.name + '\n', encoding='ascii')
for name in ['query-check.json', 'query-smoke.json', 'full-check.json', 'full-smoke.json']:
    shutil.copy2(STAGE / name, OUTPUT / (SUFFIX + '-' + name))
print('Packages ready:', SUFFIX)
