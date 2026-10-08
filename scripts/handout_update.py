"""Local takeover state helpers. No independent CLI or teaching-data migration."""
from pathlib import Path
import hashlib
import json
import os
import re
import tempfile
from datetime import datetime, timezone
from docx import Document

TOOL_VERSION = '3.1.1'
UPDATE_BEGIN = '<!-- COVER_UPDATE_3.1.1_BEGIN -->'
UPDATE_END = '<!-- COVER_UPDATE_3.1.1_END -->'

def compatibility_check(root):
    """No user song reads or writes: one legacy synthetic CLI build and audit."""
    from make_handout import DEFAULT_LAYOUT
    import ast
    import subprocess
    import sys
    root = Path(root).resolve()
    for name in ('make_handout.py', 'audit_docx.py', 'handout_update.py'):
        ast.parse((root / 'scripts' / name).read_text(encoding='utf-8-sig'))
    with tempfile.TemporaryDirectory(prefix='sign-handout-compat-') as temp:
        temp = Path(temp)
        source = temp / 'legacy.md'
        source.write_text('# 兼容检查\n合成句（*动作①++*/~~省略②~~）\n', encoding='utf-8')
        layout = temp / 'layout.json'
        layout.write_text(json.dumps(DEFAULT_LAYOUT), encoding='utf-8')
        output = temp / 'legacy.docx'
        result = subprocess.run([sys.executable, str(root / 'scripts/make_handout.py'), str(source),
                                 '--out', str(output), '--height', '0.79', '--layout', str(layout)],
                                capture_output=True, text=True, encoding='utf-8', errors='replace')
        if result.returncode:
            raise RuntimeError(result.stderr or result.stdout)
        if '合成句（动作①++/省略②）' not in '\n'.join(p.text for p in Document(output).paragraphs):
            raise RuntimeError('旧Markdown正文兼容检查失败')
        result = subprocess.run([sys.executable, str(root / 'scripts/audit_docx.py'), str(output)],
                                capture_output=True, text=True, encoding='utf-8', errors='replace')
        if result.returncode:
            raise RuntimeError(result.stderr or result.stdout)
    return ['旧命令、Markdown标记、全部旧版式字段和audit_docx调用通过合成检查']


def update_record_path(root):
    identity = hashlib.sha256(str(Path(root).resolve()).casefold().encode('utf-8')).hexdigest()[:16]
    local = Path(os.environ.get('LOCALAPPDATA', Path.home() / '.local/share'))
    return local / 'SignSongAgent' / 'Updates' / (identity + '.json')


def update_fingerprint(root):
    # Generated pages and user data are deliberately excluded.
    root = Path(root)
    names = ('scripts/make_handout.py', 'scripts/audit_docx.py', 'scripts/handout_update.py', 'scripts/TOOLING.md', '使用说明.md', 'README.md')
    result = {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in names}
    agents = (root / 'AGENTS.md').read_text(encoding='utf-8')
    without_block = re.sub(re.escape(UPDATE_BEGIN) + r'.*?' + re.escape(UPDATE_END) + r'\r?\n?', '', agents, flags=re.S)
    result['AGENTS.md（更新区块以外）'] = hashlib.sha256(without_block.encode('utf-8')).hexdigest()
    return result


def write_record(path, record):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)


def update_takeover(action, root=None, state_path=None, note=''):
    """Agent-controlled completion; no autonomous migration or song editing."""
    root = Path(root or Path(__file__).resolve().parents[1])
    state_path = Path(state_path) if state_path else update_record_path(root)
    record_problem = ''
    try:
        record = json.loads(state_path.read_text(encoding='utf-8'))
        if not isinstance(record, dict):
            raise ValueError('记录不是对象')
    except FileNotFoundError:
        record = {}
    except ValueError:
        record = {}
        record_problem = '已有本机记录无法解析，不能据此认定更新成功'
    fingerprint = update_fingerprint(root)
    same = record.get('version') == TOOL_VERSION and record.get('files') == fingerprint
    status = '已完成' if same and record.get('status') == 'complete' else '尚未完成' if record else '首次接管'
    if action == 'check':
        checks = compatibility_check(root)
        return dict(version=TOOL_VERSION, status=status, checks=checks, record=str(state_path), record_warning=record_problem)
    if action == 'fail':
        write_record(state_path, dict(version=TOOL_VERSION, files=fingerprint, status='incomplete',
                                      completed=note, remaining='接管智能体须核对原因后安全续做',
                                      updated=datetime.now(timezone.utc).isoformat()))
        return dict(status='尚未完成', record=str(state_path))
    if action != 'complete':
        raise ValueError('更新动作须为check、complete或fail')
    if not note.strip():
        raise ValueError('完成前须用--update-note简述本机实际调整、任务保留及验证结果；不写私人对话')
    checks = compatibility_check(root)  # Always recheck, even after re-covering the same version.
    agents = root / 'AGENTS.md'
    text = agents.read_text(encoding='utf-8')
    pattern = re.compile(re.escape(UPDATE_BEGIN) + r'.*?' + re.escape(UPDATE_END) + r'\r?\n?', re.S)
    if text.count(UPDATE_BEGIN) != text.count(UPDATE_END) or text.count(UPDATE_BEGIN) > 1:
        raise ValueError('更新说明标记不完整或重复；保留原文件，先检查')
    cleaned = pattern.sub('', text)
    if UPDATE_BEGIN not in text and not (same and record.get('status') == 'complete'):
        raise ValueError('首次完成时未找到本版本更新说明；先核对AGENTS是否已正确覆盖')
    original_bytes = agents.read_bytes()
    backup = state_path.parent / (state_path.stem + '-AGENTS-before-cleanup.md')
    temporary = agents.with_name('.AGENTS-update.tmp')
    completion = dict(version=TOOL_VERSION, files=fingerprint, status='complete',
                      checks=checks, result=note, updated=datetime.now(timezone.utc).isoformat())
    try:
        if cleaned != text:
            backup.parent.mkdir(parents=True, exist_ok=True)
            if not backup.exists():
                backup.write_bytes(original_bytes)
            temporary.write_text(cleaned, encoding='utf-8')
        # Persist a verified completion record before removing the instructions.
        write_record(state_path, completion)
        if cleaned != text:
            temporary.replace(agents)
    except Exception:
        temporary.unlink(missing_ok=True)
        # A failed cleanup is not complete. Keep/recover the instructions.
        if agents.read_bytes() != original_bytes:
            agents.write_bytes(original_bytes)
        failed = completion.copy()
        failed.update(status='incomplete', remaining='记录或区块清理失败；说明已保留，修复后续做')
        try:
            write_record(state_path, failed)
        except OSError:
            pass
        raise
    return dict(status='已完成', checks=checks, record=str(state_path), removed_block=cleaned != text)


