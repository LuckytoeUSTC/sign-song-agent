from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
try:
    from desktop_query import main
    main()
except Exception as error:
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, f'无法启动规则与例句查询：\n{error}\n\n请让智能体检查 Python、pywebview 和 WebView2 配置。', '规则与例句查询', 16)
