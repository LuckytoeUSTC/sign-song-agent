"""Download selected dictionary folders from the public resource catalog."""
import argparse
import json
import sys
import time
from http.client import IncompleteRead
from pathlib import Path, PurePosixPath
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]


def local_path(root, key):
    parts = PurePosixPath(key).parts
    if len(parts) < 3 or parts[0] != "dictionaries" or ".." in parts:
        raise ValueError("词典目录中的文件路径无效")
    target = root.joinpath(*parts[1:]).resolve()
    if not target.is_relative_to(root.resolve()):
        raise ValueError("文件路径超出词典目录")
    return target


def download(base_url, resource, root):
    target = local_path(root, resource["key"])
    size = resource["size"]
    if target.is_file():
        if target.stat().st_size == size:
            print("已安装：", target.name, flush=True)
            return
        raise ValueError(f"已有同名文件，保留原文件，请先确认版本：{target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(target.name + ".part")
    url = base_url.rstrip("/") + "/" + quote(resource["key"], safe="/")
    for attempt in range(3):
        try:
            offset = partial.stat().st_size if partial.exists() else 0
            if offset > size:
                raise ValueError(f"临时下载文件过大，请先处理：{partial}")
            if offset != size:
                headers = {"User-Agent": "sign-song-agent", "Accept-Encoding": "identity"}
                if offset:
                    headers["Range"] = f"bytes={offset}-"
                with urlopen(Request(url, headers=headers), timeout=60) as response:
                    resumed = response.status == 206
                    if resumed and not response.headers.get("Content-Range", "").startswith(f"bytes {offset}-"):
                        raise ValueError("下载服务返回了不匹配的续传位置")
                    if response.status not in (200, 206):
                        raise ValueError(f"下载服务返回 HTTP {response.status}")
                    mode = "ab" if resumed else "wb"
                    received = offset if resumed else 0
                    with partial.open(mode) as output:
                        while chunk := response.read(1024 * 1024):
                            received += len(chunk)
                            if received > size:
                                raise ValueError("下载内容超出目录记录的文件长度")
                            output.write(chunk)
            if partial.stat().st_size != size:
                raise OSError("下载尚未完成")
            if target.exists():
                raise ValueError(f"下载期间出现同名文件，保留原文件：{target}")
            partial.replace(target)
            print("下载完成：", target.name, flush=True)
            return
        except HTTPError as error:
            if error.code in (401, 403, 404):
                raise ValueError(f"公开下载地址不可用（HTTP {error.code}），请联系项目维护者") from None
            if attempt == 2:
                raise
        except (URLError, OSError, IncompleteRead):
            if attempt == 2:
                raise
        time.sleep(attempt + 1)


def main():
    parser = argparse.ArgumentParser(description="自动下载国家通用词典，或下载用户选择的其它词典")
    parser.add_argument("--list", action="store_true", help="显示可下载词典，不访问网络")
    parser.add_argument("--dictionary", action="append", default=[], help="选择词典名称，可重复使用")
    parser.add_argument("--default", action="store_true", help="下载默认国家通用手语词典")
    parser.add_argument("--manifest", type=Path, default=ROOT / "dictionaries/resources.json")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dictionaries")
    args = parser.parse_args()
    catalog = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    dictionaries = catalog["dictionaries"]
    if args.list:
        for name, item in dictionaries.items():
            mb = sum(r["size"] for r in item["files"]) / 1_000_000
            print(f"{name}　约 {mb:.1f} MB" + ("（接管后自动下载）" if item.get("default") else "（用户选择）"))
        return
    names = list(dict.fromkeys(args.dictionary))
    if args.default or not names:
        names = list(dict.fromkeys([n for n, d in dictionaries.items() if d.get("default")] + names))
    for name in names:
        if name not in dictionaries:
            raise ValueError(f"目录中没有这套词典：{name}。可使用 --list 查看名称。")
    base_url = catalog.get("base_url", "")
    parsed = urlsplit(base_url)
    if parsed.scheme not in ("https", "http") or not parsed.netloc or parsed.username or parsed.password:
        raise ValueError("尚未配置有效的公开下载域名，请联系项目维护者；无需填写 R2 密钥")
    for name in names:
        print("准备词典：", name, flush=True)
        for resource in dictionaries[name]["files"]:
            download(base_url, resource, args.output_dir)
    print("所选词典已就绪。", flush=True)


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, URLError, KeyError) as error:
        print(f"下载未完成：{error}", file=sys.stderr)
        sys.exit(1)
