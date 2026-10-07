"""Minimal read-only MDict 2.0 reader for the local sign-language dictionaries.

Supports unencrypted MDX/MDD files using zlib or uncompressed blocks.  It is
deliberately small and dependency-free so the original dictionaries remain
untouched.  Typical usage:

    python mdict_reader.py query dictionary.mdx "太阳"
    python mdict_reader.py resources dictionary.mdd --contains "太阳"
    python mdict_reader.py extract dictionary.mdd "\\path\\image.jpg" out.jpg
"""

from __future__ import annotations

import argparse
import bisect
import html
import re
import struct
import sys
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Iterable


def _u16be(data: bytes) -> int:
    return struct.unpack(">H", data)[0]


def _u32be(data: bytes) -> int:
    return struct.unpack(">I", data)[0]


def _u64be(data: bytes) -> int:
    return struct.unpack(">Q", data)[0]


def _read_exact(stream: BinaryIO, size: int) -> bytes:
    data = stream.read(size)
    if len(data) != size:
        raise EOFError(f"wanted {size} bytes, got {len(data)}")
    return data


def _decompress_block(block: bytes, expected_size: int | None = None) -> bytes:
    if len(block) < 8:
        raise ValueError("truncated compressed block")
    kind = block[:4]
    payload = block[8:]
    if kind == b"\x00\x00\x00\x00":
        result = payload
    elif kind == b"\x02\x00\x00\x00":
        result = zlib.decompress(payload)
    elif kind == b"\x01\x00\x00\x00":
        raise NotImplementedError("LZO-compressed MDict blocks are not supported")
    else:
        raise ValueError(f"unknown MDict compression marker {kind.hex()}")
    if expected_size is not None and len(result) != expected_size:
        raise ValueError(
            f"decompressed size mismatch: expected {expected_size}, got {len(result)}"
        )
    return result


@dataclass(frozen=True)
class Entry:
    key: str
    start: int
    end: int


@dataclass(frozen=True)
class RecordBlock:
    file_offset: int
    compressed_size: int
    decompressed_size: int
    logical_start: int

    @property
    def logical_end(self) -> int:
        return self.logical_start + self.decompressed_size


class MDictReader:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.header: dict[str, str] = {}
        self.encoding = "utf-8"
        self.is_mdd = self.path.suffix.lower() == ".mdd"
        self.entries: list[Entry] = []
        self._record_blocks: list[RecordBlock] = []
        self._block_starts: list[int] = []
        self._parse()

    @staticmethod
    def _parse_header_attrs(text: str) -> dict[str, str]:
        return {
            key: html.unescape(value)
            for key, value in re.findall(r'([A-Za-z0-9_]+)="([^"]*)"', text)
        }

    def _decode_key(self, raw: bytes) -> str:
        return raw.decode(self.encoding, errors="replace")

    @property
    def _text_width(self) -> int:
        return 2 if self.encoding.lower().startswith("utf-16") else 1

    def _read_number(self, stream: BinaryIO) -> int:
        return _u64be(_read_exact(stream, 8))

    def _parse(self) -> None:
        with self.path.open("rb") as stream:
            header_size = _u32be(_read_exact(stream, 4))
            header_raw = _read_exact(stream, header_size)
            _read_exact(stream, 4)  # header Adler-32
            header_text = header_raw.decode("utf-16le", errors="replace").rstrip("\x00\r\n")
            self.header = self._parse_header_attrs(header_text)
            if self.header.get("Encrypted", "No") not in {"No", "0", ""}:
                raise NotImplementedError("encrypted MDict files are not supported")
            declared_encoding = self.header.get("Encoding", "UTF-8")
            if self.is_mdd and not declared_encoding:
                # MDD resource keys are UTF-16LE even when the header leaves
                # Encoding blank (the binary resource payload is not decoded).
                declared_encoding = "UTF-16LE"
            self.encoding = declared_encoding.lower().replace("utf8", "utf-8")

            key_block_count = self._read_number(stream)
            entry_count = self._read_number(stream)
            key_info_decompressed_size = self._read_number(stream)
            key_info_size = self._read_number(stream)
            key_blocks_size = self._read_number(stream)
            _read_exact(stream, 4)  # key-header Adler-32

            info_block = _read_exact(stream, key_info_size)
            key_info = _decompress_block(info_block, key_info_decompressed_size)
            key_block_specs = list(self._parse_key_block_info(key_info, key_block_count))
            if sum(count for count, _, _ in key_block_specs) != entry_count:
                raise ValueError("key-block entry count mismatch")
            if sum(comp for _, comp, _ in key_block_specs) != key_blocks_size:
                raise ValueError("key-block byte size mismatch")

            key_blocks_raw = _read_exact(stream, key_blocks_size)
            pairs: list[tuple[str, int]] = []
            cursor = 0
            for expected_entries, compressed_size, decompressed_size in key_block_specs:
                block = key_blocks_raw[cursor : cursor + compressed_size]
                cursor += compressed_size
                decoded = _decompress_block(block, decompressed_size)
                block_pairs = list(self._parse_key_block(decoded))
                if len(block_pairs) != expected_entries:
                    raise ValueError("decoded key-block entry count mismatch")
                pairs.extend(block_pairs)

            record_block_count = self._read_number(stream)
            record_entry_count = self._read_number(stream)
            record_info_size = self._read_number(stream)
            record_blocks_size = self._read_number(stream)
            if record_entry_count != entry_count:
                raise ValueError("record/key entry count mismatch")

            record_specs = []
            for _ in range(record_block_count):
                record_specs.append((self._read_number(stream), self._read_number(stream)))
            if record_info_size != record_block_count * 16:
                raise ValueError("record-block info size mismatch")
            if sum(comp for comp, _ in record_specs) != record_blocks_size:
                raise ValueError("record-block byte size mismatch")

            record_data_start = stream.tell()
            file_cursor = record_data_start
            logical_cursor = 0
            for compressed_size, decompressed_size in record_specs:
                self._record_blocks.append(
                    RecordBlock(
                        file_offset=file_cursor,
                        compressed_size=compressed_size,
                        decompressed_size=decompressed_size,
                        logical_start=logical_cursor,
                    )
                )
                file_cursor += compressed_size
                logical_cursor += decompressed_size
            self._block_starts = [block.logical_start for block in self._record_blocks]

            if len(pairs) != entry_count:
                raise ValueError("decoded key count mismatch")
            for index, (key, start) in enumerate(pairs):
                end = pairs[index + 1][1] if index + 1 < len(pairs) else logical_cursor
                self.entries.append(Entry(key=key, start=start, end=end))

    def _parse_key_block_info(
        self, data: bytes, key_block_count: int
    ) -> Iterable[tuple[int, int, int]]:
        cursor = 0
        width = self._text_width
        for _ in range(key_block_count):
            entry_count = _u64be(data[cursor : cursor + 8])
            cursor += 8
            head_len = _u16be(data[cursor : cursor + 2])
            cursor += 2 + head_len * width + width
            tail_len = _u16be(data[cursor : cursor + 2])
            cursor += 2 + tail_len * width + width
            compressed_size = _u64be(data[cursor : cursor + 8])
            decompressed_size = _u64be(data[cursor + 8 : cursor + 16])
            cursor += 16
            yield entry_count, compressed_size, decompressed_size
        if cursor != len(data):
            raise ValueError(f"unparsed key-block-info bytes: {len(data) - cursor}")

    def _parse_key_block(self, data: bytes) -> Iterable[tuple[str, int]]:
        cursor = 0
        term = b"\x00\x00" if self._text_width == 2 else b"\x00"
        while cursor < len(data):
            start = _u64be(data[cursor : cursor + 8])
            cursor += 8
            if self._text_width == 1:
                end = data.index(term, cursor)
            else:
                end = cursor
                while end + 1 < len(data) and data[end : end + 2] != term:
                    end += 2
            key = self._decode_key(data[cursor:end])
            cursor = end + len(term)
            yield key, start

    def find(self, term: str, exact: bool = True, limit: int = 20) -> list[Entry]:
        if exact:
            folded = term.casefold()
            return [entry for entry in self.entries if entry.key.casefold() == folded][:limit]
        folded = term.casefold()
        return [entry for entry in self.entries if folded in entry.key.casefold()][:limit]

    def _read_record_block(self, index: int) -> bytes:
        block = self._record_blocks[index]
        with self.path.open("rb") as stream:
            stream.seek(block.file_offset)
            compressed = _read_exact(stream, block.compressed_size)
        return _decompress_block(compressed, block.decompressed_size)

    def read_entry(self, entry: Entry) -> bytes:
        if not self._record_blocks:
            return b""
        first = bisect.bisect_right(self._block_starts, entry.start) - 1
        last = bisect.bisect_right(self._block_starts, max(entry.start, entry.end - 1)) - 1
        chunks = []
        for index in range(first, last + 1):
            block = self._record_blocks[index]
            raw = self._read_record_block(index)
            local_start = max(entry.start, block.logical_start) - block.logical_start
            local_end = min(entry.end, block.logical_end) - block.logical_start
            chunks.append(raw[local_start:local_end])
        return b"".join(chunks)


def _make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    inspect_cmd = sub.add_parser("inspect", help="show dictionary metadata")
    inspect_cmd.add_argument("dictionary", type=Path)

    query_cmd = sub.add_parser("query", help="query text records")
    query_cmd.add_argument("dictionary", type=Path)
    query_cmd.add_argument("term")
    query_cmd.add_argument("--contains", action="store_true")
    query_cmd.add_argument("--limit", type=int, default=20)
    query_cmd.add_argument("--summary", action="store_true")

    resources_cmd = sub.add_parser("resources", help="list resource keys")
    resources_cmd.add_argument("dictionary", type=Path)
    resources_cmd.add_argument("--contains", default="")
    resources_cmd.add_argument("--limit", type=int, default=50)

    extract_cmd = sub.add_parser("extract", help="extract one binary resource")
    extract_cmd.add_argument("dictionary", type=Path)
    extract_cmd.add_argument("key")
    extract_cmd.add_argument("output", type=Path)

    images_cmd = sub.add_parser(
        "images", help="resolve a word entry and extract every referenced original image"
    )
    images_cmd.add_argument("dictionary", type=Path, help="source MDX")
    images_cmd.add_argument("resources", type=Path, help="companion MDD")
    images_cmd.add_argument("term")
    images_cmd.add_argument("output_dir", type=Path)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _make_parser().parse_args(argv)
    reader = MDictReader(args.dictionary)
    if args.command == "inspect":
        print(f"Title: {reader.header.get('Title', '')}")
        print(f"Encoding: {reader.encoding}")
        print(f"Entries: {len(reader.entries)}")
        print(f"Record blocks: {len(reader._record_blocks)}")
        return 0
    if args.command == "resources":
        needle = args.contains.casefold()
        matches = (e for e in reader.entries if needle in e.key.casefold())
        for entry in list(matches)[: args.limit]:
            print(entry.key)
        return 0
    if args.command == "query":
        entries = reader.find(args.term, exact=not args.contains, limit=args.limit)
        for index, entry in enumerate(entries):
            if index:
                print("\n---")
            print(f"[{entry.key}]")
            raw = reader.read_entry(entry)
            record = raw.decode(reader.encoding, errors="replace").rstrip("\x00")
            if args.summary and not record.startswith("@@@LINK="):
                label_match = re.search(
                    r'<span\s+class=["\']word["\']>(.*?)</span>', record, re.I | re.S
                )
                contents = re.findall(
                    r'<p\s+class=["\']content["\']>(.*?)</p>', record, re.I | re.S
                )
                images = re.findall(
                    r'<img\b[^>]*\bsrc=["\']([^"\']+)["\']', record, re.I
                )
                clean = lambda value: html.unescape(re.sub(r"<[^>]+>", "", value)).strip()
                if label_match:
                    print(f"label: {clean(label_match.group(1))}")
                for content in contents:
                    print(f"description: {clean(content)}")
                for source in images:
                    print(f"image: {source}")
            else:
                print(record)
        return 0 if entries else 1
    if args.command == "extract":
        entries = reader.find(args.key, exact=True, limit=2)
        if not entries:
            print(f"resource not found: {args.key}", file=sys.stderr)
            return 1
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(reader.read_entry(entries[0]))
        print(args.output)
        return 0
    if args.command == "images":
        resource_reader = MDictReader(args.resources)
        pending = [args.term]
        visited: set[str] = set()
        definitions: list[tuple[str, str]] = []
        while pending:
            current = pending.pop(0)
            if current.casefold() in visited:
                continue
            visited.add(current.casefold())
            for entry in reader.find(current, exact=True, limit=100):
                record = reader.read_entry(entry).decode(reader.encoding, errors="replace").rstrip("\x00")
                link = re.fullmatch(r"\s*@@@LINK=(.*?)\s*", record, flags=re.DOTALL)
                if link:
                    pending.append(link.group(1).strip())
                else:
                    definitions.append((entry.key, record))
        if not definitions:
            print(f"entry not found: {args.term}", file=sys.stderr)
            return 1
        args.output_dir.mkdir(parents=True, exist_ok=True)
        extracted = 0
        for definition_index, (key, record) in enumerate(definitions, start=1):
            print(f"[{key}]")
            for source in re.findall(r'<img\b[^>]*\bsrc=["\']([^"\']+)["\']', record, re.I):
                resource_key = "\\" + source.replace("/", "\\").lstrip("\\")
                resources = resource_reader.find(resource_key, exact=True, limit=2)
                if not resources:
                    print(f"missing resource: {resource_key}", file=sys.stderr)
                    continue
                extracted += 1
                basename = Path(source.replace("\\", "/")).name
                destination = args.output_dir / f"{definition_index:02d}_{extracted:02d}_{basename}"
                destination.write_bytes(resource_reader.read_entry(resources[0]))
                print(destination)
        return 0 if extracted else 1
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
