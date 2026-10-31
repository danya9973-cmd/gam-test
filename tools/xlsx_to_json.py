# -*- coding: utf-8 -*-
"""xlsx 원본을 JSON 사본으로 옮긴다 — 형식만 바꾸고 내용은 거르지 않는다.

data/raw/gcgf_*.xlsx 는 받은 그대로 두고, 모두 묶어 data/converted/gcgf.json 하나로 만든다.
모양은 {"supply": [...], "gmoney": [...], ...} — 파일 이름에서 gcgf_ 를 뗀 것이 열쇠다.
첫 줄을 머리줄로 삼아 한 줄을 {머리줄: 값} 하나로 옮긴다.
합계 줄 · 엑셀식 날짜(45408 같은 숫자) · 빈 칸도 그대로 둔다 — 거르고 푸는 것은 계산 단계의 일이다.

openpyxl 없이 파이썬 기본 기능만 쓴다. xlsx 는 XML 여러 개를 zip 으로 묶은 파일이다.

실행 (저장소 최상위에서):
    python tools/xlsx_to_json.py
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

DATA = Path(__file__).resolve().parent.parent / "data"
RAW = DATA / "raw"
OUT = DATA / "converted"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_index(ref: str) -> int:
    """'B12' → 1. 칸 주소의 알파벳으로 몇 번째 열인지 센다 (빈 칸을 건너뛰어도 열이 밀리지 않게)."""
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_value(c: ET.Element, shared: list[str]):
    t = c.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.iterfind(".//m:t", NS))
    v = c.find("m:v", NS)
    if v is None or v.text is None:
        return None
    if t == "s":
        return shared[int(v.text)]
    if t in ("str", "e"):
        return v.text
    if t == "b":
        return v.text == "1"
    # 숫자는 엑셀에 적힌 그대로 읽는다. 정수로 떨어지면 정수로 둔다
    x = float(v.text)
    return int(x) if x.is_integer() and re.fullmatch(r"-?\d+", v.text) else x


def read_rows(path: Path) -> list[list]:
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iterfind(".//m:t", NS))
                      for si in root.iterfind("m:si", NS)]
        sheets = sorted((n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
                        key=lambda n: int(re.search(r"\d+", n).group(0)))
        found = []
        for name in sheets:
            rows = []
            for r in ET.fromstring(z.read(name)).iterfind(".//m:sheetData/m:row", NS):
                row: list = []
                for c in r.iterfind("m:c", NS):
                    i = col_index(c.get("r"))
                    row += [None] * (i + 1 - len(row))
                    row[i] = cell_value(c, shared)
                if any(x is not None for x in row):
                    rows.append(row)
            if rows:
                found.append(rows)
    # 내용이 든 시트가 둘 이상이면 어느 것을 옮길지 사람이 정해야 한다
    assert len(found) == 1, f"{path.name}: 내용이 든 시트가 {len(found)}개"
    return found[0]


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    files = sorted(RAW.glob("gcgf_*.xlsx"))
    bundle = {}
    for src in files:
        head, *body = read_rows(src)
        assert len(set(head)) == len(head), f"{src.name}: 머리줄에 같은 이름이 있음"
        key = src.stem.removeprefix("gcgf_")
        bundle[key] = [dict(zip(head, row + [None] * (len(head) - len(row)))) for row in body]
        print(f"  {src.name:<26} → {key:<12} {len(bundle[key]):>6,}행")
    dest = OUT / "gcgf.json"
    dest.write_text(json.dumps(bundle, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n총 {len(files)}건 → {dest} ({dest.stat().st_size/1024:,.0f}KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
