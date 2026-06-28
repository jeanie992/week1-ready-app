# -*- coding: utf-8 -*-
"""
dart_parse.py : DART 감사/사업보고서 추출 텍스트(01_parsed/DART_*.txt)를
                재무상태표·손익계산서·현금흐름표의 구조화 표로 파싱

핵심 : 보고서마다 당기/전기 2개년이 들어있으므로 3개 보고서를 합쳐 다개년 시계열 구성
반환 : {"years":[...], "BS":{계정:{연도:값}}, "IS":{...}, "CF":{...}}  (값 단위 = 원)
회사 무관 범용 파서 : 계정명을 미리 알 필요 없이 '계정 라벨 → 당기·전기 숫자' 순서로 추출
"""
from __future__ import annotations
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PARSED = ROOT / "01_parsed"

SECTIONS = [
    ("BS", ["재 무 상 태 표", "재무상태표"], ["손 익 계 산 서", "손익계산서", "포 괄 손 익"]),
    ("IS", ["손 익 계 산 서", "손익계산서"], ["자 본 변 동 표", "자본변동표"]),
    ("CF", ["현 금 흐 름 표", "현금흐름표"], ["주석", "주 석", "외부감사"]),
]


def _is_noise(l: str) -> bool:
    if re.search(r"제\s*\d+\s*[\(기]", l):
        return True
    if any(k in l for k in ("단위", "과 목", "과목", "회사명", "별첨", "별 첨", "보고기간")):
        return True
    if re.search(r"주식회사|㈜", l):          # 회사명 라인
        return True
    if l in ("당기", "전기", "당 기", "전 기"):
        return True
    if re.search(r"\((당|전)\)", l):
        return True
    if re.fullmatch(r"[-=·.\s]+", l):
        return True
    return False


def _is_num(s: str) -> bool:
    return bool(re.fullmatch(r"\(?-?[\d,]+(\.\d+)?\)?", s)) and any(c.isdigit() for c in s)


def _to_num(s: str):
    s = s.strip()
    if s in ("-", "–", "—", ""):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[(),]", "", s)
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _clean_label(s: str) -> str:
    s = s.replace("|", " ")
    s = re.sub(r"\(주석[^)]*\)", "", s)
    s = re.sub(r"^[IVXⅠ-Ⅻ]+\.\s*", "", s)         # 로마숫자 머리표
    s = re.sub(r"^[0-9]+\.\s*", "", s)             # 숫자 머리표
    s = re.sub(r"^\([0-9]+\)\s*", "", s)           # (1) 머리표
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _detect_years(text: str):
    head = "\n".join(text.splitlines()[:80])
    ys = re.findall(r"(20\d\d)\s*년\s*12\s*월\s*31\s*일", head)
    if len(ys) < 2:
        ys = re.findall(r"(20\d\d)\s*년", head)
    seen = []
    for y in ys:
        if y not in seen:
            seen.append(y)
        if len(seen) == 2:
            break
    if len(seen) < 2:
        return None
    cur, prev = int(seen[0]), int(seen[1])
    if cur < prev:
        cur, prev = prev, cur
    return cur, prev


def _section_text(text: str, starts, stops):
    """실제 표 구간을 슬라이스 : 목차(숫자 적음)는 건너뛰고 숫자가 많은 첫 등장을 채택"""
    best = -1
    for s in starts:
        idx = 0
        while True:
            i = text.find(s, idx)
            if i == -1:
                break
            window = text[i:i + 600]
            # 실제 재무제표 본문에만 있는 '콤마 단위 큰 숫자'(백만 이상)로 목차와 구분
            if len(re.findall(r"\d{1,3}(?:,\d{3}){2,}", window)) >= 2:
                best = i
                break
            idx = i + len(s)
        if best != -1:
            break
    if best == -1:
        return ""
    hi = len(text)
    for s in stops:
        j = text.find(s, best + 10)
        if j != -1:
            hi = min(hi, j)
    return text[best:hi]


def _parse_section(section: str):
    lines = [l.strip() for l in section.splitlines() if l.strip()]
    out = []
    label, nums = None, []

    def flush():
        nonlocal label, nums
        if label and nums:
            out.append((label, nums[0], nums[1] if len(nums) > 1 else None))
        label, nums = None, []

    for l in lines[1:]:                       # 첫 줄(섹션 제목) 건너뜀
        if _is_num(l):
            nums.append(_to_num(l))
            continue
        if l in ("-", "–", "—"):
            nums.append(0.0)
            continue
        if _is_noise(l):
            continue
        # 텍스트 라벨
        if label and nums:                    # 직전 계정 마감
            flush()
        lc = _clean_label(l)
        if not lc or lc in ("자 산", "부 채", "자 본", "자산", "부채", "자본"):
            continue
        if label is None:
            label = lc
        elif lc.startswith("(") or "주석" in l:
            label = (label + " " + lc).strip()
        else:
            label = lc                        # 헤더성 텍스트는 실제 계정으로 교체
    flush()
    return out


def parse_report(txt_path: Path):
    text = txt_path.read_text(encoding="utf-8")
    yrs = _detect_years(text)
    if not yrs:
        return None
    cur, prev = yrs
    data = {"years": (cur, prev), "BS": {}, "IS": {}, "CF": {}}
    for key, starts, stops in SECTIONS:
        sec = _section_text(text, starts, stops)
        for label, v_cur, v_prev in _parse_section(sec):
            data[key].setdefault(label, {})
            if label not in data[key] or cur not in data[key][label]:
                data[key][label][cur] = v_cur
            if v_prev is not None:
                data[key][label][prev] = v_prev
    return data


def parse_all(parsed_dir: Path = PARSED):
    """01_parsed/DART_*.txt 전부 파싱해 다개년 통합"""
    merged = {"BS": {}, "IS": {}, "CF": {}, "years": set()}
    order = {"BS": [], "IS": [], "CF": []}
    reports = sorted(parsed_dir.glob("DART_*.txt"))
    for rp in reports:
        d = parse_report(rp)
        if not d:
            continue
        merged["years"].update(d["years"])
        for st in ("BS", "IS", "CF"):
            for label, yv in d[st].items():
                if label not in merged[st]:
                    merged[st][label] = {}
                    order[st].append(label)
                merged[st][label].update(yv)
    merged["years"] = sorted(merged["years"])
    merged["order"] = order
    return merged


if __name__ == "__main__":
    m = parse_all()
    print("연도:", m["years"])
    for st in ("IS", "BS", "CF"):
        print(f"\n===== {st} (계정 {len(m[st])}개) =====")
        for label in m["order"][st][:40]:
            yv = m[st][label]
            vals = "  ".join(f"{y}:{(yv.get(y)/1e8):,.1f}" if yv.get(y) is not None else f"{y}:-" for y in m["years"])
            print(f"  {label[:24]:24s} {vals}")
