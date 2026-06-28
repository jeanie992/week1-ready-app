# -*- coding: utf-8 -*-
"""
dart_fetch.py — DART(전자공시)에서 회사의 정기보고서(감사보고서/사업보고서)를 받아 텍스트로 추출.

비상장 외부감사 대상 회사는 보통 '감사보고서'가 올라온다(상장사는 '사업보고서').
받은 원본은 00_inbox/dart/ 에, 추출 텍스트는 01_parsed/DART_*.txt 로 저장 → 기존 파이프라인에 합류.
실제 해석·정리는 Claude 가 추출 텍스트를 읽고 수행한다.

사용:
    python scripts/dart_fetch.py "회사명" --years 3
API 키:
    환경변수 DART_API_KEY, 또는 프로젝트 루트의 .dart_key 파일에서 읽음. (외부 공유 금지)
표준 라이브러리만 사용(추가 설치 불필요).
"""
from __future__ import annotations
import io
import os
import re
import sys
import json
import html
import zipfile
import argparse
import urllib.parse
import urllib.request
from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
DART_RAW = ROOT / "00_inbox" / "dart"
PARSED = ROOT / "01_parsed"
CACHE = ROOT / "scripts" / ".corpcode_cache.xml"
BASE = "https://opendart.fss.or.kr/api"


def load_key() -> str:
    key = os.environ.get("DART_API_KEY", "").strip()
    if not key:
        f = ROOT / ".dart_key"
        if f.exists():
            key = f.read_text(encoding="utf-8").strip()
    if not key:
        sys.exit("[오류] DART API 키가 없습니다. 환경변수 DART_API_KEY 또는 프로젝트 루트 .dart_key 파일을 두세요.")
    return key


def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "week1-ready/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def get_corp_code(key: str, name: str) -> str:
    """회사명 → corp_code(8자리). 전체 목록(zip)을 받아 캐시 후 매칭."""
    if not CACHE.exists():
        print("  corpCode 전체 목록 내려받는 중(최초 1회)…")
        data = http_get(f"{BASE}/corpCode.xml?crtfc_key={key}")
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            CACHE.write_bytes(z.read("CORPCODE.xml"))
    root = ET.fromstring(CACHE.read_text(encoding="utf-8"))
    exact, contains = [], []
    for el in root.iter("list"):
        cn = (el.findtext("corp_name") or "").strip()
        cc = (el.findtext("corp_code") or "").strip()
        if cn == name:
            exact.append((cn, cc))
        elif name in cn:
            contains.append((cn, cc))
    cands = exact or contains
    if not cands:
        sys.exit(f"[오류] '{name}' 에 해당하는 회사를 DART 목록에서 못 찾음.")
    if len(cands) > 1:
        print(f"  후보 {len(cands)}개: " + ", ".join(f"{n}({c})" for n, c in cands[:8]))
    print(f"  선택: {cands[0][0]} (corp_code={cands[0][1]})")
    return cands[0][1]


def list_reports(key: str, corp_code: str, bgn: str, end: str) -> list[dict]:
    """정기보고서/감사보고서 공시 목록. 감사보고서·사업보고서만 필터."""
    out = []
    for ty in ("F", "A"):  # F=외부감사관련(감사보고서), A=정기공시(사업보고서)
        url = (f"{BASE}/list.json?crtfc_key={key}&corp_code={corp_code}"
               f"&bgn_de={bgn}&end_de={end}&pblntf_ty={ty}&page_count=100&page_no=1")
        try:
            j = json.loads(http_get(url).decode("utf-8"))
        except Exception as e:
            print(f"  목록 조회 실패(ty={ty}): {e}", file=sys.stderr)
            continue
        if j.get("status") != "000":
            print(f"  목록 응답(ty={ty}): {j.get('status')} {j.get('message')}")
            continue
        for it in j.get("list", []):
            nm = it.get("report_nm", "")
            if ("감사보고서" in nm) or ("사업보고서" in nm):
                out.append(it)
    return out


def pick_annual(reports: list[dict], years: int) -> list[dict]:
    """연도별 1건씩, 최신순 N개. 연결보다 별도/개별 우선, 정정은 최신만."""
    def score(it):
        nm = it["report_nm"]
        s = 0
        if "연결" in nm:
            s -= 1            # 별도/개별 우선
        if "사업보고서" in nm:
            s += 2            # 사업보고서가 있으면 우선(상장사)
        return s
    by_year: dict[str, dict] = {}
    for it in sorted(reports, key=lambda x: x.get("rcept_dt", ""), reverse=True):
        y = it.get("rcept_dt", "")[:4]
        if not y:
            continue
        if y not in by_year or score(it) > score(by_year[y]):
            # 같은 해 더 적합한 보고서로 교체(이미 최신 정렬이라 날짜는 충분)
            by_year.setdefault(y, it)
    chosen = [by_year[y] for y in sorted(by_year, reverse=True)[:years]]
    return chosen


def xml_to_text(raw: bytes) -> str:
    """DART 문서 XML(zip 내부) → 대략적 평문. 태그 제거 + 엔티티 복원."""
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            s = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        s = raw.decode("utf-8", errors="ignore")
    s = re.sub(r"(?is)<(script|style).*?</\1>", " ", s)
    s = re.sub(r"(?i)</(p|tr|div|title|table|br|li)\s*>", "\n", s)
    s = re.sub(r"(?i)<td[^>]*>", " | ", s)
    s = re.sub(r"<[^>]+>", " ", s)          # 나머지 태그 제거
    s = html.unescape(s)
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n\s*\n\s*\n+", "\n\n", s)
    return s.strip()


def fetch_document(key: str, rcept_no: str) -> bytes:
    return http_get(f"{BASE}/document.xml?crtfc_key={key}&rcept_no={rcept_no}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("company", nargs="?", default="")
    ap.add_argument("--years", type=int, default=3)
    ap.add_argument("--corp_code", default=None)
    ap.add_argument("--bgn", default=None, help="시작일 YYYYMMDD (기본: years+1년 전)")
    args = ap.parse_args()

    key = load_key()
    DART_RAW.mkdir(parents=True, exist_ok=True)
    PARSED.mkdir(parents=True, exist_ok=True)

    corp_code = args.corp_code or get_corp_code(key, args.company)
    # 보고서 제출은 회계연도 다음 해 → 넉넉히 (years+1)년 범위
    from datetime import date  # date.today()는 환경상 사용 불가할 수 있어 sys로 대체
    end = "".join(filter(str.isdigit, os.environ.get("TODAY", "")))[:8]
    if len(end) != 8:
        # TODAY 미지정 시 넉넉한 상한
        end = "20991231"
    bgn = args.bgn or f"{max(2000, int((end[:4] if end[:4].isdigit() else '2026')) - (args.years + 1))}0101"

    print(f"[1/3] 공시 목록 조회 ({bgn}~{end})")
    reports = list_reports(key, corp_code, bgn, end)
    if not reports:
        sys.exit("[오류] 감사보고서/사업보고서를 찾지 못했습니다. 회사명/기간을 확인하세요.")
    chosen = pick_annual(reports, args.years)
    print(f"[2/3] 대상 {len(chosen)}건:")
    for it in chosen:
        print(f"   - {it['rcept_dt']}  {it['report_nm']}  (rcept_no={it['rcept_no']})")

    print("[3/3] 원본 다운로드 + 텍스트 추출")
    manifest = []
    for it in chosen:
        rno = it["rcept_no"]
        nm = re.sub(r"[^0-9A-Za-z가-힣]", "_", it["report_nm"])[:40]
        year = it["rcept_dt"][:4]
        try:
            blob = fetch_document(key, rno)
            (DART_RAW / f"{rno}.zip").write_bytes(blob)
            texts = []
            with zipfile.ZipFile(io.BytesIO(blob)) as z:
                for n in z.namelist():
                    if n.lower().endswith((".xml", ".html", ".htm", ".txt")):
                        texts.append(xml_to_text(z.read(n)))
            text = "\n\n".join(t for t in texts if t)
            out = PARSED / f"DART_{year}_{nm}.txt"
            out.write_text(f"# 출처: DART {it['report_nm']} (접수 {it['rcept_dt']}, rcept_no={rno})\n\n{text}",
                           encoding="utf-8")
            manifest.append((out.name, len(text)))
            print(f"   OK {out.name}  ({len(text):,}자)")
        except Exception as e:
            print(f"   XX {rno} 실패: {e}", file=sys.stderr)

    print(f"\n완료: {len(manifest)}건 추출 → {PARSED}")
    print("원본 보관:", DART_RAW)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
