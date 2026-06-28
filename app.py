# -*- coding: utf-8 -*-
"""
app.py : Week 1 현황파악 도구 (Streamlit, 로컬 단독)

실행 : 가상환경 활성화 후 streamlit run app.py  (또는 .venv/Scripts/streamlit.exe run app.py)
역할 : 자료를 끌어다 넣으면 기존 스크립트(convert·analyze_excel·dart_fetch·build_master)를 실행하고
       결과(추출 상태·Excel 분석·DART 공시·마스터 엑셀)를 보여준다
문체 : 끝 온점 없음, 하이픈 대신 콜론
"""
from __future__ import annotations
import os, sys, subprocess
from pathlib import Path
import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
INBOX = ROOT / "00_inbox"
PARSED = ROOT / "01_parsed"
ANALYSIS = ROOT / "02_analysis"
OUTPUT = ROOT / "03_output"
PY = ROOT / ".venv" / "Scripts" / "python.exe"
PYEXE = str(PY) if PY.exists() else sys.executable

# 확장자 → 00_inbox 하위 폴더
EXT_MAP = {
    ".xlsx": "excel", ".xls": "excel", ".csv": "excel",
    ".pdf": "pdf", ".docx": "word", ".pptx": "ppt",
    ".hwp": "hwp", ".hwpx": "hwp",
}

st.set_page_config(page_title="Week1 현황파악", layout="wide")


def run(args, label, env=None):
    """venv 파이썬으로 스크립트 실행 후 stdout 반환"""
    with st.status(label, expanded=False) as box:
        p = subprocess.run([PYEXE, *args], cwd=str(ROOT),
                           capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
        st.code((p.stdout or "") + (p.stderr or ""), language="text")
        box.update(state="complete" if p.returncode == 0 else "error")
    return p.returncode == 0


def inbox_summary():
    rows = []
    for sub in ["excel", "pdf", "word", "ppt", "hwp"]:
        d = INBOX / sub
        n = len([f for f in d.iterdir() if f.is_file()]) if d.exists() else 0
        rows.append({"형식": sub, "파일수": n})
    return pd.DataFrame(rows)


# ── 사이드바 : 작업공간 ──────────────────────────────────────────────────────
with st.sidebar:
    st.header("작업공간")
    company = st.text_input("회사명", value="", placeholder="예 : 회사명", help="DART 검색·표기에 사용")
    st.dataframe(inbox_summary(), hide_index=True, use_container_width=True)
    if st.button("결과 초기화 (01~03 비우기)", type="secondary"):
        import shutil
        for d in [PARSED, ANALYSIS, OUTPUT]:
            if d.exists():
                shutil.rmtree(d)
            d.mkdir(parents=True, exist_ok=True)
        st.success("01_parsed · 02_analysis · 03_output 비움")
    st.caption("원본(00_inbox)은 유지됩니다")

st.title("Week 1 현황파악 도구")
st.caption("자료를 넣고 → 분석 실행 → 마스터 엑셀을 받습니다")

# ── 1. 자료 업로드 ───────────────────────────────────────────────────────────
st.subheader("1. 자료 업로드")
st.caption("엑셀·PDF·Word·PPT·HWP 를 한 번에 끌어다 놓으세요 : 확장자로 자동 분류됩니다")
ups = st.file_uploader("파일 선택", accept_multiple_files=True,
                       type=["xlsx", "xls", "csv", "pdf", "docx", "pptx", "hwp", "hwpx"])
if ups:
    saved = []
    for f in ups:
        ext = Path(f.name).suffix.lower()
        cat = EXT_MAP.get(ext)
        if not cat:
            continue
        dest = INBOX / cat
        dest.mkdir(parents=True, exist_ok=True)
        (dest / f.name).write_bytes(f.getbuffer())
        saved.append({"파일": f.name, "분류": cat, "크기(KB)": round(len(f.getbuffer()) / 1024, 1)})
    if saved:
        st.success(f"{len(saved)}개 저장됨")
        st.dataframe(pd.DataFrame(saved), hide_index=True, use_container_width=True)
with st.expander("엑셀 입력 규약 (자동 분석 적중률↑)"):
    st.markdown("첫 행 = 연/분기 헤더 : 첫 열 = 항목명 : 한 시트에 표 하나 : 병합셀 지양")

# ── 2. 웹 기사 URL ───────────────────────────────────────────────────────────
st.subheader("2. 웹 기사 URL (선택)")
urls_path = INBOX / "urls.txt"
cur = urls_path.read_text(encoding="utf-8") if urls_path.exists() else ""
urls = st.text_area("줄당 하나, '#' 은 주석", value=cur, height=100)
if st.button("URL 저장"):
    INBOX.mkdir(parents=True, exist_ok=True)
    urls_path.write_text(urls, encoding="utf-8")
    st.success("urls.txt 저장")

# ── 3. DART 공시 ─────────────────────────────────────────────────────────────
st.subheader("3. DART 공시 (선택)")
c1, c2 = st.columns([2, 2])
dart_company = c1.text_input("DART 회사명", value=company)
dart_key = c2.text_input("DART API 키", type="password",
                         help=".dart_key 파일이 있으면 비워두세요")
if dart_key:
    (ROOT / ".dart_key").write_text(dart_key.strip(), encoding="utf-8")
anthropic_key = st.text_input("Anthropic API 키 (NOTE 인사이트, 선택)", type="password",
                              help="비우면 : 로컬은 Claude Code(Max 구독)로, 클라우드는 규칙 기반으로 NOTE 생성")

# ── 4. 분석 실행 ─────────────────────────────────────────────────────────────
st.subheader("4. 분석 실행")
b1, b2, b3, b4 = st.columns(4)
if b1.button("문서 추출", use_container_width=True):
    run(["scripts/convert.py"], "문서 → 텍스트 추출")
if b2.button("Excel 분석", use_container_width=True):
    run(["scripts/analyze_excel.py"], "Excel 시계열·비교 분석")
if b3.button("DART 수집", use_container_width=True):
    if dart_company:
        run(["scripts/dart_fetch.py", dart_company, "--years", "3"], f"DART {dart_company} 3개년 수집")
    else:
        st.warning("DART 회사명을 입력하세요")
if b4.button("전체 실행", type="primary", use_container_width=True):
    ok = run(["scripts/convert.py"], "1/3 문서 추출")
    if dart_company:
        run(["scripts/dart_fetch.py", dart_company, "--years", "3"], "2/3 DART 수집")
    run(["scripts/analyze_excel.py"], "3/3 Excel 분석")
    st.success("완료 : 아래 결과 확인")

st.markdown("**재무 마스터 자동 생성 (모든 회사 : DART 3표 파싱 + NOTE 인사이트)**")
st.caption("DART 수집이 끝난 뒤 실행 : Anthropic 키 있으면 Claude 인사이트, 없으면 규칙 기반")
if st.button("DART → 재무 마스터 자동 생성", type="primary", use_container_width=True):
    env = os.environ.copy()
    args = ["scripts/build_master_generic.py", dart_company or company]
    if anthropic_key:
        env["ANTHROPIC_API_KEY"] = anthropic_key.strip()
        args.append("--llm")
    run(args, "재무 마스터 자동 생성 (파싱 + 인사이트)", env=env)

# ── 5. 결과 ──────────────────────────────────────────────────────────────────
st.subheader("5. 결과")
man = PARSED / "_manifest.csv"
if man.exists():
    st.markdown("**문서 추출 상태**")
    st.dataframe(pd.read_csv(man), hide_index=True, use_container_width=True)

digest = ANALYSIS / "digest.md"
scan = ANALYSIS / "_scan.md"
if digest.exists():
    with st.expander("Excel 분석 요약 (digest)", expanded=True):
        st.markdown(digest.read_text(encoding="utf-8"))
elif scan.exists():
    with st.expander("Excel 구조 스캔"):
        st.markdown(scan.read_text(encoding="utf-8"))

dart_dir = PARSED
dart_files = sorted(dart_dir.glob("DART_*.txt")) if dart_dir.exists() else []
if dart_files:
    st.markdown("**DART 공시 추출** : " + ", ".join(f.stem for f in dart_files))

st.markdown("**다운로드**")
d1, d2 = st.columns(2)
ax = ANALYSIS / "analysis.xlsx"
if ax.exists():
    d1.download_button("Excel 분석표", ax.read_bytes(),
                       file_name="analysis.xlsx", use_container_width=True)
gen = OUTPUT / "재무마스터_자동.xlsx"
if gen.exists():
    d2.download_button("재무 마스터 (자동)", gen.read_bytes(),
                       file_name="재무마스터_자동.xlsx", use_container_width=True)
