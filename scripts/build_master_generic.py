# -*- coding: utf-8 -*-
"""
build_master_generic.py : 어떤 회사든 DART 텍스트에서 3개 재무제표를 자동 파싱해 마스터 엑셀 생성

흐름 : dart_parse(파싱) → llm_notes(인사이트) → excel_style(양식) 렌더
산출 : 03_output/재무마스터_자동.xlsx (탭 : Cover / 1.손익 / 2.재무상태 / 3.현금흐름)
인사이트 : ANTHROPIC_API_KEY 있으면 Claude, 없으면 규칙 기반
사용 : python scripts/build_master_generic.py "회사명" [--llm]
문체 : 끝 온점 없음, 하이픈 대신 콜론
"""
from __future__ import annotations
import sys, re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

import excel_style as S
import dart_parse
import llm_notes

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "03_output" / "재무마스터_자동.xlsx"
EOK = "#,##0.0;[Red]\\-#,##0.0;\\-;@"
PCT = "0.0%"
GROW = '+0.0%;[Red]-0.0%;"-"'
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

SUBTOTAL_KW = ("총계", "영업수익", "영업비용", "영업이익", "당기순이익", "현금흐름",
               "순이익", "법인세비용차감전", "포괄손익", "유동자산", "비유동자산",
               "유동부채", "비유동부채", "자본금", "이익잉여금")


def is_sub(label):
    return any(k in label.replace(" ", "") for k in [k.replace(" ", "") for k in SUBTOTAL_KW])


def _norm(s):
    return re.sub(r"\s+", "", s)


def _match_note(label, notes_for_stmt):
    n = _norm(label)
    for acc, note in notes_for_stmt.items():
        if _norm(acc) == n or _norm(acc) in n or n in _norm(acc):
            return note
    return ""


def eok(v):
    return v / 1e8 if v is not None else None


def supp(vals):
    """YoY(마지막), CAGR(처음→마지막)"""
    vals2 = [v for v in vals]
    a = next((v for v in vals2 if v is not None), None)
    c = next((v for v in reversed(vals2) if v is not None), None)
    b = vals2[-2] if len(vals2) >= 2 else None
    n = sum(1 for v in vals2 if v is not None) - 1
    yoy = (c / b - 1) if (b and b > 0 and c is not None) else None
    cagr = ((c / a) ** (1 / n) - 1) if (a and a > 0 and c and c > 0 and n >= 1) else None
    return yoy, cagr


def statement_sheet(wb, name, title, years, accounts, data, notes):
    ws = wb.create_sheet(name)
    S.label(ws.cell(2, 2), title, bold=True)
    ws.cell(3, 2).value = S.clean_text("단위 : 억원 : 보조 YoY·CAGR : 출처 DART 감사보고서")
    ws.cell(3, 2).font = S.FONT_NOTE
    hr = 4
    heads = ["항목"] + [str(y) for y in years] + [f"YoY '{str(years[-1])[2:]}", f"CAGR '{str(years[0])[2:]}-'{str(years[-1])[2:]}", "NOTE"]
    for j, h in enumerate(heads):
        S.header(ws.cell(hr, 2 + j), h)
    ncol = len(years)
    r = hr + 1
    for acc in accounts:
        vals = [eok(data[acc].get(y)) for y in years]
        if all(v is None for v in vals):
            continue
        sub = is_sub(acc)
        S.label(ws.cell(r, 2), acc, bold=sub)
        for j, v in enumerate(vals):
            S.num(ws.cell(r, 3 + j), v, EOK, bold=sub)
        yoy, cagr = supp(vals)
        S.num(ws.cell(r, 3 + ncol), yoy, GROW)
        S.num(ws.cell(r, 3 + ncol + 1), cagr, GROW)
        nc = ws.cell(r, 3 + ncol + 2)
        nc.value = S.clean_text(_match_note(acc, notes))
        nc.font = S.FONT_NOTE
        nc.alignment = WRAP
        nc.border = S.BORDER_THIN
        if sub:
            for cc in range(2, 3 + ncol + 2):
                ws.cell(r, cc).fill = S.FILL_BAND
        r += 1
    ws.column_dimensions["A"].width = S.W_MARGIN
    ws.column_dimensions["B"].width = 26
    for j in range(ncol + 2):
        ws.column_dimensions[get_column_letter(3 + j)].width = 11
    ws.column_dimensions[get_column_letter(3 + ncol + 2)].width = 50
    ws.sheet_view.showGridLines = False


def cover(wb, company, years, used_llm):
    ws = wb.create_sheet("Cover")
    S.header(ws.cell(2, 2), f"{company} 재무 마스터 (자동 생성)")
    meta = [("대상회사", company), ("기간", f"{years[0]}~{years[-1]}"),
            ("단위", "억원 ( DART 원 → ÷1e8 )"),
            ("출처", "DART 감사/사업보고서 자동 파싱(dart_parse)"),
            ("인사이트", "Claude API" if used_llm else "규칙 기반(YoY 급변) : ANTHROPIC_API_KEY 설정 시 Claude"),
            ("탭", "1.손익계산서 2.재무상태표 3.현금흐름표"),
            ("문체", "문장 끝 온점 없음 : 하이픈 대신 콜론")]
    r = 4
    for k, v in meta:
        S.label(ws.cell(r, 2), k, bold=True)
        ws.cell(r, 3).value = S.clean_text(v); ws.cell(r, 3).font = S.FONT_LABEL; ws.cell(r, 3).alignment = WRAP
        r += 1
    ws.column_dimensions["A"].width = S.W_MARGIN
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 72
    ws.sheet_view.showGridLines = False


def main():
    company = next((a for a in sys.argv[1:] if not a.startswith("-")), "회사")
    use_llm = "--no-llm" not in sys.argv   # 기본 : LLM 시도(API 키 없으면 로컬 Claude Code=Max)
    m = dart_parse.parse_all()
    if not m["years"]:
        sys.exit("[오류] 01_parsed 에 DART_*.txt 가 없습니다 (dart_fetch.py 먼저 실행)")
    years = m["years"]
    notes_all = llm_notes.generate(m, years, company=company, use_llm=use_llm)
    notes = {st: {acc: nt for (s2, acc), nt in notes_all.items() if s2 == st} for st in ("IS", "BS", "CF")}

    wb = Workbook(); wb.remove(wb.active)
    cover(wb, company, years, use_llm and bool(notes_all))
    statement_sheet(wb, "1.손익계산서", f"{company} 손익계산서 (억원)", years, m["order"]["IS"], m["IS"], notes["IS"])
    statement_sheet(wb, "2.재무상태표", f"{company} 재무상태표 (억원)", years, m["order"]["BS"], m["BS"], notes["BS"])
    statement_sheet(wb, "3.현금흐름표", f"{company} 현금흐름표 (억원)", years, m["order"]["CF"], m["CF"], notes["CF"])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print("저장 :", OUT)
    print(f"탭 : {wb.sheetnames} : 연도 {years} : 노트 {len(notes_all)}건 ({'Claude' if use_llm else '규칙기반'})")


if __name__ == "__main__":
    main()
