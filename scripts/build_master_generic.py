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


# ── 내부 Excel(사업부 P&L) → 추가 분석 탭 ───────────────────────────────────
PPF = '+0.0"%p";[Red]-0.0"%p";"-"'


def _val(tidy, ent, cat, y):
    s = tidy[(tidy["entity"] == ent) & (tidy["category"] == cat) & (tidy["year"] == y)]["value"]
    return float(s.sum()) if len(s) else None


def _find_entity(ents, keys):
    for k in keys:
        for e in ents:
            if str(e).strip() == k:
                return e
    for k in keys:
        for e in ents:
            if k in str(e):
                return e
    return None


def _iwidths(ws, ny):
    ws.column_dimensions["A"].width = S.W_MARGIN
    ws.column_dimensions["B"].width = 22
    for j in range(ny + 2):
        ws.column_dimensions[get_column_letter(3 + j)].width = 10
    ws.sheet_view.showGridLines = False


def _seg_pl_tab(wb, tidy, years, cats, rev, oi):
    ws = wb.create_sheet("4.사업부손익")
    S.label(ws.cell(2, 2), f"사업부별 {rev}·{oi}·이익률 ({years[0]}~{years[-1]})", bold=True)
    ws.cell(3, 2).value = S.clean_text("단위 : 억원 (이익률 %) : 보조 CAGR·YoY(비율 %p) : 출처 내부 P&L")
    ws.cell(3, 2).font = S.FONT_NOTE
    hr = 4
    S.header(ws.cell(hr, 2), "사업부 / 지표")
    for j, y in enumerate(years):
        S.header(ws.cell(hr, 3 + j), str(y))
    base = 3 + len(years)
    S.header(ws.cell(hr, base), "CAGR"); S.header(ws.cell(hr, base + 1), "YoY")
    totlike = [c for c in cats if str(c).strip() in ("Total", "합계", "계", "총계")]
    order = [c for c in cats if c not in totlike] + totlike
    r = hr + 1
    for cat in order:
        isT = cat in totlike
        for metric, lm in [(rev, "매출"), (oi, "영업이익"), ("__m__", "영업이익률")]:
            S.label(ws.cell(r, 2), f"{cat} : {lm}", bold=isT)
            vals = []
            for j, y in enumerate(years):
                if metric == "__m__":
                    rv, ov = _val(tidy, rev, cat, y), _val(tidy, oi, cat, y)
                    v = (ov / rv) if (rv and ov is not None) else None
                    S.num(ws.cell(r, 3 + j), v, PCT); vals.append(v)
                else:
                    v = _val(tidy, metric, cat, y)
                    vv = (v / 1e8) if v is not None else None
                    S.num(ws.cell(r, 3 + j), vv, EOK, bold=isT); vals.append(vv)
            a = next((x for x in vals if x is not None), None)
            c2 = vals[-1]; b = vals[-2] if len(vals) >= 2 else None
            n = sum(1 for x in vals if x is not None) - 1
            if metric == "__m__":
                S.num(ws.cell(r, base), ((c2 - a) * 100 if (a is not None and c2 is not None) else None), PPF)
                S.num(ws.cell(r, base + 1), ((c2 - b) * 100 if (b is not None and c2 is not None) else None), PPF)
            else:
                cagr = ((c2 / a) ** (1 / n) - 1) if (a and a > 0 and c2 and c2 > 0 and n >= 1) else None
                yoy = (c2 / b - 1) if (b and b > 0 and c2 is not None) else None
                S.num(ws.cell(r, base), cagr, GROW); S.num(ws.cell(r, base + 1), yoy, GROW)
            if isT:
                for cc in range(2, base + 2):
                    ws.cell(r, cc).fill = S.FILL_BAND
            r += 1
    _iwidths(ws, len(years))


def _commonsize_tab(wb, tidy, years, totcat, rev):
    ws = wb.create_sheet("5.비용구조_매출대비")
    S.label(ws.cell(2, 2), f"공통비율 : 매출 100원당 항목 비중 ({totcat})", bold=True)
    ws.cell(3, 2).value = S.clean_text("단위 : % : 보조 Δ(첫→끝, %p) : 출처 내부 P&L")
    ws.cell(3, 2).font = S.FONT_NOTE
    show = years if len(years) <= 6 else [years[0], years[len(years) // 2], years[-3], years[-2], years[-1]]
    hr = 4
    S.header(ws.cell(hr, 2), "항목")
    for j, y in enumerate(show):
        S.header(ws.cell(hr, 3 + j), str(y))
    S.header(ws.cell(hr, 3 + len(show)), "Δ %p")
    ents = list(tidy[tidy["category"] == totcat]["entity"].unique())
    ents.sort(key=lambda e: abs(_val(tidy, e, totcat, years[-1]) or 0), reverse=True)
    r = hr + 1
    for e in ents:
        S.label(ws.cell(r, 2), e, bold=(e == rev))
        ratios = []
        for j, y in enumerate(show):
            ev, rvy = _val(tidy, e, totcat, y), _val(tidy, rev, totcat, y)
            rt = (ev / rvy) if (ev is not None and rvy) else None
            S.num(ws.cell(r, 3 + j), rt, PCT); ratios.append(rt)
        d = ((ratios[-1] - ratios[0]) * 100) if (ratios and ratios[0] is not None and ratios[-1] is not None) else None
        S.num(ws.cell(r, 3 + len(show)), d, PPF)
        if e == rev:
            for cc in range(2, 4 + len(show)):
                ws.cell(r, cc).fill = S.FILL_BAND
        r += 1
    _iwidths(ws, len(show))


def internal_tabs(wb):
    """내부 00_inbox/excel 의 P&L 을 분석해 사업부손익·비용구조 탭 추가 : 반환 추가탭 목록"""
    try:
        import analyze_excel as AX
        sheets = list(AX.load_tables())
    except Exception:
        return []
    added = []
    for fname, sheet, raw in sheets:
        try:
            tidy, note, warns = AX.to_tidy(raw)
        except Exception:
            continue
        if tidy is None or tidy.empty:
            continue
        years = sorted(int(y) for y in tidy["year"].unique())
        cats = [c for c in tidy["category"].unique() if str(c).strip()]
        ents = list(tidy["entity"].unique())
        rev = _find_entity(ents, ["매출", "영업수익", "매출액", "수익"])
        oi = _find_entity(ents, ["영업이익"])
        totcat = next((c for c in cats if str(c).strip() in ("Total", "합계", "계", "총계")), (cats[0] if cats else None))
        if rev and oi and len(cats) >= 2 and len(years) >= 2:
            _seg_pl_tab(wb, tidy, years, cats, rev, oi)
            if totcat:
                _commonsize_tab(wb, tidy, years, totcat, rev)
            added += ["4.사업부손익", "5.비용구조_매출대비"]
            break
    return added


def main():
    company = next((a for a in sys.argv[1:] if not a.startswith("-")), "회사")
    use_llm = "--no-llm" not in sys.argv   # 기본 : LLM 시도(API 키 없으면 로컬 Claude Code=Max)
    m = dart_parse.parse_all()
    years = m["years"]
    wb = Workbook(); wb.remove(wb.active)
    notes_all = {}
    if years:
        notes_all = llm_notes.generate(m, years, company=company, use_llm=use_llm)
        notes = {st: {acc: nt for (s2, acc), nt in notes_all.items() if s2 == st} for st in ("IS", "BS", "CF")}
        cover(wb, company, years, use_llm and bool(notes_all))
        statement_sheet(wb, "1.손익계산서", f"{company} 손익계산서 (억원)", years, m["order"]["IS"], m["IS"], notes["IS"])
        statement_sheet(wb, "2.재무상태표", f"{company} 재무상태표 (억원)", years, m["order"]["BS"], m["BS"], notes["BS"])
        statement_sheet(wb, "3.현금흐름표", f"{company} 현금흐름표 (억원)", years, m["order"]["CF"], m["CF"], notes["CF"])
    else:
        print("[안내] DART 공시 없음 → 내부 Excel 분석만 진행")

    internal = internal_tabs(wb)   # 내부 P&L → 사업부손익·비용구조 탭

    if not wb.sheetnames:
        sys.exit("[오류] 분석할 자료가 없습니다 (DART 또는 내부 Excel 필요)")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print("저장 :", OUT)
    print(f"탭 : {wb.sheetnames}")
    print(f"  DART 연도 {years or '없음'} : 노트 {len(notes_all)}건 : 내부탭 {internal or '없음'}")


if __name__ == "__main__":
    main()
