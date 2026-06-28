# -*- coding: utf-8 -*-
"""
analyze_excel.py  —  00_inbox/excel 의 Excel/CSV 를 시계열·항목별로 분석한다.

핵심 아이디어:
  스키마(표 모양)는 파일마다 다르다. 그래서 어떤 표든 먼저
  '긴 형식(tidy long)' = [entity_code, entity, category, year, value] 로 변환한 뒤,
  그 위에서 YoY / CAGR / 항목별 비교를 '똑같은 방식'으로 계산한다.

지원하는 표 모양:
  1) 멀티헤더 가로형 : 0행=연도(병합 반복), 1행=구분/지표, 왼쪽 몇 칸=항목(코드·이름), 각 행=항목
                       (예: '프로퍼티별 10개년 financial현황' — 컨설팅 재무자료 전형)
  2) 단일헤더 가로형 : 헤더 한 줄에 연도들이 가로로, 각 행=항목
  3) 세로형(long)    : 어떤 열이 연도/날짜, 나머지 숫자열이 지표

감지에 실패하면 '추측하지 않고' digest.md 에 "확인 필요"로 남기고 원본 미리보기만 싣는다.

출력:
  02_analysis/analysis.xlsx  (시트: 원본미리보기 / tidy / 시계열증감 / 항목비교)
  02_analysis/digest.md      (Claude /week1 이 읽을 자연어 요약)
"""
from __future__ import annotations
import re
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent.parent
EXCEL_DIR = ROOT / "00_inbox" / "excel"
OUT = ROOT / "02_analysis"
XLSX_OUT = OUT / "analysis.xlsx"
DIGEST_OUT = OUT / "digest.md"

YEAR_MIN, YEAR_MAX = 1990, 2099


# ── 값 정제 ────────────────────────────────────────────────────────────────
def parse_number(x):
    """'1,234' '(1,234)'(음수) '12.3%' '₩1,000' 등을 float 로. 불가면 NaN."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "" or s in {"-", "–", "—", "N/A", "n/a", "na", "NaN"}:
        return np.nan
    neg = s.startswith("(") and s.endswith(")")
    pct = s.endswith("%")
    s = re.sub(r"[(),%₩$€£¥\s]", "", s)
    s = s.replace("−", "-")  # 유니코드 마이너스
    try:
        v = float(s)
    except ValueError:
        return np.nan
    if neg:
        v = -v
    if pct:
        v = v / 100.0
    return v


def year_token(x):
    """셀 값에서 연도(int)를 추출. 연/FY/분기 시작연도 등. 아니면 None."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, float)) and float(x).is_integer():
        y = int(x)
        return y if YEAR_MIN <= y <= YEAR_MAX else None
    s = str(x).strip()
    m = re.search(r"(19|20)\d{2}", s)            # 2015, 2024-03, FY2024
    if m:
        return int(m.group(0))
    m = re.fullmatch(r"FY ?'?(\d{2})", s, re.I)   # FY23, FY'23
    if m:
        return 2000 + int(m.group(1))
    m = re.fullmatch(r"'?(\d{2})", s)             # '23
    if m:
        yy = int(m.group(1))
        return 2000 + yy if yy < 80 else 1900 + yy
    return None


def is_texty(series: pd.Series) -> float:
    """비어있지 않은 셀 중 '숫자로 못 읽는' 비율(=텍스트성). 0~1."""
    vals = [v for v in series.tolist() if not (v is None or (isinstance(v, float) and pd.isna(v)))]
    if not vals:
        return 0.0
    non_num = sum(1 for v in vals if pd.isna(parse_number(v)))
    return non_num / len(vals)


# ── 레이아웃 감지 + tidy 변환 ───────────────────────────────────────────────
def detect_year_row(raw: pd.DataFrame, scan=6):
    """앞쪽 행들 중 연도 토큰이 가장 많은 행 index. (행idx, 연도개수) 또는 (None,0)."""
    best, best_n = None, 0
    for r in range(min(scan, len(raw))):
        n = sum(year_token(v) is not None for v in raw.iloc[r].tolist())
        if n > best_n:
            best, best_n = r, n
    return best, best_n


def detect_year_col(raw: pd.DataFrame, scan=6):
    """앞쪽 열들 중 연도 토큰이 가장 많은 열 index (세로형 감지용)."""
    best, best_n = None, 0
    for c in range(min(scan, raw.shape[1])):
        n = sum(year_token(v) is not None for v in raw.iloc[:, c].tolist())
        if n > best_n:
            best, best_n = c, n
    return best, best_n


def to_tidy(raw: pd.DataFrame):
    """raw(header=None) → (tidy_df, 설명문자열, 경고리스트). 실패 시 tidy_df=None."""
    notes, warns = [], []
    year_row, n_row = detect_year_row(raw)
    year_col, n_col = detect_year_col(raw)

    # ── 가로형(연도가 행에 다수) ──────────────────────────────────────────
    if n_row >= 2 and n_row >= n_col:
        yrow = raw.iloc[year_row]
        years_ffill = yrow.copy()
        # 병합셀로 비어있는 연도 칸을 좌→우로 채움
        last = None
        filled = []
        for v in years_ffill.tolist():
            y = year_token(v)
            if y is not None:
                last = y
            filled.append(last)
        data_cols = [c for c in range(raw.shape[1]) if filled[c] is not None]
        if not data_cols:
            return None, "연도 열을 찾지 못함", ["가로형 추정했으나 연도 데이터열 없음"]
        first_data = min(data_cols)
        label_cols = [c for c in range(first_data) if c < raw.shape[1]]

        # 아래 행이 '텍스트 위주'면 멀티헤더(구분/지표 행), 숫자 위주면 단일헤더
        below = year_row + 1
        multi = below < len(raw) and is_texty(raw.iloc[below, data_cols]) >= 0.5
        if multi:
            cat_row = below
            data_start = cat_row + 1
            categories = {c: raw.iat[cat_row, c] for c in data_cols}
            notes.append(f"멀티헤더 가로형: {year_row}행=연도, {cat_row}행=구분, 좌측 {len(label_cols)}칸=항목")
        else:
            cat_row = None
            data_start = year_row + 1
            categories = {c: "" for c in data_cols}
            notes.append(f"단일헤더 가로형: {year_row}행=연도, 좌측 {len(label_cols)}칸=항목")

        # 항목(엔티티) 열: 숫자 위주면 코드, 텍스트 위주면 이름
        code_col = name_col = None
        for c in label_cols:
            col_below = raw.iloc[data_start:, c]
            if is_texty(col_below) >= 0.5:
                name_col = c if name_col is None else name_col
            else:
                code_col = c if code_col is None else code_col
        if name_col is None:
            name_col = label_cols[-1] if label_cols else None

        records = []
        for r in range(data_start, len(raw)):
            name = raw.iat[r, name_col] if name_col is not None else f"row{r}"
            if name is None or (isinstance(name, float) and pd.isna(name)):
                continue
            code = raw.iat[r, code_col] if code_col is not None else ""
            for c in data_cols:
                val = parse_number(raw.iat[r, c])
                if pd.isna(val):
                    continue
                records.append({
                    "entity_code": "" if pd.isna(code) else code,
                    "entity": str(name).strip(),
                    "category": str(categories[c]).strip() if categories[c] is not None else "",
                    "year": filled[c],
                    "value": val,
                })
        tidy = pd.DataFrame.from_records(records)
        if tidy.empty:
            return None, "; ".join(notes), ["데이터 행에서 숫자를 찾지 못함"]
        return tidy, "; ".join(notes), warns

    # ── 세로형(연도가 한 열에) ────────────────────────────────────────────
    if n_col >= 3:
        # 헤더 행 추정: 연도열 위쪽에서 텍스트 헤더가 있는 첫 행
        hdr = 0
        df = pd.read_excel  # noqa (자리표시) — 아래에서 raw 기반으로 직접 구성
        header_row = 0
        for r in range(min(6, len(raw))):
            if year_token(raw.iat[r, year_col]) is not None:
                header_row = max(0, r - 1)
                break
        cols = [str(raw.iat[header_row, c]).strip() for c in range(raw.shape[1])]
        body = raw.iloc[header_row + 1:].copy()
        body.columns = cols
        ycolname = cols[year_col]
        body["__year__"] = body[ycolname].map(year_token)
        body = body[body["__year__"].notna()]
        # 숫자열 = 값, 텍스트열 = 항목/구분
        value_cols, label_cols = [], []
        for c in cols:
            if c == ycolname:
                continue
            if is_texty(body[c]) >= 0.5:
                label_cols.append(c)
            else:
                value_cols.append(c)
        if not value_cols:
            return None, "세로형 추정했으나 숫자 지표열 없음", ["확인 필요"]
        ent = label_cols[0] if label_cols else None
        records = []
        for _, row in body.iterrows():
            name = str(row[ent]).strip() if ent else "전체"
            for vc in value_cols:
                val = parse_number(row[vc])
                if pd.isna(val):
                    continue
                records.append({"entity_code": "", "entity": name,
                                "category": vc, "year": int(row["__year__"]), "value": val})
        tidy = pd.DataFrame.from_records(records)
        notes.append(f"세로형: '{ycolname}'열=연도, 지표열 {len(value_cols)}개")
        if tidy.empty:
            return None, "; ".join(notes), ["데이터에서 숫자를 찾지 못함"]
        return tidy, "; ".join(notes), warns

    return None, "시간축(연도) 자동 감지 실패", ["표 구조 확인 필요 — 첫 행을 연/분기 헤더로 정리해 재실행"]


# ── tidy → 분석 지표 ────────────────────────────────────────────────────────
def cagr(first_val, last_val, n_years):
    if n_years <= 0 or first_val is None or last_val is None:
        return np.nan
    if pd.isna(first_val) or pd.isna(last_val):
        return np.nan
    # 시작값이 0 이하이거나 종료값이 음수면 CAGR 무의미(거듭제곱이 복소수가 됨)
    if first_val <= 0 or last_val < 0:
        return np.nan
    return float((last_val / first_val) ** (1.0 / n_years) - 1.0)


def build_growth(tidy: pd.DataFrame) -> pd.DataFrame:
    """(entity, category) 시계열별 첫·끝값, CAGR, 최근 YoY."""
    rows = []
    for (code, ent, cat), g in tidy.groupby(["entity_code", "entity", "category"], dropna=False):
        g = g.sort_values("year")
        years = g["year"].tolist()
        vals = g["value"].tolist()
        if len(years) < 2:
            continue
        first_y, last_y = years[0], years[-1]
        first_v, last_v = vals[0], vals[-1]
        span = last_y - first_y
        last_yoy = np.nan
        if len(vals) >= 2 and vals[-2] not in (0, None) and not pd.isna(vals[-2]):
            last_yoy = vals[-1] / vals[-2] - 1.0
        rows.append({
            "항목코드": code, "항목": ent, "구분": cat,
            "시작연도": first_y, "종료연도": last_y,
            "시작값": first_v, "종료값": last_v,
            "기간": span, "CAGR": cagr(first_v, last_v, span),
            "최근YoY": last_yoy,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["구분", "종료값"], ascending=[True, False])
    return df


def build_timeseries_wide(tidy: pd.DataFrame) -> pd.DataFrame:
    """연도를 열로 펼친 피벗(항목·구분 × 연도). 사람이 보기 좋게."""
    pv = tidy.pivot_table(index=["entity_code", "entity", "category"],
                          columns="year", values="value", aggfunc="sum")
    pv = pv.reset_index().rename(columns={"entity_code": "항목코드", "entity": "항목", "category": "구분"})
    return pv


HEADLINE_NAMES = ["매출", "매출액", "영업수익", "수익", "총매출", "revenue", "sales"]


def build_comparison(tidy: pd.DataFrame, growth: pd.DataFrame) -> pd.DataFrame:
    """대표항목(예: 매출)을 골라 '구분(사업부)별' 최신연도 구성을 비교한다.

    P&L 표는 행(항목)이 매출·비용·영업이익처럼 이질적이라 단순 합산 구성비가 무의미하다.
    대신 한 대표항목을 잡고, 그 항목이 구분(사업부)별로 어떻게 나뉘는지 본다.
    """
    last_year = int(tidy["year"].max())
    cats = tidy["category"].unique().tolist()
    total_like = [c for c in cats if str(c).strip().lower() in {"total", "합계", "계", "총계"}]
    total_cat = total_like[0] if total_like else None

    # 대표항목 선택: 이름 우선 매칭, 없으면 (Total 구분의) 최신연도 절대값 최대 항목
    ents = tidy["entity"].unique().tolist()
    headline = next((e for e in ents for k in HEADLINE_NAMES
                     if k.lower() in str(e).lower()), None)
    if headline is None:
        base_cat = total_cat or (cats[0] if cats else "")
        latest = tidy[(tidy["year"] == last_year) & (tidy["category"] == base_cat)]
        if not latest.empty:
            headline = latest.loc[latest["value"].abs().idxmax(), "entity"]
        elif ents:
            headline = ents[0]

    sub = tidy[(tidy["year"] == last_year) & (tidy["entity"] == headline)].copy()
    sub = sub.groupby("category", as_index=False)["value"].sum()
    # 분모: Total 구분 값(있으면), 없으면 Total 외 구분 합
    if total_cat is not None and (sub["category"] == total_cat).any():
        denom = float(sub.loc[sub["category"] == total_cat, "value"].iloc[0])
        sub = sub[sub["category"] != total_cat]  # 구성 비교에서는 Total 제외
    else:
        denom = float(sub["value"].sum())
    sub["구성비"] = sub["value"] / denom if denom else np.nan
    sub = sub.sort_values("value", ascending=False)
    sub = sub.rename(columns={"category": "구분", "value": f"{last_year}_{headline}"})
    sub.attrs["meta"] = (last_year, headline, total_cat)
    return sub


# ── digest.md 생성 ──────────────────────────────────────────────────────────
def fmt_pct(x):
    return "n/a" if pd.isna(x) else f"{x*100:+.1f}%"


def fmt_num(x):
    return "n/a" if pd.isna(x) else f"{x:,.0f}"


def write_digest(per_file_results):
    lines = ["# Excel 분석 digest (analyze_excel.py 자동 생성)\n",
             "> 이 파일은 Claude `/week1` 이 읽어 현황요약·시사점에 인용합니다. 수치 출처는 `analysis.xlsx`.\n"]
    for fr in per_file_results:
        lines.append(f"\n## 파일: {fr['file']}  (시트: {fr['sheet']})")
        lines.append(f"- 감지: {fr['note']}")
        if fr.get("warns"):
            for w in fr["warns"]:
                lines.append(f"- ⚠️ {w}")
        if fr.get("tidy") is None:
            lines.append("- ❌ 시계열 분석 불가 — 표 구조 확인 필요. 아래 미리보기 참고.")
            lines.append("```")
            lines.append(fr["preview"])
            lines.append("```")
            continue
        tidy = fr["tidy"]
        yrs = sorted(tidy["year"].unique())
        lines.append(f"- 기간: {yrs[0]}~{yrs[-1]} ({len(yrs)}개 구간), 항목 {tidy['entity'].nunique()}개, 구분 {tidy['category'].nunique()}개")
        g = fr["growth"]
        if not g.empty:
            # 최신값 상위 시계열의 CAGR/최근YoY
            top = g.sort_values("종료값", ascending=False).head(8)
            lines.append("\n**핵심 시계열 (최신값 상위)**\n")
            lines.append("| 항목 | 구분 | 시작→종료 | 종료값 | CAGR | 최근YoY |")
            lines.append("|---|---|---|---:|---:|---:|")
            for _, r in top.iterrows():
                lines.append(f"| {r['항목']} | {r['구분']} | {r['시작연도']}→{r['종료연도']} | "
                             f"{fmt_num(r['종료값'])} | {fmt_pct(r['CAGR'])} | {fmt_pct(r['최근YoY'])} |")
            # 급증·급감 Top (3년 이상 구간만 — 신규 항목의 1년 폭증을 배제)
            gv = g[g["CAGR"].notna() & (g["기간"] >= 3)].sort_values("CAGR", ascending=False)
            if not gv.empty:
                lines.append("\n**CAGR 상위 3 / 하위 3 (3년+ 구간, 성장·역성장 항목)**\n")
                for label, part in [("상위", gv.head(3)), ("하위", gv.tail(3))]:
                    for _, r in part.iterrows():
                        lines.append(f"- [{label}] {r['항목']} / {r['구분']}: CAGR {fmt_pct(r['CAGR'])} "
                                     f"({r['시작연도']}→{r['종료연도']}, {fmt_num(r['시작값'])}→{fmt_num(r['종료값'])})")
        comp = fr.get("comparison")
        if comp is not None and not comp.empty:
            ly, headline, _tot = comp.attrs.get("meta", ("", "", None))
            valcol = [c for c in comp.columns if c not in ("구분", "구성비")][0]
            lines.append(f"\n**{ly}년 '{headline}'의 구분(사업부)별 구성**\n")
            lines.append("| 구분 | 값 | 구성비 |")
            lines.append("|---|---:|---:|")
            for _, r in comp.head(12).iterrows():
                lines.append(f"| {r['구분']} | {fmt_num(r[valcol])} | {fmt_pct(r['구성비'])} |")
    lines.append("\n---\n_주의: CAGR는 시작값이 0 이하이면 n/a. 수치는 반드시 analysis.xlsx 로 교차확인하세요._\n")
    DIGEST_OUT.write_text("\n".join(lines), encoding="utf-8")


# ── 메인 ────────────────────────────────────────────────────────────────────
def load_tables():
    """00_inbox/excel 의 파일들을 (파일명, 시트명, raw_df) 로 yield."""
    if not EXCEL_DIR.exists():
        return
    for path in sorted(EXCEL_DIR.iterdir()):
        if not path.is_file():
            continue
        ext = path.suffix.lower()
        try:
            if ext in {".xlsx", ".xls"}:
                xl = pd.ExcelFile(path)
                for s in xl.sheet_names:
                    yield path.name, s, pd.read_excel(path, sheet_name=s, header=None)
            elif ext == ".csv":
                yield path.name, "(csv)", pd.read_csv(path, header=None)
        except Exception as e:
            print(f"  XX 읽기 실패 {path.name}: {e}", file=sys.stderr)


def write_scan(results):
    """--scan: 계산 없이 '무슨 표인지'만 파악해 _scan.md 로. (분석 전 파악용)"""
    lines = ["# Excel 구조 스캔 (analyze_excel.py --scan)\n",
             "> 계산은 하지 않고 표의 모양만 파악했습니다. `/triage` 가 이걸 읽어 분석계획을 세웁니다.\n"]
    for fr in results:
        lines.append(f"\n## {fr['file']}  (시트: {fr['sheet']})")
        lines.append(f"- 감지: {fr['note']}")
        for w in fr.get("warns", []):
            lines.append(f"- ⚠️ {w}")
        if fr.get("tidy") is None:
            lines.append("- ❌ 시계열 자동 인식 실패 — 표 구조 확인 필요. 미리보기:")
            lines.append("```\n" + fr["preview"] + "\n```")
            continue
        t = fr["tidy"]
        yrs = sorted(t["year"].unique())
        ents = t["entity"].unique().tolist()
        cats = [c for c in t["category"].unique().tolist() if str(c).strip()]
        lines.append(f"- 기간: {yrs[0]}~{yrs[-1]} ({len(yrs)}개 구간)")
        lines.append(f"- 항목(행) {len(ents)}개 — 예: {', '.join(map(str, ents[:8]))}")
        if cats:
            lines.append(f"- 구분(열그룹) {len(cats)}개 — {', '.join(map(str, cats))}")
    (OUT / "_scan.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"  → {OUT / '_scan.md'}")


def main() -> int:
    scan_only = "--scan" in sys.argv
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    sheets = list(load_tables())
    if not sheets:
        print("[안내] 00_inbox/excel 에 .xlsx/.csv 를 넣어주세요.")
        if scan_only:
            (OUT / "_scan.md").write_text("# Excel 구조 스캔\n\n(00_inbox/excel 비어 있음)\n", encoding="utf-8")
        else:
            write_digest([])  # 빈 결과라도 digest 는 남긴다
        return 0

    # ── --scan: 구조만 파악(계산 X) ──────────────────────────────────────
    if scan_only:
        for fname, sheet, raw in sheets:
            tidy, note, warns = to_tidy(raw)
            print(f"  [{fname} / {sheet}] {note}")
            results.append({"file": fname, "sheet": sheet, "note": note,
                            "warns": warns, "preview": raw.iloc[:12, :10].to_string(),
                            "tidy": tidy})
        write_scan(results)
        print(f"\n스캔 완료: {len(results)}개 시트. (분석은 /week1 또는 --scan 없이 실행)")
        return 0

    with pd.ExcelWriter(XLSX_OUT, engine="openpyxl") as writer:
        for fname, sheet, raw in sheets:
            preview = raw.iloc[:12, :10].to_string()
            tidy, note, warns = to_tidy(raw)
            print(f"  [{fname} / {sheet}] {note}")
            res = {"file": fname, "sheet": sheet, "note": note, "warns": warns, "preview": preview}
            safe = re.sub(r"[^0-9A-Za-z가-힣]", "_", sheet)[:20] or "sheet"
            # 원본 미리보기는 항상 기록
            raw.iloc[:20, :].to_excel(writer, sheet_name=f"미리보기_{safe}"[:31], index=False, header=False)
            if tidy is None:
                res["tidy"] = None
                results.append(res)
                continue
            growth = build_growth(tidy)
            ts = build_timeseries_wide(tidy)
            comp = build_comparison(tidy, growth)
            res.update({"tidy": tidy, "growth": growth, "comparison": comp})
            results.append(res)
            tidy.to_excel(writer, sheet_name=f"tidy_{safe}"[:31], index=False)
            ts.to_excel(writer, sheet_name=f"시계열_{safe}"[:31], index=False)
            growth.to_excel(writer, sheet_name=f"증감_{safe}"[:31], index=False)
            comp.to_excel(writer, sheet_name=f"비교_{safe}"[:31], index=False)

    write_digest(results)
    n_ok = sum(1 for r in results if r.get("tidy") is not None)
    print(f"\n완료: {len(results)}개 시트 중 {n_ok}개 시계열 분석 성공.")
    print(f"  → {XLSX_OUT}")
    print(f"  → {DIGEST_OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
