# -*- coding: utf-8 -*-
"""
llm_notes.py : 재무제표 행별 NOTE 인사이트를 생성

- ANTHROPIC_API_KEY 가 있으면 Claude API 로 행별 한 줄 인사이트 생성(회사 무관 범용)
- 없으면 규칙 기반(YoY 급변 표시)으로 폴백
문체 : 끝 온점 없음, 하이픈 대신 콜론
"""
from __future__ import annotations
import os, json

MODEL = "claude-sonnet-4-6"   # 노트 생성용(빠르고 충분) : 더 깊게는 claude-opus-4-8


def _rule_based(statements, years):
    """폴백 : 마지막 해 YoY가 ±30% 넘으면 표시"""
    notes = {}
    y2, y1 = years[-1], years[-2]
    for st in ("IS", "BS", "CF"):
        for acc, yv in statements[st].items():
            a, b = yv.get(y1), yv.get(y2)
            if a and a > 0 and b is not None:
                g = b / a - 1
                if abs(g) >= 0.3:
                    notes[(st, acc)] = f"{y2} 전년비 {'급증' if g>0 else '급감'} {g*100:+.0f}%"
    return notes


def _to_eok_table(statements, years):
    lines = []
    for st, title in [("IS", "손익계산서"), ("BS", "재무상태표"), ("CF", "현금흐름표")]:
        lines.append(f"## {title} (억원)")
        for acc, yv in statements[st].items():
            vals = " ".join(f"{y}:{(yv[y]/1e8):.1f}" if yv.get(y) is not None else f"{y}:-" for y in years)
            lines.append(f"{acc} | {vals}")
    return "\n".join(lines)


def generate(statements, years, company="회사", extra_context="", use_llm=True):
    """반환 : {(statement, account): note}"""
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not (use_llm and key):
        return _rule_based(statements, years)
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        table = _to_eok_table(statements, years)
        prompt = (
            f"너는 재무 분석가다. 아래는 {company}의 3개 재무제표(단위 억원)다.\n"
            f"{extra_context}\n\n{table}\n\n"
            "각 '핵심 계정'에 대해 추세·비중·이상치를 근거로 한 줄 인사이트를 쓴다.\n"
            "규칙 : 한국어 : 문장 끝 온점 없음 : 하이픈 대신 콜론( : ) : 각 노트 40자 이내 : "
            "모든 계정이 아니라 의미있는 계정만(20~30개) : 출처 없는 추정 금지\n"
            'JSON 배열만 출력 : [{"statement":"IS|BS|CF","account":"계정명","note":"..."}]'
        )
        msg = client.messages.create(
            model=MODEL, max_tokens=4000,
            messages=[{"role": "user", "content": prompt}],
        )
        txt = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        s = txt[txt.find("["): txt.rfind("]") + 1]
        arr = json.loads(s)
        return {(d["statement"], d["account"]): d["note"] for d in arr if d.get("note")}
    except Exception as e:
        print(f"[llm_notes] API 실패 → 규칙 기반 폴백 : {e}")
        return _rule_based(statements, years)
