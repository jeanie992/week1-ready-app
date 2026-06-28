# -*- coding: utf-8 -*-
"""
llm_notes.py : 재무제표 행별 NOTE 인사이트를 생성

인사이트 생성 우선순위 :
  1) ANTHROPIC_API_KEY 있으면 Claude API (배포 환경 등)
  2) 로컬에 Claude Code CLI(claude) 있으면 `claude -p` 로 생성 → Max 구독 사용(API 키 불필요)
  3) 둘 다 없으면 규칙 기반(YoY 급변 표시)
문체 : 끝 온점 없음, 하이픈 대신 콜론
"""
from __future__ import annotations
import os, json, shutil, subprocess

MODEL = "claude-sonnet-4-6"   # API 경로용 노트 생성 모델


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


def _build_prompt(statements, years, company, extra_context):
    table = _to_eok_table(statements, years)
    return (
        f"너는 BCG 재무 분석가다. 아래 {company}의 3개 재무제표(억원)와 (있으면) 추가 맥락을 본다.\n"
        f"{extra_context}\n\n{table}\n\n"
        "분석 방식 : 단순 YoY/CAGR 나열 금지 : **Top-down 원인 추적**으로 인사이트를 만든다\n"
        " 1) 큰 변화(대항목)를 먼저 짚는다 : 매출·영업이익·순이익·자산의 증감과 이상치\n"
        " 2) 어디서 비롯됐나(중항목) : 어떤 하위계정/부문이 그 변화를 견인·악화시켰나\n"
        " 3) 왜 그런가(세부항목) : 구체 원인 : 다른 계정에 단서가 있나(예 영업이익 하락 ← 지급임차료 급증)\n"
        " 4) 거시환경 영향인지 회사 특정이슈인지 구분(단정 못하면 가설로)\n"
        "각 핵심 계정의 NOTE 에는 위 추적의 **가장 관련된 한 고리**(원인·연결)를 담는다 : 단순 증감률 반복 금지\n"
        "규칙 : 한국어 : 끝 온점 없음 : 하이픈 대신 콜론( : ) : 각 45자 이내 : 출처 없는 단정 금지(불확실은 '추정')\n"
        '의미있는 계정 20~30개만. JSON 배열만 출력(다른 말 금지) : [{"statement":"IS|BS|CF","account":"계정명","note":"..."}]'
    )


def _parse_arr(txt):
    s = txt[txt.find("["): txt.rfind("]") + 1]
    arr = json.loads(s)
    return {(d["statement"], d["account"]): d["note"] for d in arr if d.get("note")}


def _via_api(prompt, key):
    import anthropic
    client = anthropic.Anthropic(api_key=key)
    msg = client.messages.create(model=MODEL, max_tokens=4000,
                                 messages=[{"role": "user", "content": prompt}])
    txt = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    return _parse_arr(txt)


def _via_cli(prompt):
    """로컬 Claude Code(Max 구독)로 생성 : claude CLI 가 있을 때만"""
    claude = shutil.which("claude")
    if not claude:
        return None
    p = subprocess.run([claude, "-p", prompt, "--output-format", "text"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=240)
    if p.returncode != 0 or not (p.stdout or "").strip():
        return None
    return _parse_arr(p.stdout)


def generate(statements, years, company="회사", extra_context="", use_llm=True):
    """반환 : {(statement, account): note}"""
    if not use_llm:
        return _rule_based(statements, years)
    prompt = _build_prompt(statements, years, company, extra_context)
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if key:
        try:
            r = _via_api(prompt, key)
            print(f"[llm_notes] Claude API 로 생성 : {len(r)}건")
            return r
        except Exception as e:
            print(f"[llm_notes] API 실패 → 다음 경로 : {e}")
    try:
        r = _via_cli(prompt)
        if r:
            print(f"[llm_notes] Claude Code(Max 구독)로 생성 : {len(r)}건")
            return r
    except Exception as e:
        print(f"[llm_notes] CLI 실패 → 규칙 기반 : {e}")
    return _rule_based(statements, years)
