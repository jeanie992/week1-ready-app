# Week 1 현황파악 도구

회사 자료를 빠르게 파악하는 웹 도구 : 자료를 올리면 문서 추출 · Excel 시계열 분석 · DART 공시 수집 · 재무 마스터 엑셀을 만든다

## 기능
- 문서(PDF/Word/PPT/HWP) → 텍스트 추출
- 임의 스키마 Excel 시계열·비교 분석 (YoY/CAGR)
- DART 3개년 감사/사업보고서 수집 + 재무제표 자동 파싱
- 재무 마스터 엑셀 생성 (손익·재무상태·현금흐름, NOTE 인사이트 포함)

## 로컬 실행
```bash
python -m venv .venv
. .venv/bin/activate            # Windows : .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## 키 (선택)
- DART API 키 : 앱 입력란 또는 환경변수 `DART_API_KEY` / Streamlit Secrets
- Anthropic API 키 : NOTE 인사이트 자동 생성용 : 없으면 규칙 기반으로 동작

## 구성
- `app.py` : Streamlit UI
- `scripts/convert.py` : 문서 → 텍스트
- `scripts/analyze_excel.py` : Excel 분석
- `scripts/dart_fetch.py` : DART 보고서 수집
- `scripts/dart_parse.py` : DART → 재무제표 구조화
- `scripts/build_master_generic.py` : 재무 마스터 엑셀
- `scripts/llm_notes.py` : NOTE 인사이트
- `scripts/excel_style.py` : 엑셀 양식 모듈
