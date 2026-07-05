# 혜초 기획안 자동생성 (프로토타입)

사업부 원본자료(docx)를 기획자 스타일에 맞는 기획안 PPTX로 자동 생성합니다.

## 실행 방법
```
pip install -r requirements.txt
export ANTHROPIC_API_KEY="본인 키"   # 없으면 프롬프트 미리보기까지만 동작
python3 -m streamlit run app.py
```

## 구성
- `parser.py` — 사업부 워드 자료 파싱 (유형 A~E 자동 판별, draft_copy 감지)
- `prompt_builder.py` — few-shot 예시 선택 + 스타일별 프롬프트 조립
- `generator.py` — Claude API 호출
- `assembler.py` — 생성된 JSON을 PPTX 도형에 주입
- `template_map.json` — 기획자별 템플릿 경로 및 필드 매핑 (현재 정현지·신윤정만 등록)
- `app.py` — Streamlit 웹 UI

## 현재 상태
- 정현지(안데스)·신윤정(방글라데시) 실제 상품으로 엔드투엔드 검증 완료
- 박소설 스타일은 사업부 원본자료 미확보로 아직 미등록
