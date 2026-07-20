"""Gemini 기반 검수(QA) 모듈 — GEMINI_API_KEY 환경변수 필요.

Claude가 생성한 카피를 클로드가 아닌 다른 AI(Gemini)로 교차 검수한다:
왜곡/날조, 문장 단위 저작권(표절), 사실확인.
"""
import os, json

REVIEW_MODEL = "gemini-3.5-flash"

_RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "issues": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "field": {"type": "string"},
                    "category": {
                        "type": "string",
                        "enum": ["왜곡/날조", "저작권/표절", "사실확인"],
                    },
                    "quote": {"type": "string"},
                    "explanation": {"type": "string"},
                    "severity": {"type": "string", "enum": ["높음", "중간", "낮음"]},
                },
                "required": ["field", "category", "quote", "explanation", "severity"],
            },
        },
        "summary": {"type": "string"},
    },
    "required": ["issues", "summary"],
}

_SYSTEM_INSTRUCTION = (
    "당신은 여행 상품 기획안 카피를 검수하는 팩트체커입니다. "
    "[원본 자료]와 [AI가 생성한 카피(JSON)]를 비교해서 다음 세 가지 문제만 찾아주세요.\n"
    "1. 왜곡/날조: 원본에 없는 사실을 지어내거나, 원본 내용을 과장/왜곡한 부분\n"
    "2. 저작권/표절: 원본 자료나 잘 알려진 외부 자료의 문장을 문장 구조·어순까지 거의 "
    "그대로 베낀 부분 (단순 표현 유사는 제외)\n"
    "3. 사실확인: 지명, 연도, 수치, 고유명사 등 확인이 필요한데 근거가 불명확하거나 "
    "틀린 것으로 보이는 부분\n"
    "각 문제는 어느 JSON 키(field)에서 발견됐는지, 문제 문장을 quote에 그대로 인용해서 "
    "지적하세요. 원본에 없는 정보라도 잘 알려진 일반 상식 수준이면 지적하지 마세요. "
    "문제가 없으면 issues를 빈 배열로 반환하세요."
)


def review_content(content_json, source_material_text, dry_run=False):
    if dry_run or not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        return {
            "_dry_run": True,
            "_note": "GEMINI_API_KEY가 설정되지 않아 검수를 건너뜁니다.",
            "issues": [],
            "summary": "",
        }

    from google import genai
    from google.genai import types

    client = genai.Client()

    user_content = (
        f"[원본 자료]\n{source_material_text}\n\n"
        f"[AI가 생성한 카피 (JSON)]\n{json.dumps(content_json, ensure_ascii=False, indent=2)}"
    )

    resp = client.models.generate_content(
        model=REVIEW_MODEL,
        contents=user_content,
        config=types.GenerateContentConfig(
            system_instruction=_SYSTEM_INSTRUCTION,
            response_mime_type="application/json",
            response_json_schema=_RESPONSE_SCHEMA,
            temperature=0,
        ),
    )

    try:
        return json.loads(resp.text)
    except (json.JSONDecodeError, TypeError) as e:
        raise RuntimeError(
            f"Gemini 검수 응답을 JSON으로 파싱하는 데 실패했습니다: {e}\n"
            f"응답 원문: {resp.text[:500] if resp.text else '(empty)'}"
        ) from e
