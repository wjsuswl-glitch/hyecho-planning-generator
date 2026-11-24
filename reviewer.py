"""검수(QA) 모듈 — provider="gemini"(GEMINI_API_KEY) 또는 "claude"(ANTHROPIC_API_KEY).

카피를 쓴 AI와 다른 AI가 교차 검수한다(app.py가 생성의 반대편 AI를 고른다):
왜곡/날조, 문장 단위 저작권(표절), 사실확인, 비문/맞춤법.
"""
import os, json

REVIEW_MODEL = "gemini-3.5-flash"

# experience_points는 "이 상품만의 사실"이 아니라 혜초의 모든 상품에 정책적으로 항상
# 들어가는 공통 서비스 문구(No 쇼핑/No 옵션, 전 일정 인솔자 동행, 항공/호텔 등급, 전용
# 이동수단)다. prompt_builder.py의 [고정 문구 뱅크]가 이 필드에만 이 문구를 쓰도록
# 지시하고 있고, 사업부 자료(개별 상품 소개)에는 원래부터 안 나오는 게 정상이다 —
# 그런데 Gemini 검수가 "원본에 없다"는 이유만으로 이걸 왜곡/날조로 잘못 지적하는
# 문제가 있었다(카라코람 재테스트에서 확인됨). 그래서 이 필드는 왜곡/날조·사실확인
# 검수 대상에서 제외한다고 명시한다.
_EXEMPT_FIELDS_NOTE = (
    "단, experience_points 필드는 예외입니다 — 이건 특정 상품에 대한 사실 주장이 아니라 "
    "혜초여행의 모든 상품에 공통으로 적용되는 표준 서비스 정책 문구(No 쇼핑/No 옵션, 전 "
    "일정 인솔자 동행, 항공/호텔 등급, 전용 이동수단 등)입니다. 사업부 자료(개별 상품 "
    "소개)에 이런 문구가 없는 것은 정상이므로, experience_points 필드는 왜곡/날조와 "
    "사실확인 검수 대상에서 제외하세요. (sections 배열을 쓰는 새 구성에서는 \"type\":\"service_cards\"인 "
    "섹션이 이에 해당합니다.)\n"
    "또한 composition_note 필드와 각 섹션의 why 필드는 기획자에게 구성 이유를 설명하는 내부 "
    "메모이며 고객 화면에 나오지 않으므로, 검수 대상에서 제외하세요."
)

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
                        "enum": ["왜곡/날조", "저작권/표절", "사실확인", "비문/맞춤법", "구성/배치"],
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
    "당신은 여행 상품 기획안 카피를 검수하는 팩트체커 겸 교정자입니다. "
    "[원본 자료]와 [AI가 생성한 카피(JSON)]를 비교해서 다음 다섯 가지 문제만 찾아주세요.\n"
    "1. 왜곡/날조: 여행지·코스·일정 등 이 상품 고유의 사실에 대해, 원본에 없는 내용을 "
    "지어내거나 원본 내용을 과장/왜곡한 부분(지역 설명, 코스 정보, 역사·지리적 사실 등)\n"
    "2. 저작권/표절: 원본 자료나 잘 알려진 외부 자료의 문장을 문장 구조·어순까지 거의 "
    "그대로 베낀 부분 (단순 표현 유사는 제외)\n"
    "3. 사실확인: 지명, 연도, 수치, 고유명사 등 확인이 필요한데 근거가 불명확하거나 "
    "틀린 것으로 보이는 부분\n"
    "4. 비문/맞춤법: 맞춤법·띄어쓰기·조사 오용, 주어-서술어 불일치, 문장이 안 끝나거나 "
    "어색하게 꼬인 비문 등 국어 표현 자체의 오류(사실 여부와는 무관)\n"
    "5. 구성/배치: 섹션 제목과 항목의 성격이 어긋난 곳(예: '쉼·숙박' 섹션에 짐 운반 같은 서비스 항목이 "
    "들어간 경우), 같은 서비스·혜택이 3곳 이상에서 반복되거나 Q&A가 다른 블록의 내용을 그대로 되묻는 "
    "중복. 사실 오류가 아니므로 severity는 '낮음' 또는 '중간'으로 하세요.\n"
    "특히 다음 두 가지는 꼼꼼히 확인하세요(2026-10-27 추가 — '호수 3곳 + 암봉 3곳'을 "
    "'여섯 개의 호수'라고 쓰고도 검수를 통과한 사례가 있었음): (a) '여섯 개의 호수', "
    "'세 곳' 같은 개수·수량 표현이 원본 자료의 항목을 실제로 센 값과 일치하는지, 서로 "
    "다른 종류를 합쳐 한 종류의 개수처럼 쓰지 않았는지 — 어긋나면 왜곡/날조(높음)로 "
    "지적. (b) 최적기·시즌 문구와 본문 설명이 서로 모순되지 않는지, 원본이 말한 월 "
    "범위와 맞는지.\n"
    f"{_EXEMPT_FIELDS_NOTE}\n"
    "각 문제는 어느 JSON 키(field)에서 발견됐는지, 문제 문장을 quote에 그대로 인용해서 "
    "지적하세요. 원본에 없는 정보라도 잘 알려진 일반 상식 수준이면 지적하지 마세요. "
    "문제가 없으면 issues를 빈 배열로 반환하세요."
)


def _review_with_gemini(content_json, source_material_text, dry_run=False):
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


_CLAUDE_REVIEW_MODEL = "claude-sonnet-5"


def _loads_loose(text):
    text = (text or "").strip()
    cands = [text]
    if text.startswith("```"):
        parts = text.split("```")
        inner = parts[1] if len(parts) > 1 else text
        cands.append(inner[4:].strip() if inner.startswith("json") else inner.strip())
    a, b = text.find("{"), text.rfind("}")
    if a != -1 and b > a:
        cands.append(text[a:b + 1])
    for c in cands:
        try:
            return json.loads(c)
        except (json.JSONDecodeError, TypeError):
            continue
    return None


def _review_with_claude(content_json, source_material_text, dry_run=False):
    """Gemini가 쓴 카피를 Claude가 검수한다(같은 AI의 자기 검수를 피하기 위한 교차검수)."""
    if dry_run or not os.environ.get("ANTHROPIC_API_KEY"):
        return {
            "_dry_run": True,
            "_note": "ANTHROPIC_API_KEY가 설정되지 않아 검수를 건너뜁니다.",
            "issues": [],
            "summary": "",
        }

    import anthropic

    client = anthropic.Anthropic()
    system = (
        _SYSTEM_INSTRUCTION
        + "\n\n응답은 설명 없이 JSON 객체 하나만 출력하세요. 형식: "
        + '{"issues":[{"field":str,"category":"왜곡/날조|저작권/표절|사실확인|비문/맞춤법|구성/배치",'
        + '"quote":str,"explanation":str,"severity":"높음|중간|낮음"}],"summary":str}'
    )
    user_content = (
        f"[원본 자료]\n{source_material_text}\n\n"
        f"[AI가 생성한 카피 (JSON)]\n{json.dumps(content_json, ensure_ascii=False, indent=2)}"
    )
    resp = client.messages.create(
        model=_CLAUDE_REVIEW_MODEL,
        max_tokens=8000,
        thinking={"type": "disabled"},
        system=system,
        messages=[{"role": "user", "content": user_content}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    result = _loads_loose(text)
    if not isinstance(result, dict) or not isinstance(result.get("issues"), list):
        raise RuntimeError(
            "Claude 검수 응답을 JSON으로 파싱하는 데 실패했습니다.\n"
            f"응답 원문: {text[:500] if text else '(empty)'}"
        )
    result.setdefault("summary", "")
    return result


def review_content(content_json, source_material_text, dry_run=False, provider="gemini"):
    """provider: 검수를 맡을 AI("gemini" 또는 "claude"). 반환 형식은 같다."""
    if str(provider).lower() == "claude":
        return _review_with_claude(content_json, source_material_text, dry_run=dry_run)
    return _review_with_gemini(content_json, source_material_text, dry_run=dry_run)
