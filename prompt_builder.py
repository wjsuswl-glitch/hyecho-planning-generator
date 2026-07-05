"""few-shot 예시 선택 + 시스템 프롬프트 조립 모듈"""
import json
import unicodedata

import os
FEWSHOT_PATH = os.path.join(os.path.dirname(__file__), "data", "fewshot_examples.json")


def _nfc(s):
    """macOS 등에서 NFD(자모 분리형)로 저장된 한글 문자열을
    NFC(완성형)로 정규화. JSON 파일의 값과 코드 내 문자열 리터럴의
    유니코드 정규화 형식이 다르면 '정현지' == '정현지' 비교가 실패한다."""
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s

STYLE_RULES = {
    "박소설": "문학적·서정적 문체, 형용사와 비유를 적극 활용. 표지에 Design Direction(톤앤매너/색상/키워드)을 명시.",
    "신윤정": "정보 밀도 높은 구조적 문체, 넘버링 카드(특별함 N가지)와 비교표를 활용.",
    "정현지": "함축적·담백한 문체. 표지는 2줄 대구 형태 태그라인. 3~5장으로 압축.",
}

LAYOUT_HINT = {"박소설": "separate", "신윤정": "combined", "정현지": "separate"}
BANNER_MAP_INCLUDE = {"박소설": True, "신윤정": True, "정현지": False}

# template_map.json의 field_map / repeatable_groups와 1:1로 맞춘 스키마.
# 여기가 template_map.json과 어긋나면 assembler.py에서 "NO DATA"만 계속 쌓인다 —
# 필드를 추가/삭제할 땐 반드시 template_map.json도 같이 바꿀 것.
SCHEMA_HINTS = {
    "정현지": """{
  "cover": {"tagline": str, "product_name": str, "region_tag": str, "subtitle": str, "intro_copy": str},
  "brand_tagline": str,
  "brand_points": [str, str],
  "watermark_label": str,
  "why_hyecho": {
    "title": str, "subtitle1": str, "subtitle2": str,
    "badge": str, "section_title": str, "theme_line": str
  },
  "destinations": [ {"title": str, "description": str} ],
  "season": {"title": str, "content": str}
}
※ watermark_label: 표지 우하단에 작게 들어가는 영문 1~2단어 (여정/지역명)
※ destinations: 위 예시는 배열 안에 원소 1개만 보여준 것입니다. 실제로는
  {"title": str, "description": str} 형태의 원소를 실제 입력에 있는 개수만큼 반복하세요.""",
    "신윤정": """{
  "cover": {"tagline": str, "subtitle": str, "intro_copy": str},
  "why_hyecho": {"title": str, "points": [str, str, str, str]}
}""",
}

DESTINATIONS_RULE = (
    "[destinations 배열 규칙]\n"
    "destinations는 실제 입력에 실제로 등장하는 경유지/명소 개수만큼만 생성하세요.\n"
    "예를 들어 하이라이트가 3곳이면 정확히 3개만 만드세요. 템플릿에 슬롯이 몇 개 있든 "
    "상관없이, 있지도 않은 경유지를 지어내서 슬롯을 채우면 안 됩니다. "
    "부족한 슬롯은 조립 단계에서 자동으로 삭제됩니다."
)

def load_fewshot_examples(writer_style, category, k=3):
    with open(FEWSHOT_PATH, encoding="utf-8") as f:
        all_examples = json.load(f)

    writer_style = _nfc(writer_style)
    category = _nfc(category)

    same = [e for e in all_examples
            if _nfc(e["writer_style"]) == writer_style and _nfc(e["category"]) == category]
    other_cat = [e for e in all_examples
                 if _nfc(e["writer_style"]) == writer_style and _nfc(e["category"]) != category]
    picked = (same[:2] + other_cat[:1])[:k]
    return picked

def build_system_prompt(writer_style, category, parsed_sections, format_info):
    examples = load_fewshot_examples(writer_style, category)
    draft_copy = format_info.get("draft_copy")

    if draft_copy:
        copy_instruction = (
            f"[표지 카피 생성 규칙 — 다듬기 모드]\n"
            f"사업부 자료에 이미 카피 초안이 있습니다: \"{draft_copy}\"\n"
            f"이 문구를 거의 그대로 유지하되, {writer_style}의 문체에 맞게 어미와 리듬만 다듬으세요. "
            f"의미나 핵심 단어는 바꾸지 마세요."
        )
    else:
        copy_instruction = (
            f"[표지 카피 생성 규칙 — 창작 모드]\n"
            f"사업부 자료에 카피 초안이 없습니다. '컨셉'과 '담당자 기획 의도'에 나온 핵심 개념을 "
            f"재료로 삼아 {writer_style}의 문체로 2줄 태그라인을 새로 창작하세요."
        )

    prompt = f"""역할: 당신은 혜초여행사 콘텐츠팀의 {writer_style} 기획자입니다.

[스타일 규칙]
{STYLE_RULES[writer_style]}

[레이아웃 규칙]
why_hyecho와 season 섹션은 {"같은 슬라이드에 합쳐서" if LAYOUT_HINT[writer_style]=="combined" else "별도 슬라이드로 나눠서"} 구성하세요.
배너/지도 슬라이드는 {"포함" if BANNER_MAP_INCLUDE[writer_style] else "생략"}하세요.

[고정 문구 뱅크]
"혜초와 함께하면", "No 쇼핑! No 옵션!", "국내 유일", "전 일정 인솔자 동행"
→ 문맥에 자연스럽게 녹여 쓰되 남발하지 않음

{copy_instruction}

[출력 형식]
반드시 JSON으로만 응답하세요. 다른 텍스트를 포함하지 마세요.
스키마 (이 구조를 정확히 따르세요. 필드를 빼거나 이름을 바꾸지 마세요):
{SCHEMA_HINTS.get(writer_style, "(스키마 미정의 — 담당자에게 문의)")}

{DESTINATIONS_RULE if "destinations" in SCHEMA_HINTS.get(writer_style, "") else ""}

[Few-shot 예시 {len(examples)}개 — 문체·구조 참고 전용]
아래 예시들은 전혀 다른 여행 상품(지명, 코스, 하이라이트 등)에 대한 과거 결과물입니다.
문장 톤·문단 구성 방식·섹션 나누는 방식만 참고하세요.
예시에 등장하는 지명, 상품명, 문구, 숫자, 이미지 캡션은 절대 그대로 재사용하지 마세요.
아래 [실제 입력]에 없는 내용(예: 안데스, 페루, 마추픽추 등 예시 속 고유명사)이
출력에 등장하면 안 됩니다. 반드시 [실제 입력]에 있는 사실만으로 콘텐츠를 생성하세요.
{json.dumps(examples, ensure_ascii=False, indent=2)}

[실제 입력 — 사업부 원본자료 파싱 결과 (이 내용만을 근거로 생성)]
{json.dumps(parsed_sections, ensure_ascii=False, indent=2)}
"""
    return prompt
