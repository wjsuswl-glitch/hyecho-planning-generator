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
    "박소설": "문학적·서정적 문체, 형용사와 비유를 적극 활용.",
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
  "watermark_label": str,
  "background_story": {"kicker": str, "title": str, "content": str},
  "why_reasons": [ {"title": str, "content": str} ],
  "destinations_heading": str,
  "destinations": [ {"title": str, "description": str, "region_tag": str} ],
  "route_compare": {
    "title": str,
    "routes": [ {"name": str, "course": str, "scenery": str, "appeal": str, "summary": str} ]
  },
  "brand_tagline": str,
  "experience_points": [ {"title": str, "description": str} ],
  "highlights_heading": str,
  "highlights": [ {"title": str, "description": str} ],
  "season": {"title": str, "content": str, "stat_line": str},
  "season_table": [ {"month": str, "high": str, "low": str} ],
  "altitude_profile": [ {"name": str, "altitude": str} ],
  "safety_note": {"question": str, "answer": str}
}
※ cover.tagline: 표지 맨 위에 작게 들어가는 짧은 감성 문구 (2줄 이내, 꾸미는 말)
※ cover.product_name: 실제 상품명입니다. 사업부 자료의 상품 제목에 있는 핵심 내용(시리즈명,
  테마 태그 등)은 전부 반영하되, "[차마고도 3편] [감탄절로] 사천에서 티벳까지 천장공로 12일"
  처럼 대괄호 태그를 그대로 나열하지 말고 자연스러운 하나의 문장/구로 풀어서 녹여내세요
  (예: "차마고도 3편, 감탄절로 사천에서 티벳까지 천장공로 12일"). 대괄호는 절대 쓰지 마세요.
  이 필드가 표지의 메인 타이틀(가장 큰 글씨)과 배너 슬라이드에 그대로 노출되므로, 절대
  비워두거나 다른 감성 문구로 대체하지 마세요.
※ cover.subtitle: 이미지 아래에 들어가는 보조 설명 한 줄입니다. 반드시 여행객이 읽을
  마케팅 카피여야 합니다 (예: "사천에서 티벳까지, 2,140km 고원의 서사"). 이 문서 자체를
  설명하는 메타 문구("상품 소개 기획안", "디자인팀 전달용", "기획안입니다" 등)는 절대
  쓰지 마세요 — 그런 문구는 실제 고객이 보는 화면에 그대로 노출되는 심각한 오류입니다.
※ watermark_label: 표지 우상단에 작게 들어가는 영문 1~2단어 (여정/지역명)
※ background_story: "OOO란?" 같은 여행지/노선의 배경·역사·유래를 설명하는 섹션입니다.
  사업부 자료에 이런 배경 설명이 없으면, 잘 알려진 일반 상식 수준에서만 채우고
  구체적 사실(연도, 수치 등)을 지어내지 마세요.
※ why_reasons: "왜 이 지역/노선인가"를 설명하는 이유 목록입니다. 실제 입력 내용에 근거해
  2~4개 작성하세요.
※ destinations: 위 예시는 배열 안에 원소 1개만 보여준 것입니다. 실제로는
  {"title": str, "description": str, "region_tag": str} 형태의 원소를 실제 입력에 있는 개수만큼
  반복하세요. region_tag는 그 목적지가 속한 지역/국가/성(省) 이름이며, 없으면 빈 문자열로 두세요.
  사업부 자료에 없는 지명을 지어내진 마세요. description은 2~3문장 이내로 간결하게 쓰세요
  (너무 길면 레이아웃이 깨집니다).
※ route_compare: 사업부 자료에 대안 코스/노선 비교 내용이 있을 때만 채우세요. 없으면
  routes를 빈 배열로 두세요 (있지도 않은 대안 코스를 지어내지 마세요).
※ destinations_heading: 목적지 소개 섹션의 제목입니다 (예: "OOO 하이라이트"). brand_tagline과는
  다른 문구로 작성하세요 (같은 말 반복 금지).
※ experience_points: 노쇼핑/편안한 이동/독보적 일정 같은 "혜초만의 차별점" 카드 2~3개.
※ highlights_heading, highlights: destinations와는 다른, 더 큰 테마 단위의 "여정 하이라이트"
  입니다 (예: "국내 유일 육로 횡단", "매일 변화하는 풍경", "문화의 공존" 등 3~4개). 특정
  지명 하나가 아니라 여정 전체를 관통하는 주제로 작성하세요.
  highlights_heading은 highlights가 하나라도 있으면 절대 빈 값으로 두지 마세요 — 아래
  카드들을 한 문장으로 아우르는 제목입니다 (예: "한눈에 보는 여정 하이라이트!").
※ season.stat_line: 계절 섹션 상단의 짧은 강조 배너 문구 (예: "최적기: O월~O월")
※ season_table: 월별 기온 등 계절 통계가 사업부 자료에 있을 때만 채우세요. 없으면 빈
  배열로 두세요.
※ altitude_profile, safety_note: 고산 트레킹은 "고산증", 도보순례는 "체력/보험",
  일반 하이킹은 "난이도" 등 카테고리에 맞는 안전/난이도 안내가 필요한 상품에만 채우세요.
  해당 없는 상품(저지대 여행 등)이면 둘 다 빈 값/생략하세요. safety_note는 혜초 홈페이지의
  표준 안내 톤(과장 없이 사실 위주)을 따르세요.
※ 사업부 자료에 정보가 부족한 필드(예: experience_points 문구)는 빈 값으로 두지
  말고, 사업부 자료의 사실에 기반해 정현지 문체로 자연스럽게 채워서 완성하세요. 단, destinations나
  route_compare, season_table처럼 사실 데이터가 필요한 항목에 없는 내용을 새로 지어내는 것은
  금지입니다 — 채우기는 "표현"에 대한 것이지 "사실 날조"가 아닙니다.""",
    "신윤정": "__SAME_AS_정현지__",
    "박소설": "__SAME_AS_정현지__",
}
SCHEMA_HINTS["박소설"] = SCHEMA_HINTS["정현지"]  # 박소설도 동일한 동적 빌더(builder.py) 스키마 사용
SCHEMA_HINTS["신윤정"] = SCHEMA_HINTS["정현지"]  # 신윤정도 동일한 동적 빌더(builder.py) 스키마 사용

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

def build_system_prompt(writer_style, category, parsed_sections, format_info, has_images=False):
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

[중요 — 모든 필드 공통 규칙]
아래 스키마의 모든 값은 실제 고객이 보게 될 화면에 그대로 노출됩니다. 어떤 필드에도
이 작업/문서 자체에 대한 메타 설명("상품 소개 기획안입니다", "디자인팀 전달용",
"기획안", "AI가 생성한", "다음은 ~입니다" 등)을 쓰지 마세요. 모든 텍스트는 실제
여행 상품을 소개하는 마케팅 카피여야 합니다.

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
{"""
[첨부 이미지]
이 메시지에는 사업부에서 제공한 이미지 파일(사진, 지도, 옛 자료 스크린샷 등)이 함께
첨부되어 있습니다. 이미지 안에 보이는 지명, 설명, 표, 일정 같은 실제 정보도 위 텍스트
자료와 동등한 '사업부 원본자료'로 취급해 반영하세요. 다만 이미지가 단순 풍경/분위기
참고용인 경우 억지로 사실 정보를 추출하려 하지 말고, 명확히 읽을 수 있는 텍스트/데이터만
사용하세요.""" if has_images else ""}
"""
    return prompt
