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
    # 실제 제작한 상품소개/기획전 기획안 4건(인도 건축기행, 귀주성 트레킹, 제주도
    # 춘하추동, 설국열차 이벤트)의 어투를 분석해 도출 — 체험(직접·온몸으로) 중심의
    # 생동감 있는 문체가 핵심. "\"...\"" 인용구로 후킹 문구를 열고, 계절/테마명은
    # 한자를 병기하며(예: [봄 春]), ①②③ 원문자 번호로 소주제를 구분한다. "~해
    # 보세요!" 같은 강한 청유형 CTA로 문단을 자주 마무리하고, "세계 최대", "국내
    # 유일", "아시아 최대" 같은 최상급 비교 표현을 적극 활용한다. 대표이사·임원
    # 답사처럼 신뢰를 줄 수 있는 구체적 근거(실명, 직접 답사 등)가 사업부 자료에
    # 있으면 놓치지 말고 살려서 쓴다.
    "최정인": (
        "체험 중심의 생동감 있는 문체 — '직접', '온몸으로', '걸으며' 등 몸으로 겪는 "
        "표현을 즐겨 쓴다. \"...\" 인용구로 감성적인 후킹 문구를 열고, ①②③ 같은 "
        "원문자 번호로 소주제를 나눈다. 계절/테마명에는 한자를 병기할 수 있다(예: "
        "\"[봄 春] 유채꽃과 생명의 숨결\"). 문단은 종종 '~해 보세요!' 같은 강한 "
        "청유형 CTA로 마무리한다. '세계 최대', '국내 유일', '아시아 최대'처럼 "
        "확인 가능한 최상급 비교 표현을 적극 활용해 상품의 규모/희소성을 강조한다."
    ),
}

LAYOUT_HINT = {"박소설": "separate", "신윤정": "combined", "정현지": "separate", "최정인": "separate"}
BANNER_MAP_INCLUDE = {"박소설": True, "신윤정": True, "정현지": False, "최정인": True}

# template_map.json의 field_map / repeatable_groups와 1:1로 맞춘 스키마.
# 여기가 template_map.json과 어긋나면 assembler.py에서 "NO DATA"만 계속 쌓인다 —
# 필드를 추가/삭제할 땐 반드시 template_map.json도 같이 바꿀 것.
SCHEMA_HINTS = {
    "정현지": """{
  "cover": {"tagline": str, "product_name": str, "region_tag": str, "subtitle": str, "intro_copy": str},
  "watermark_label": str,
  "product_variant_type": str,
  "background_story": {"title": str, "content": str},
  "why_reasons": [ {"title": str, "content": str} ],
  "destinations_heading": str,
  "destinations": [ {"title": str, "description": str, "region_tag": str} ],
  "tour_spots_heading": str,
  "tour_spots": [ {"title": str, "description": str, "region_tag": str} ],
  "route_compare": {
    "title": str,
    "routes": [ {"name": str, "course": str, "scenery": str, "appeal": str, "summary": str} ]
  },
  "transport_spec": {"title": str, "specs": [ {"label": str, "value": str} ]},
  "brand_tagline": str,
  "experience_points": [ {"title": str, "description": str} ],
  "guide_profile": [ {"name": str, "title": str, "bio": str} ],
  "season": {"title": str, "content": str, "stat_line": str},
  "season_table": [ {"month": str, "high": str, "low": str} ],
  "meal_info": {"question": str, "answer": str},
  "altitude_profile": [ {"name": str, "altitude": str, "distance": str, "duration": str, "highlight": str} ],
  "safety_note": {"question": str, "answer": str},
  "banner_copy": {"kicker": str, "title": str}
}
※ cover.tagline: 표지 맨 위에 작게 들어가는 짧은 감성 문구 (2줄 이내, 꾸미는 말)
※ cover.region_tag: 여러 국가·지역을 넘나드는 상품일 때만 그 범위를 짚어주는 상위
  지역명(예: "남미 5개국 완전일주"라면 "남아메리카")을 채우세요. 단일 국가/지역만
  방문하는 상품이면 반드시 빈 문자열로 두세요 — product_name에 이미 그 나라/지역명이
  들어가는 경우가 많아(예: "멕시코 문명기행 13일") 표지에 같은 지명이 또 한 번
  나오면 불필요한 중복입니다.
※ cover.product_name: 실제 상품명입니다. 여행지·노선·기간처럼 고객이 실제로 궁금해하는
  핵심 정보만 자연스러운 하나의 문장/구로 담으세요 (예: "사천에서 티벳까지 천장공로 12일").
  대괄호는 절대 쓰지 마세요. 사업부 자료 제목의 대괄호 태그는 다음 두 종류를 구분해서
  처리하세요:
  1) "이지트레킹", "프리미엄", "특가", "얼리버드" 같은 내부 상품군/난이도/캠페인 분류
     태그는 고객에게 의미 없는 내부 라벨이므로 product_name에서 완전히 제외하세요
     (product_variant_type 필드로 이미 별도 반영됩니다).
  2) "차마고도 3편", "감탄절로"처럼 시리즈명·캠페인명 등 실제 마케팅 정보인 태그는
     product_name 앞에 붙이지 말고 cover.tagline(표지 맨 위 감성 문구)에 자연스럽게
     녹여내거나, tagline에도 안 맞으면 과감히 생략하세요.
  이 필드가 표지의 메인 타이틀(가장 큰 글씨)과 배너 슬라이드에 그대로 노출되므로, 절대
  비워두거나 다른 감성 문구로 대체하지 마세요.
※ cover.subtitle: 이미지 아래에 들어가는 보조 설명 한 줄입니다. 반드시 여행객이 읽을
  마케팅 카피여야 합니다 (예: "사천에서 티벳까지, 2,140km 고원의 서사"). 이 문서 자체를
  설명하는 메타 문구("상품 소개 기획안", "디자인팀 전달용", "기획안입니다" 등)는 절대
  쓰지 마세요 — 그런 문구는 실제 고객이 보는 화면에 그대로 노출되는 심각한 오류입니다.
※ watermark_label: 표지 우상단에 작게 들어가는 영문 1~2단어 (여정/지역명)
※ product_variant_type: 상품의 이동수단/난이도 성격 태그입니다. "육상"(일반 도보·차량
  이동 여정), "크루즈"(선박이 핵심 이동수단), "고소·극한등반"(6,000m급 이상이거나 신청
  자격 제한이 있는 등반) 중 사업부 자료 내용에 맞는 하나를 고르세요. 애매하면 "육상"으로
  두세요.
※ background_story: 여행지/노선의 배경·역사·유래를 설명하는 섹션입니다. title(임팩트 있는
  헤드라인, 예: "세계의 지붕을 잇는, 실크로드의 마지막 길")과 content(설명 문단)만
  작성하세요. "OOO란?" 같은 형식적인 소제목(kicker)은 모든 상품 기획안에 기계적으로
  반복되는 상투적 표현이라 절대 만들지 마세요 — title로 바로 임팩트 있게 시작하세요.
  사업부 자료에 이런 배경 설명이 없으면 웹 검색 도구로 사실을 확인한 뒤 채우세요
  (연도, 유래, 지리적 사실 등 구체적 정보일수록 검색으로 확인하고, 지어내지 마세요).
  검색 결과가 서로 다르거나 확인이 안 되는 세부사항은 그 부분만 빼고 확실한 내용만
  쓰세요. 검색으로도 못 찾으면 잘 알려진 일반 상식 수준에서만 채우세요.
※ why_reasons: "이 상품만의 차별점"을 설명하는 이유 목록입니다 (2~4개). 반드시 장소·코스·
  풍광·일정·계절 한정성처럼 이 상품 고유의 요소만 다루세요 (예: "1년에 단 두 번만 열리는
  계절", "타사가 가지 않는 숨은 트레킹 코스", "국내 유일의 노선"). "No 쇼핑/No 옵션",
  "전 일정 인솔자 동행", 항공/호텔/이동수단처럼 상품과 무관하게 항상 제공되는 혜초 공통
  서비스 요소는 여기 쓰지 마세요 — 그건 experience_points의 역할입니다. 타사 대비 명확한
  차별점(유일 노선 등)이 사업부 자료에 있다면 "~와는 다릅니다", "비교해보면 답은 혜초"
  같은 직접적인 비교 어조도 자연스럽게 쓸 수 있습니다 — 다만 실제로 비교할 근거가 있을
  때만 이 톤을 쓰세요.
※ destinations: 위 예시는 배열 안에 원소 1개만 보여준 것입니다. 실제로는
  {"title": str, "description": str, "region_tag": str} 형태의 원소를 실제 입력에 있는 개수만큼
  반복하세요. 사업부 자료에 없는 지명을 지어내진 마세요. description은 2~3문장 이내로
  간결하게 쓰세요(너무 길면 레이아웃이 깨집니다). 사업부 자료가 "트레킹"과 "관광"처럼 코스
  성격을 구분해 놓았다면(트레킹 상품 중 일부는 트레킹 코스와 시내/유적지 관광 코스가 함께
  있습니다), destinations에는 트레킹·액티비티 코스만 담고, 관광지·박물관·도시 방문처럼 걷거나
  체험하는 트레킹이 아닌 코스는 아래 tour_spots로 분리하세요. 자료에 이런 구분이 없는 일반
  상품(트레킹이 아니거나, 방문지 전체가 같은 성격인 상품)이면 tour_spots는 비워두고
  destinations 하나에 전부 담으세요.
※ destinations[].region_tag — 방문지의 "상위 카테고리"(그룹 라벨): 방문지가 여러 개일 때,
  성격이 비슷한 것끼리(예: 같은 문명/테마) 또는 같은 지역/도시에 있는 것끼리 묶어서 그
  묶음 전체를 대표하는 카테고리 이름입니다. 방문지가 나열식으로만 쭉 이어지면 읽기
  피곤하다는 피드백과, 그룹핑을 넣었는데도 상위 카테고리가 안 보인다는 피드백이 모두
  있었습니다 — region_tag는 destinations_heading(섹션 전체 제목, 예: "방문지 하이라이트")
  보다 한 단계 아래, 개별 방문지 제목보다 한 단계 위에 있는 "카테고리 제목"이라는 걸
  잊지 마세요. 실제로 PPTX에서는 destinations_heading과 똑같은 색 배경 바(banner)
  스타일로 렌더링되니, region_tag 값도 방문지 이름 하나가 아니라 "카테고리처럼 읽히는
  짧은 구"로 쓰세요.
  - intro_copy나 background_story에서 이미 여러 개의 테마/레이어를 명시적으로 짚어준
    상품(예: "사포테카 → 아스테카 → 식민도시 → 마야, 문명 4개 레이어" 같은 문구)이라면,
    그 레이어/테마 이름을 그대로 region_tag로 재사용해 destinations를 나누세요 —
    앞에서 이미 정의한 카테고리를 방문지 섹션에서 다시 안 쓰면, 앞뒤 내용이 따로
    노는 것처럼 보이고 상위 카테고리가 사라져 보입니다.
  - 사업부 자료에 이미 그룹 구조가 있다면(예: "[ 사포테카 문명 | 와하카 ]", "[ 식민
    기획도시 | 과달라하라·과나후아토·케레타로 ]"처럼 테마와 지역을 함께 묶어놓은 괄호
    표기) 그 구조를 그대로 살려서 region_tag에 짧은 그룹명으로 담으세요(예: "사포테카
    문명 · 와하카"). 테마 구분이 없고 지역/도시만 구분된다면 지역명만 써도 됩니다
    (예: "와하카").
  - 방문지가 3개 이상이고 사업부 자료나 여정 순서상 자연스럽게 2개 이상의 묶음으로
    나뉜다면(지리적으로 인접, 같은 문명/시대, 같은 액티비티 유형 등) region_tag를
    비워두지 말고 적극적으로 채우세요 — region_tag를 전부 비우는 선택은 방문지들이
    정말로 서로 무관하고(예: 전부 한 도시 안의 개별 명소라 더 나눌 필요가 없음) 사업부
    자료에도 그룹 지을 근거가 전혀 없을 때만 하세요.
  - 반드시 지켜야 할 규칙: 같은 그룹에 속하는 방문지는 destinations 배열에서 반드시
    서로 붙어 있어야 합니다(그룹 A의 방문지를 전부 나열한 다음에 그룹 B로 넘어가는
    식) — 그룹이 배열 중간에 흩어져 있으면 화면에 같은 그룹명이 여러 번 끊겨서 나옵니다.
※ tour_spots_heading, tour_spots: destinations와 형태는 동일하지만
  ("title": str, "description": str, "region_tag": str}), 트레킹 코스가 아닌 관광지(도시,
  유적지, 박물관 등)만 담는 별도 섹션입니다. 트레킹 코스와 관광 코스가 모두 있는 상품에서만
  채우고, tour_spots_heading도 그때만 채우세요(예: "함께 즐기는 관광 코스"). 해당 없으면
  tour_spots_heading은 빈 문자열, tour_spots는 빈 배열로 두세요. region_tag는 destinations와
  동일한 그룹 라벨 규칙을 따르세요(위 destinations[].region_tag 설명 참고).
※ route_compare: 사업부 자료에 대안 코스/노선 비교 내용이 있을 때만 채우세요. 없으면
  routes를 빈 배열로 두세요 (있지도 않은 대안 코스를 지어내지 마세요).
※ transport_spec: 열차·크루즈처럼 이동수단 자체가 상품의 핵심 매력인 경우에만 채우세요
  (예: "The Ghan 럭셔리 열차", "모션 알바트로스 크루즈"). specs는 객실타입/부대시설/톤수/
  안전등급처럼 사업부 자료에 실제로 있는 스펙만 담으세요. 일반적인 항공/버스 이동에는
  채우지 말고 title을 빈 문자열, specs를 빈 배열로 두세요.
※ destinations_heading: 목적지 소개 섹션의 제목입니다 (예: "OOO 하이라이트"). brand_tagline과는
  다른 문구로 작성하세요 (같은 말 반복 금지).
※ experience_points: "혜초와 함께라면 편안한 이유" — 상품과 무관하게 항상 제공되는 혜초
  공통 서비스 차별점 카드 2~4개입니다. No 쇼핑/No 옵션, 전 일정 인솔자 동행, 항공/호텔
  등급, 전용 이동수단(전용차량·전세버스 등) 중 사업부 자료에서 확인되는 항목 위주로
  구성하세요. 이 상품만의 장소·코스·계절 이야기는 절대 여기 쓰지 마세요 — 그건
  why_reasons의 역할입니다. why_reasons와 문장이 겹치면 안 됩니다.
※ guide_profile: 인솔자·가이드·담당자의 실제 이력(경력, 등정/순례 횟수, 자격 등)이
  사업부 자료에 있을 때만 채우세요. 있지도 않은 인물이나 이력을 지어내지 마세요. 없으면
  빈 배열로 두세요.
※ season: 사계절 내내 상시 운영되는 상품이 아닌 이상 반드시 채우세요 (season.content를
  비워두지 마세요). 최적기, 피해야 할 시기, 특정 계절에만 볼 수 있는 것(개화·단풍·적설 등)
  처럼 이 상품에 실제로 해당하는 계절 정보를 구체적으로 담으세요. 사계절 상시 운영
  상품(계절 영향이 없는 도심 관광 등)일 때만 비워도 됩니다.
※ season.stat_line: 계절 섹션 상단의 짧은 강조 배너 문구 (예: "최적기: O월~O월")
※ season_table: 월별 기온 등 계절 통계가 사업부 자료에 있을 때만 채우세요. 없으면 빈
  배열로 두세요.
※ meal_info: "여행/트레킹 중 식사는 어떻게 하나요?" 같은 실용 정보 Q&A입니다. safety_note와
  같은 형식(question/answer)이며, 사업부 자료에 식사 관련 정보(산장식/현지식/포함 여부 등)가
  있을 때만 채우세요. 없으면 question, answer 모두 빈 문자열로 두세요.
※ altitude_profile, safety_note: 고산 트레킹은 "고산증", 도보순례는 "체력/보험",
  일반 하이킹은 "난이도" 등 카테고리에 맞는 안전/난이도 안내가 필요한 상품에만 채우세요.
  해당 없는 상품(저지대 여행 등)이면 둘 다 빈 값/생략하세요. safety_note는 혜초 홈페이지의
  표준 안내 톤(과장 없이 사실 위주)을 따르되, 신청 자격 제한이나 환불 불가 조건처럼 사업부
  자료에 강한 경고성 유의사항이 있다면 완곡하게 순화하지 말고 사실대로 명확히 전달하세요.
  altitude_profile의 distance(구간 거리, 예: "15km")와 duration(소요시간, 예: "약 6시간")은
  사업부 자료에 있을 때만 채우고, 없으면 빈 문자열로 두세요. altitude_profile의 highlight는
  그 구간에서 볼 수 있는 풍경이나 경험(예: "빙하와 협곡이 빚어낸 장관", "만년설 봉우리
  파노라마")을 8~14자 이내로 간결하게 채우세요 — 이름과 거리만 나열하지 말고 각 구간이
  실제로 왜 매력적인지 알 수 있게 하세요. 사업부 자료의 설명에 근거해 작성하고, 자료에
  없으면 지형·코스 특징에 기반해 일반적인 수준으로 자연스럽게 표현해도 됩니다(사실
  날조는 금지, 표현은 허용).
※ banner_copy: 배너에 들어갈 문구로, cover.tagline이나 cover.product_name을 그대로
  재사용하지 마세요. 훨씬 짧고 강렬하게 압축한 별도 카피입니다. kicker는 감성적인 후킹
  문구 1~2줄(예: "봄으로 물든 카라코람을 걷다"), title은 굵고 임팩트 있는 핵심 키워드
  1~2줄(예: "1년에 단 두 번, 카라코람 하이라이트")입니다. 실제 상품명·표지 카피와 정확히
  일치할 필요는 없습니다 — 상품의 가장 눈에 띄는 포인트만 뽑아 짧게 후킹하는 것이
  목적입니다.
※ 사업부 자료에 정보가 부족한 필드(예: experience_points 문구)는 빈 값으로 두지
  말고, 사업부 자료의 사실에 기반해 정현지 문체로 자연스럽게 채워서 완성하세요. 단, destinations나
  route_compare, season_table, transport_spec.specs, guide_profile, meal_info,
  altitude_profile의 distance/duration처럼 사실 데이터가 필요한 항목에 없는 내용을 새로
  지어내는 것은 금지입니다 — 채우기는 "표현"에 대한 것이지 "사실 날조"가 아닙니다.""",
    "신윤정": "__SAME_AS_정현지__",
    "박소설": "__SAME_AS_정현지__",
}
SCHEMA_HINTS["박소설"] = SCHEMA_HINTS["정현지"]  # 박소설도 동일한 동적 빌더(builder.py) 스키마 사용
SCHEMA_HINTS["신윤정"] = SCHEMA_HINTS["정현지"]  # 신윤정도 동일한 동적 빌더(builder.py) 스키마 사용
SCHEMA_HINTS["최정인"] = SCHEMA_HINTS["정현지"]  # 최정인도 동일한 동적 빌더(builder.py) 스키마 사용

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
    # load_fewshot_examples는 내부적으로 _nfc()로 정규화해서 비교하지만, 그건 그
    # 함수 안에서만 쓰이는 로컬 변수라 여기 build_system_prompt의 writer_style
    # 자체는 정규화되지 않은 채로 남아있었다 — 그 결과 STYLE_RULES[writer_style]처럼
    # 바로 아래에서 하는 딕셔너리 조회는 그대로 원본 인코딩을 써서, Streamlit
    # 화면의 selectbox에서 넘어온 문자열이 자모 분리형(NFD)이면(예: 맥OS 환경에서
    # 저장된 template_map.json의 키가 NFD인 경우) 여기 파일에 NFC로 적힌 키와
    # 안 맞아 KeyError가 났다("최정인" 추가 후 실제로 발생 — 화면엔 똑같이
    # "최정인"으로 보여도 바이트 단위로는 다른 문자열이라 딕셔너리 조회가 실패함).
    # 함수 맨 앞에서 한 번 정규화해두면 이 함수 안의 모든 딕셔너리 조회
    # (STYLE_RULES/LAYOUT_HINT/BANNER_MAP_INCLUDE/SCHEMA_HINTS)와 f-string
    # 삽입까지 전부 안전해진다.
    writer_style = _nfc(writer_style)
    category = _nfc(category)

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
why_reasons와 season 섹션은 {"같은 슬라이드에 합쳐서" if LAYOUT_HINT[writer_style]=="combined" else "별도 슬라이드로 나눠서"} 구성하세요.

[웹 검색 도구]
web_search 도구를 사용할 수 있습니다. 사업부 자료에 없는 배경지식·역사·지리적 사실이
필요할 때(특히 background_story) 검색으로 확인한 뒤 반영하세요. 검색 없이 추측으로
연도·수치·고유명사를 지어내지 마세요 — 확인 안 되면 그 세부사항은 빼세요.
검색은 꼭 필요한 최소한(가능하면 1~2회)으로만 사용하세요. "검색을 진행하겠습니다",
"검색 결과가 비어 있네요", "확인했습니다", "이제 ~하겠습니다" 같은 진행 상황 설명은
단 한 글자도 출력하지 마세요 — 이런 문장이 섞이면 최종 응답 전체가 무효 처리됩니다.
검색이 끝나면 다른 말 없이 곧바로 JSON 응답을 시작하세요. 최종 응답은 오직 JSON
객체 하나만이어야 하며, 그 앞이나 뒤에 어떤 텍스트도 있으면 안 됩니다.
배너/지도 슬라이드는 {"포함" if BANNER_MAP_INCLUDE[writer_style] else "생략"}하세요.

[고정 문구 뱅크 — experience_points 전용]
"혜초와 함께하면", "No 쇼핑! No 옵션!", "전 일정 인솔자 동행"
→ experience_points에서만 문맥에 자연스럽게 녹여 쓰되 남발하지 않음. why_reasons에는
이 문구들을 쓰지 마세요 (역할이 다릅니다 — 위 experience_points/why_reasons 설명 참고).

{copy_instruction}

[버전 분기 판단 — 대부분의 상품엔 해당 없음]
사업부 자료 안에 "봄 버전과 가을 버전으로 2개로 해주세요"처럼, 계절이나 시기에 따라
결과물 자체를 여러 개로 나눠 만들어달라는 명시적 요청이 있는지 확인하세요. 이런 요청이
없으면(대부분의 경우) 아래 [출력 형식]대로 스키마 객체 하나만 그대로 출력하세요 — 이
섹션은 무시하세요.
명시적 요청이 있을 때만, 최종 응답을 아래처럼 감싸서 요청된 버전 개수만큼 각각 완전한
스키마 객체를 만들어 담으세요:
{{
  "multi_version": true,
  "versions": [
    {{"version_label": "봄", ... 아래 스키마의 모든 필드 ...}},
    {{"version_label": "가을", ... 아래 스키마의 모든 필드 ...}}
  ]
}}
각 버전은 완전히 독립된 콘텐츠여야 합니다 — 사업부 자료가 지시한 대로 그 시기에 맞는
표현(예: 봄=살구꽃, 가을=황금빛 미루나무)을 표지, 배경 설명, 계절 섹션, 목적지 설명 등
관련된 모든 필드에 일관되게 반영하세요. 같은 문장을 복붙하지 말고 각 버전마다 그 계절/
시기에 맞게 다시 쓰세요. version_label은 사업부 자료에 쓰인 표현을 그대로 쓰세요
(예: "봄", "가을").

[출력 형식]
반드시 JSON으로만 응답하세요. 다른 텍스트를 포함하지 마세요. 응답은 반드시 중괄호
{{ }}로 시작하는 JSON 객체 하나여야 합니다 — 대괄호 [ ]로 감싼 배열이 아닙니다.
아래 [Few-shot 예시]는 여러 개라서 배열( [ ] )로 보여지는 것일 뿐이며, 당신의
최종 답변은 그 예시들처럼 배열로 감싸지 말고 스키마와 똑같이 객체 하나만
출력하세요 (예: [{{"cover": ...}}]가 아니라 {{"cover": ...}}) — 위 [버전 분기 판단]에
해당하는 경우에만 예외적으로 "versions" 배열을 그 객체 "안에" 담습니다.

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
