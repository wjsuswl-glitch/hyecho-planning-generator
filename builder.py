"""동적 PPTX 빌더 (v2) — 기존 완성 기획안을 열어서 도형을 덮어쓰는 대신,
매번 필요한 만큼만 슬라이드를 새로 만든다.

기존 assembler.py 방식(템플릿 기획안을 열어 shape_id별로 텍스트 주입)은
- 옛 기획안의 이미지/깨진 도형이 그대로 남는 문제
- 목적지 개수가 템플릿 슬롯 수와 안 맞으면 삭제 로직이 계속 필요한 문제
가 있어서, 이 모듈은 옛 기획안(예: 안데스_템플릿.pptx)을 "디자인 참고용"으로만
쓰고, 실제 출력은 python-pptx로 슬라이드를 새로 그린다.
"""
import json
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- 스타일 상수 (정현지 스타일 기본값) ----
SLIDE_W = Emu(6858000)   # 원본 기획안과 동일한 슬라이드 크기 (약 7.5 x 10.83 inch, 모바일 세로형)
SLIDE_H = Emu(9906000)
FONT_NAME = "맑은 고딕"
ACCENT_COLOR = RGBColor(0x1B, 0x4D, 0x6B)   # 진한 티얼/네이비 (섹션 바, 강조)
TEXT_COLOR = RGBColor(0x22, 0x22, 0x22)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED_COLOR = RGBColor(0x66, 0x66, 0x66)
MARGIN = Inches(0.4)
CONTENT_W = SLIDE_W - MARGIN * 2


def _tf_setup(tf, text, size, color, bold=False, align=PP_ALIGN.LEFT, font=FONT_NAME):
    tf.word_wrap = True
    lines = str(text).split("\n")
    tf.text = lines[0]
    p0 = tf.paragraphs[0]
    p0.alignment = align
    for run in p0.runs or [p0.add_run()]:
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.text = line
        p.alignment = align
        r = p.add_run() if not p.runs else p.runs[0]
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font


def estimate_text_height(text, size_pt, width_emu, line_spacing=1.22, bold=False):
    """글자 수 기반으로 텍스트가 실제로 차지할 높이를 대략 추정한다.
    고정 간격 대신 이걸 써야 설명 길이에 따라 다음 요소와 안 겹친다.
    실제 렌더링 폭 추정은 부정확할 수 있어 여유 마진을 둔다(단, 슬라이드 장수를
    압축하기 위해 과도한 여유는 줄였다 — line_spacing 1.35→1.22, 버퍼 0.15→0.08in)."""
    if not text:
        return Inches(0.1)
    width_in = Emu(width_emu).inches
    # 한글 기준 글자 폭 대략치 (볼드면 좀 더 넓게 잡음), 안전 마진 포함
    char_w_in = (size_pt / 72) * (1.05 if bold else 0.95)
    chars_per_line = max(1, int(width_in / char_w_in))
    total_lines = 0
    for line in str(text).split("\n"):
        total_lines += max(1, -(-len(line) // chars_per_line))  # ceil
    line_h_in = (size_pt / 72) * line_spacing
    return Inches(total_lines * line_h_in + 0.08)  # 여유 마진


def add_text(slide, left, top, width, height, text, size=14, color=TEXT_COLOR,
             bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.vertical_anchor = anchor
    _tf_setup(tf, text, size, color, bold, align)
    return box


def add_section_bar(slide, top, text, height=Inches(0.45), size=15):
    """섹션 제목이 들어가는 진한 색 배경 바 (예: '천장공로 하이라이트')"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, top, CONTENT_W, height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT_COLOR
    bar.line.fill.background()
    tf = bar.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_setup(tf, text, size, WHITE, bold=True, align=PP_ALIGN.CENTER)
    return bar


def add_image_placeholder(slide, left, top, width, height, label="이미지"):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.background()
    box.line.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    tf = box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _tf_setup(tf, label, 11, MUTED_COLOR, align=PP_ALIGN.CENTER)
    return box


def add_small_image_placeholder(slide, top, width, height, label="이미지"):
    """전체 폭을 다 차지하는 큰 이미지 자리 대신, 가로 폭도 훨씬 좁힌 작은 썸네일
    크기 자리표시. CONTENT_W 안에서 가운데 정렬해서 배치한다. 실제 사진은 디자이너가
    작업하므로, 여기서는 '사진이 들어갈 위치'만 작게 표시하고 나머지 공간은 텍스트가
    위로 당겨져 채우게 한다."""
    left = MARGIN + (CONTENT_W - width) / 2
    return add_image_placeholder(slide, left, top, width, height, label)


def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # 완전 빈 레이아웃


class SlideFlow:
    """여러 섹션 함수가 슬라이드 하나를 공유해서 이어 쓸 수 있게 하는 커서.

    PPTX는 프레젠테이션 전체가 슬라이드 높이를 하나만 가질 수 있어서(슬라이드별
    높이 지정 불가), 식사 안내(1.6in)처럼 짧은 섹션도 그동안은 10.83in짜리 슬라이드를
    통째로 차지해 빈 공간이 컸다. ensure()로 "남은 공간에 들어가면 이어 붙이고,
    안 들어가면 새 슬라이드"를 섹션마다 판단해 슬라이드 장수와 여백을 줄인다."""

    def __init__(self, prs):
        self.prs = prs
        self.slide = None
        self.y = Inches(0.3)
        self.bottom_limit = SLIDE_H - Inches(0.2)

    def new_slide(self):
        self.slide = _blank_slide(self.prs)
        self.y = Inches(0.3)
        return self.slide

    def ensure(self, height, gap_before=Inches(0.35)):
        """다음 섹션(height)을 놓을 자리를 확보한다. 이어 붙일 수 있으면 그 y좌표를,
        없으면 새 슬라이드를 시작하고 그 y좌표를 반환한다."""
        if self.slide is None:
            self.new_slide()
            return self.y
        candidate_y = self.y + gap_before
        if candidate_y + height > self.bottom_limit:
            self.new_slide()
            return self.y
        self.y = candidate_y
        return self.y


# ---------------------------------------------------------------------------
# 슬라이드 빌더 함수 (정현지 스타일) — 모두 SlideFlow를 받아 가능하면 같은
# 슬라이드에 이어 그리고, 공간이 없을 때만 새 슬라이드를 시작한다.
# ---------------------------------------------------------------------------

def build_cover_slide(flow, cover, watermark_label=""):
    slide = flow.new_slide()
    y = flow.y
    add_text(slide, MARGIN, y, CONTENT_W, Inches(0.5), cover.get("tagline", ""),
              size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
    y += Inches(0.55)
    add_text(slide, MARGIN, y, CONTENT_W, Inches(0.9), cover.get("product_name", ""),
              size=26, bold=True, align=PP_ALIGN.CENTER)
    y += Inches(0.95)
    if cover.get("region_tag"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), cover["region_tag"],
                  size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += Inches(0.45)
    y += Inches(0.1)
    add_small_image_placeholder(slide, y, Inches(2.3), Inches(0.9), "메인 이미지")
    y += Inches(1.0)
    if cover.get("subtitle"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), cover["subtitle"],
                  size=14, bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.45)
    intro_h = estimate_text_height(cover.get("intro_copy", ""), 12, CONTENT_W)
    add_text(slide, MARGIN, y, CONTENT_W, intro_h, cover.get("intro_copy", ""),
              size=12, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
    y += intro_h
    if watermark_label:
        add_text(slide, SLIDE_W - Inches(1.5), Inches(0.15), Inches(1.1), Inches(0.3),
                  watermark_label, size=11, bold=True, color=RGBColor(0xCC, 0xB0, 0x00),
                  align=PP_ALIGN.RIGHT)
    flow.y = y
    return slide


def build_destination_slides(flow, destinations, section_title=None, theme_line=None):
    """목적지 개수만큼만 슬라이드를 만든다. 고정 개수로 나누지 않고, 실제 텍스트
    길이를 추정해서 한 슬라이드에 들어갈 수 있는 만큼만 채우고 넘치면 다음 슬라이드로.
    첫 슬라이드는 이전 섹션이 남긴 여백에 이어 붙일 수 있으면 이어 붙인다."""
    if not destinations:
        return []
    slides = []
    idx = 0
    first_slide = True
    header_h = (Inches(0.55) if section_title else Inches(0)) + \
               (Inches(0.4) if theme_line else Inches(0))

    while idx < len(destinations):
        if first_slide:
            y = flow.ensure(header_h + Inches(0.5))  # 헤더 + 항목 하나 들어갈 여유
            slide = flow.slide
            if section_title:
                add_section_bar(slide, y, section_title)
                y += Inches(0.55)
            if theme_line:
                add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), theme_line,
                          size=14, bold=True, align=PP_ALIGN.CENTER)
                y += Inches(0.4)
            first_slide = False
        else:
            slide = flow.new_slide()
            y = flow.y

        placed_any = False
        while idx < len(destinations):
            dest = destinations[idx]
            region_tag = dest.get("region_tag")
            title_h = estimate_text_height(dest.get("title", ""), 15, CONTENT_W, bold=True)
            image_h = Inches(0.45)
            desc_h = estimate_text_height(dest.get("description", ""), 12, CONTENT_W)
            block_h = (Inches(0.28) if region_tag else Inches(0)) + title_h + Inches(0.06) \
                + image_h + Inches(0.1) + desc_h + Inches(0.18)

            if placed_any and y + block_h > flow.bottom_limit:
                break  # 이 슬라이드엔 더 안 들어감 -> 다음 슬라이드로

            if region_tag:
                add_text(slide, MARGIN, y, Inches(1.2), Inches(0.25), region_tag,
                          size=10, bold=True, color=WHITE)
                y += Inches(0.28)
            add_text(slide, MARGIN, y, CONTENT_W, title_h, dest.get("title", ""),
                      size=15, bold=True)
            y += title_h + Inches(0.06)
            add_small_image_placeholder(slide, y, Inches(1.6), image_h, "이미지")
            y += image_h + Inches(0.1)
            add_text(slide, MARGIN, y, CONTENT_W, desc_h, dest.get("description", ""),
                      size=12, color=MUTED_COLOR)
            y += desc_h + Inches(0.18)

            placed_any = True
            idx += 1

        flow.y = y
        slides.append(slide)
    return slides


def build_background_slide(flow, background_story):
    """'차마고도란?' 같은 배경 이야기 섹션"""
    if not background_story:
        return None
    kicker = background_story.get("kicker", "")
    title = background_story.get("title", "")
    content = background_story.get("content", "")
    content_h = estimate_text_height(content, 12, CONTENT_W)
    total_h = (Inches(0.4) if kicker else Inches(0)) + (Inches(0.55) if title else Inches(0)) + content_h

    y = flow.ensure(total_h)
    slide = flow.slide
    if kicker:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), kicker,
                  size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += Inches(0.4)
    if title:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.5), title,
                  size=20, bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.55)
    add_text(slide, MARGIN, y, CONTENT_W, content_h, content, size=12, align=PP_ALIGN.CENTER)
    y += content_h
    flow.y = y
    return slide


def build_reasons_slide(flow, why_reasons, product_name=""):
    """'왜 사천성인가' 같은 이유 N가지 섹션.
    타이틀은 AI에게 맡기지 않고 "{상품명} 포인트 0N" 형태로 코드에서 자동 생성한다
    (개수 기반 기계적 표기라 AI보다 코드가 더 정확함)."""
    if not why_reasons:
        return None
    heading = f"{product_name} 포인트 {len(why_reasons):02d}".strip()
    reason_blocks = []
    total_h = Inches(0.6)
    for reason in why_reasons:
        title_h = estimate_text_height(reason.get("title", ""), 15, CONTENT_W, bold=True)
        content_h = estimate_text_height(reason.get("content", ""), 12, CONTENT_W)
        reason_blocks.append((title_h, content_h))
        total_h += title_h + Inches(0.07) + content_h + Inches(0.22)

    y = flow.ensure(total_h)
    slide = flow.slide
    add_section_bar(slide, y, heading)
    y += Inches(0.6)
    for reason, (title_h, content_h) in zip(why_reasons, reason_blocks):
        add_text(slide, MARGIN, y, CONTENT_W, title_h, reason.get("title", ""),
                  size=15, bold=True, color=ACCENT_COLOR, align=PP_ALIGN.CENTER)
        y += title_h + Inches(0.07)
        add_text(slide, MARGIN, y, CONTENT_W, content_h, reason.get("content", ""),
                  size=12, align=PP_ALIGN.CENTER)
        y += content_h + Inches(0.22)
    flow.y = y
    return slide


def build_transport_slide(flow, transport_spec):
    """열차/크루즈처럼 이동수단 자체가 상품의 핵심 매력인 경우의 스펙 섹션.
    안나푸르나(2296 남극 크루즈), 호주 더 간 열차(1827) 상품설명 이미지 분석에서
    반복 확인된 "이동수단 스펙표"(객실타입/부대시설/톤수/안전등급 등) 패턴을 반영."""
    if not transport_spec or not transport_spec.get("specs"):
        return None
    image_h = Inches(0.95)
    rows = []
    total_h = (Inches(0.55) if transport_spec.get("title") else Inches(0)) + image_h
    for spec in transport_spec["specs"]:
        label = spec.get("label", "")
        value = spec.get("value", "")
        row_h = estimate_text_height(f"{label}: {value}", 12, CONTENT_W)
        rows.append((label, value, row_h))
        total_h += row_h + Inches(0.07)

    y = flow.ensure(total_h)
    slide = flow.slide
    if transport_spec.get("title"):
        add_section_bar(slide, y, transport_spec["title"])
        y += Inches(0.55)
    add_small_image_placeholder(slide, y, Inches(2.2), Inches(0.85), "이동수단 이미지")
    y += image_h
    for label, value, row_h in rows:
        add_text(slide, MARGIN, y, Inches(1.6), row_h, label, size=12, bold=True, color=ACCENT_COLOR)
        add_text(slide, MARGIN + Inches(1.7), y, CONTENT_W - Inches(1.7), row_h, value, size=12)
        y += row_h + Inches(0.07)
    flow.y = y
    return slide


def build_guide_slide(flow, guide_profile):
    """인솔자/가이드/담당 임원 프로필 섹션. 제주도 가이드 이력, 산티아고 인솔자
    경력 카드, 트레킹(킬리만자로 40회 등정 임원) 상품설명 이미지에서 반복 확인된
    "회사 구성원 신뢰 요소" 패턴을 반영."""
    if not guide_profile:
        return None
    header_h = Inches(0.55)
    bio_heights = []
    total_h = header_h
    for guide in guide_profile:
        bio_h = estimate_text_height(guide.get("bio", ""), 11, CONTENT_W)
        bio_heights.append(bio_h)
        total_h += Inches(1.18) + Inches(0.25) + bio_h + Inches(0.18)

    y = flow.ensure(total_h)
    slide = flow.slide
    add_section_bar(slide, y, "함께하는 사람들")
    y += header_h
    for guide, bio_h in zip(guide_profile, bio_heights):
        add_small_image_placeholder(slide, y, Inches(1.1), Inches(1.1), "프로필 사진")
        y += Inches(1.18)
        name_title = guide.get("name", "")
        if guide.get("title"):
            name_title = f"{name_title}  ({guide['title']})" if name_title else guide["title"]
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.28), name_title, size=13, bold=True,
                  align=PP_ALIGN.CENTER)
        y += Inches(0.25)
        add_text(slide, MARGIN, y, CONTENT_W, bio_h, guide.get("bio", ""), size=11,
                  color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += bio_h + Inches(0.18)
    flow.y = y
    return slide


def build_meal_slide(flow, meal_info):
    """"트레킹/여행 중 식사는 어떻게 하나요?" 실용 정보 Q&A 섹션. 일본알프스,
    키르기즈스탄, 마칼루, 하얼빈 등 지역이 전혀 다른 다수 상품에서 반복 확인된
    패턴으로, safety_note와 동일한 question/answer 구조를 재사용."""
    if not meal_info or not meal_info.get("question"):
        return None
    ans_h = estimate_text_height(meal_info.get("answer", ""), 12, CONTENT_W)
    total_h = Inches(0.5) + ans_h

    y = flow.ensure(total_h)
    slide = flow.slide
    add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), meal_info["question"], size=15,
              bold=True, align=PP_ALIGN.CENTER)
    y += Inches(0.5)
    add_text(slide, MARGIN, y, CONTENT_W, ans_h, meal_info.get("answer", ""), size=12,
              align=PP_ALIGN.CENTER)
    y += ans_h
    flow.y = y
    return slide


def build_route_compare_slide(flow, route_compare):
    """두 노선/코스를 비교하는 표 섹션"""
    if not route_compare or not route_compare.get("routes"):
        return None
    routes = route_compare["routes"]
    row_h = Inches(0.9)
    header_h = Inches(0.6) if route_compare.get("title") else Inches(0)
    total_h = header_h + Inches(0.5) + row_h * 4  # 이름줄(0.5) + 기준 4행

    y = flow.ensure(total_h)
    slide = flow.slide
    if route_compare.get("title"):
        add_section_bar(slide, y, route_compare["title"])
        y += Inches(0.6)
    col_w = CONTENT_W / len(routes)
    criteria = ["course", "scenery", "appeal", "summary"]
    criteria_label = {"course": "코스", "scenery": "풍경", "appeal": "매력", "summary": "한줄 요약"}
    for ri, route in enumerate(routes):
        x = MARGIN + col_w * ri
        add_text(slide, x, y, col_w, Inches(0.4), route.get("name", ""), size=14, bold=True,
                  align=PP_ALIGN.CENTER, color=ACCENT_COLOR)
    y += Inches(0.5)
    for crit in criteria:
        for ri, route in enumerate(routes):
            x = MARGIN + col_w * ri
            val = route.get(crit, "")
            add_text(slide, x, y, col_w, row_h, f"[{criteria_label[crit]}]\n{val}", size=10,
                      align=PP_ALIGN.CENTER)
        y += row_h
    flow.y = y
    return slide


def build_experience_slide(flow, brand_tagline, experience_points):
    """브랜드 소구 문구 + 경험 포인트(아이콘 카드 N개).
    예전엔 brand_points(불릿 목록)를 따로 받아 여기 같이 나열했는데, experience_points와
    내용이 거의 그대로 중복되는 문제가 있어(예: '노쇼핑/노옵션'이 두 번 나옴) brand_points는
    제거하고 experience_points 하나로 통일한다."""
    if not brand_tagline and not experience_points:
        return None
    tagline_h = estimate_text_height(brand_tagline, 16, CONTENT_W, bold=True) if brand_tagline else Inches(0)
    col_w = CONTENT_W / len(experience_points) if experience_points else CONTENT_W
    desc_h = Inches(0)
    if experience_points:
        desc_h = max(
            estimate_text_height(ep.get("description", ""), 10, col_w - Inches(0.1))
            for ep in experience_points
        )
    total_h = (tagline_h + Inches(0.2) if brand_tagline else Inches(0)) \
        + (Inches(0.65) + Inches(0.3) + desc_h if experience_points else Inches(0))

    y = flow.ensure(total_h)
    slide = flow.slide
    if brand_tagline:
        add_text(slide, MARGIN, y, CONTENT_W, tagline_h, brand_tagline, size=16, bold=True,
                  align=PP_ALIGN.CENTER)
        y += tagline_h + Inches(0.2)
    if experience_points:
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            add_image_placeholder(slide, x + Inches(0.05), y, col_w - Inches(0.1), Inches(0.55), "아이콘")
        y += Inches(0.65)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            th = estimate_text_height(ep.get("title", ""), 12, col_w - Inches(0.1), bold=True)
            add_text(slide, x, y, col_w - Inches(0.1), th, ep.get("title", ""), size=12,
                      bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.3)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            dh = estimate_text_height(ep.get("description", ""), 10, col_w - Inches(0.1))
            add_text(slide, x, y, col_w - Inches(0.1), dh, ep.get("description", ""), size=10,
                      color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += desc_h
    flow.y = y
    return slide


def build_highlights_slides(flow, highlights, heading=None):
    """번호 매긴 여정 하이라이트 카드 (destinations와 별개 — 더 큰 테마 단위)"""
    if not highlights:
        return []
    heading = heading or "여정 하이라이트"  # AI가 빠뜨려도 타이틀 없는 슬라이드가 나가지 않도록 기본값
    slides = []
    idx = 0
    first = True
    header_h = Inches(0.55)

    while idx < len(highlights):
        if first:
            y = flow.ensure(header_h + Inches(0.5))
            slide = flow.slide
            add_section_bar(slide, y, heading)
            y += header_h
            first = False
        else:
            slide = flow.new_slide()
            y = flow.y

        placed_any = False
        while idx < len(highlights):
            item = highlights[idx]
            num_label = f"{idx + 1:02d}"
            title_h = estimate_text_height(item.get("title", ""), 14, CONTENT_W, bold=True)
            image_h = Inches(0.45)
            desc_h = estimate_text_height(item.get("description", ""), 11, CONTENT_W)
            block_h = Inches(0.22) + title_h + Inches(0.06) + image_h + Inches(0.1) + desc_h + Inches(0.18)
            if placed_any and y + block_h > flow.bottom_limit:
                break
            add_text(slide, MARGIN, y, Inches(0.6), Inches(0.25), num_label, size=13, bold=True,
                      color=ACCENT_COLOR)
            y += Inches(0.26)
            add_text(slide, MARGIN, y, CONTENT_W, title_h, item.get("title", ""), size=14, bold=True)
            y += title_h + Inches(0.06)
            add_small_image_placeholder(slide, y, Inches(1.6), image_h, "이미지")
            y += image_h + Inches(0.1)
            add_text(slide, MARGIN, y, CONTENT_W, desc_h, item.get("description", ""), size=11,
                      color=MUTED_COLOR)
            y += desc_h + Inches(0.18)
            placed_any = True
            idx += 1

        flow.y = y
        slides.append(slide)
    return slides


def build_season_slide(flow, season, season_table=None):
    if not season or (not season.get("content") and not season_table):
        return None
    header_h = Inches(0.6)
    stat_h = Inches(0.5) if season.get("stat_line") else Inches(0)
    content_h = estimate_text_height(season.get("content", ""), 12, CONTENT_W)
    table_h = Inches(0.85) + Inches(0.3) if season_table else Inches(0)
    total_h = header_h + stat_h + content_h + Inches(0.3) + table_h

    y = flow.ensure(total_h)
    slide = flow.slide
    add_section_bar(slide, y, season.get("title", "언제 가면 좋을까?"))
    y += header_h
    if season.get("stat_line"):
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, y, CONTENT_W, Inches(0.4))
        bar.fill.solid()
        bar.fill.fore_color.rgb = ACCENT_COLOR
        bar.line.fill.background()
        _tf_setup(bar.text_frame, season["stat_line"], 13, WHITE, bold=True, align=PP_ALIGN.CENTER)
        bar.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        y += Inches(0.5)
    add_text(slide, MARGIN, y, CONTENT_W, content_h, season.get("content", ""), size=12)
    y += content_h + Inches(0.3)
    if season_table:
        add_image_placeholder(slide, MARGIN, y, CONTENT_W, Inches(0.7),
                               "월별 기온 차트 (자리표시 — 실제 그래픽은 디자이너 작업)")
        y += Inches(0.85)
        header = "  ".join(f"{row.get('month','')}" for row in season_table)
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.3), header, size=10, color=MUTED_COLOR,
                  align=PP_ALIGN.CENTER)
        y += Inches(0.3)
    flow.y = y
    return slide


def build_safety_slide(flow, altitude_profile, safety_note):
    """경유지 고도 프로필(있는 경우) + 안전/난이도 관련 표준 안내.
    고산 트레킹의 '고산증', 도보순례의 '체력/보험', 일반 하이킹의 '난이도' 등
    카테고리에 따라 톤이 다른 표준 안내문을 담는 범용 섹션."""
    if not altitude_profile and not safety_note:
        return None
    ans_h = Inches(0)
    qa_h = Inches(0)
    if safety_note and safety_note.get("question"):
        ans_h = estimate_text_height(safety_note.get("answer", ""), 12, CONTENT_W)
        qa_h = Inches(0.5) + ans_h + Inches(0.35)
    profile_h = (Inches(0.4) + Inches(0.6) + Inches(0.65)) if altitude_profile else Inches(0)
    total_h = qa_h + profile_h

    y = flow.ensure(total_h)
    slide = flow.slide
    if safety_note and safety_note.get("question"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), safety_note["question"], size=15,
                  bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.5)
        add_text(slide, MARGIN, y, CONTENT_W, ans_h, safety_note.get("answer", ""), size=12,
                  align=PP_ALIGN.CENTER)
        y += ans_h + Inches(0.35)
    if altitude_profile:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.3),
                  "구간별 고도 프로필 (자리표시 — 실제 그래픽은 디자이너 작업)",
                  size=10, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += Inches(0.4)
        n = len(altitude_profile)
        col_w = CONTENT_W / max(n, 1)
        gap = Inches(0.06)
        for i, stop in enumerate(altitude_profile):
            x = MARGIN + col_w * i
            add_image_placeholder(slide, x + gap, y, col_w - gap * 2, Inches(0.5), "숙박")
        y += Inches(0.6)
        for i, stop in enumerate(altitude_profile):
            x = MARGIN + col_w * i
            label = f"{stop.get('name','')}\n{stop.get('altitude','')}"
            extra = " / ".join(v for v in (stop.get("distance"), stop.get("duration")) if v)
            if extra:
                label += f"\n{extra}"
            add_text(slide, x + gap, y, col_w - gap * 2, Inches(0.65), label, size=9, align=PP_ALIGN.CENTER)
        y += Inches(0.65)
    flow.y = y
    return slide


def build_banner_request_slide(flow, cover):
    """배너 기획 페이지 — 실제 회사 배너제작 템플릿(배너제작.pptx)의 레이아웃을
    그대로 재현. 4개 배너 슬롯(메인 와이드/서브메인 띠/서브메인 2단/지역 리스트)에
    각각 이미지 자리와 '태그라인+타이틀' 텍스트를 넣는다. 스펙 라벨/안내 문구는
    회사 표준이라 고정값이며, 지역/상품과 무관하게 항상 그대로 포함. 절대 위치로
    실제 배너 템플릿과 맞춰야 해서 다른 섹션과 공유하지 않고 항상 전용 슬라이드로 만든다."""
    slide = flow.new_slide()
    tagline_title = f"{cover.get('tagline','')}\n{cover.get('product_name','')}"

    add_section_bar(slide, Inches(0), "배너 기획", height=Inches(0.47), size=14)
    add_text(slide, Inches(0.106), Inches(0.675), Inches(1.717), Inches(0.404),
              "배너제작요청", size=13, bold=True)
    add_text(slide, Inches(1.823), Inches(0.733), Inches(4.672), Inches(0.303),
              "홈페이지 개편에 따라, 배너 디자인이 전면 교체되었습니다.", size=9, color=MUTED_COLOR)

    # 메인 와이드 배너
    add_text(slide, Inches(0.106), Inches(1.271), Inches(5.701), Inches(0.303),
              "메인 와이드 배너 (PC: 1920x700, MO: 750x510), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.217), Inches(1.630), Inches(5.475), Inches(1.817), "이미지")
    add_text(slide, Inches(0.388), Inches(2.158), Inches(5.133), Inches(0.640),
              tagline_title, size=13, bold=True, align=PP_ALIGN.CENTER)

    # 서브메인 띠배너
    add_text(slide, Inches(0.081), Inches(3.581), Inches(5.642), Inches(0.303),
              "서브메인 띠배너 (PC: 1920x200, MO: 750x200), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.136), Inches(4.021), Inches(7.028), Inches(0.979), "이미지")
    add_text(slide, Inches(1.184), Inches(4.173), Inches(5.133), Inches(0.640),
              tagline_title, size=13, bold=True, align=PP_ALIGN.CENTER)

    # 서브메인 2단배너
    add_text(slide, Inches(0.136), Inches(5.137), Inches(5.642), Inches(0.303),
              "서브메인 2단배너 (PC: 590x370, MO: 585x670), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.221), Inches(5.521), Inches(2.835), Inches(1.771), "이미지")
    add_text(slide, Inches(0.288), Inches(6.138), Inches(2.835), Inches(0.454),
              tagline_title, size=11, bold=True, align=PP_ALIGN.CENTER)

    # 지역 리스트 배너
    add_text(slide, Inches(0.205), Inches(7.481), Inches(5.701), Inches(0.303),
              "지역 리스트 배너 (PC: 1200x207, MO: 750x207), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.316), Inches(7.989), Inches(6.871), Inches(1.336), "이미지")
    add_text(slide, Inches(1.184), Inches(8.280), Inches(5.133), Inches(0.640),
              tagline_title, size=13, bold=True, align=PP_ALIGN.CENTER)

    return slide


def build(content_json, out_path):
    """content_json(정현지 스키마) -> 새 PPTX 파일 생성"""
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    flow = SlideFlow(prs)

    cover = content_json.get("cover", {})
    build_cover_slide(flow, cover, content_json.get("watermark_label", ""))
    build_background_slide(flow, content_json.get("background_story"))
    build_reasons_slide(flow, content_json.get("why_reasons"), product_name=cover.get("product_name", ""))
    build_destination_slides(
        flow,
        content_json.get("destinations", []),
        section_title=content_json.get("destinations_heading"),
        theme_line=None,
    )
    build_route_compare_slide(flow, content_json.get("route_compare"))
    build_transport_slide(flow, content_json.get("transport_spec"))
    build_experience_slide(
        flow,
        content_json.get("brand_tagline", ""),
        content_json.get("experience_points"),
    )
    build_guide_slide(flow, content_json.get("guide_profile"))
    build_highlights_slides(
        flow,
        content_json.get("highlights"),
        heading=content_json.get("highlights_heading"),
    )
    build_season_slide(flow, content_json.get("season", {}), content_json.get("season_table"))
    build_meal_slide(flow, content_json.get("meal_info"))
    build_safety_slide(flow, content_json.get("altitude_profile"), content_json.get("safety_note"))
    build_banner_request_slide(flow, cover)
    prs.save(out_path)
    return prs
