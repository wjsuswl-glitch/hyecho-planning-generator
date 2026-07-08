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


def estimate_text_height(text, size_pt, width_emu, line_spacing=1.35, bold=False):
    """글자 수 기반으로 텍스트가 실제로 차지할 높이를 대략 추정한다.
    고정 간격 대신 이걸 써야 설명 길이에 따라 다음 요소와 안 겹친다.
    실제 렌더링 폭 추정은 부정확할 수 있어 여유 마진을 넉넉히 둔다."""
    if not text:
        return Inches(0.15)
    width_in = Emu(width_emu).inches
    # 한글 기준 글자 폭 대략치 (볼드면 좀 더 넓게 잡음), 안전 마진 포함
    char_w_in = (size_pt / 72) * (1.05 if bold else 0.95)
    chars_per_line = max(1, int(width_in / char_w_in))
    total_lines = 0
    for line in str(text).split("\n"):
        total_lines += max(1, -(-len(line) // chars_per_line))  # ceil
    line_h_in = (size_pt / 72) * line_spacing
    return Inches(total_lines * line_h_in + 0.15)  # 여유 마진 추가


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


def new_presentation():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # 완전 빈 레이아웃


# ---------------------------------------------------------------------------
# 슬라이드 빌더 함수 (정현지 스타일)
# ---------------------------------------------------------------------------

def build_cover_slide(prs, cover, watermark_label=""):
    slide = _blank_slide(prs)
    y = Inches(0.5)
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
    y += Inches(0.15)
    add_image_placeholder(slide, MARGIN, y, CONTENT_W, Inches(3.2), "메인 이미지")
    y += Inches(3.4)
    if cover.get("subtitle"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), cover["subtitle"],
                  size=14, bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.45)
    intro_h = estimate_text_height(cover.get("intro_copy", ""), 12, CONTENT_W)
    add_text(slide, MARGIN, y, CONTENT_W, intro_h, cover.get("intro_copy", ""),
              size=12, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
    if watermark_label:
        add_text(slide, SLIDE_W - Inches(1.5), Inches(0.15), Inches(1.1), Inches(0.3),
                  watermark_label, size=11, bold=True, color=RGBColor(0xCC, 0xB0, 0x00),
                  align=PP_ALIGN.RIGHT)
    return slide



def build_destination_slides(prs, destinations, section_title=None, theme_line=None, per_slide=None):
    """목적지 개수만큼만 슬라이드를 만든다. 고정 개수로 나누지 않고, 실제 텍스트
    길이를 추정해서 한 슬라이드에 들어갈 수 있는 만큼만 채우고 넘치면 다음 슬라이드로."""
    if not destinations:
        return []
    slides = []
    bottom_limit = SLIDE_H - Inches(0.3)
    idx = 0
    first_slide = True
    while idx < len(destinations):
        slide = _blank_slide(prs)
        y = Inches(0.4)
        if first_slide:
            if section_title:
                add_section_bar(slide, y, section_title)
                y += Inches(0.6)
            if theme_line:
                add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), theme_line,
                          size=14, bold=True, align=PP_ALIGN.CENTER)
                y += Inches(0.5)
            first_slide = False

        placed_any = False
        while idx < len(destinations):
            dest = destinations[idx]
            region_tag = dest.get("region_tag")
            title_h = estimate_text_height(dest.get("title", ""), 15, CONTENT_W, bold=True)
            image_h = Inches(1.6)
            desc_h = estimate_text_height(dest.get("description", ""), 12, CONTENT_W)
            block_h = (Inches(0.35) if region_tag else Inches(0)) + title_h + Inches(0.1) \
                + image_h + Inches(0.15) + desc_h + Inches(0.3)

            if placed_any and y + block_h > bottom_limit:
                break  # 이 슬라이드엔 더 안 들어감 -> 다음 슬라이드로

            if region_tag:
                add_text(slide, MARGIN, y, Inches(1.2), Inches(0.3), region_tag,
                          size=10, bold=True, color=WHITE)
                y += Inches(0.35)
            add_text(slide, MARGIN, y, CONTENT_W, title_h, dest.get("title", ""),
                      size=15, bold=True)
            y += title_h + Inches(0.1)
            add_image_placeholder(slide, MARGIN, y, CONTENT_W, image_h, "이미지")
            y += image_h + Inches(0.15)
            add_text(slide, MARGIN, y, CONTENT_W, desc_h, dest.get("description", ""),
                      size=12, color=MUTED_COLOR)
            y += desc_h + Inches(0.3)

            placed_any = True
            idx += 1

        slides.append(slide)
    return slides


def build_background_slide(prs, background_story):
    """'차마고도란?' 같은 배경 이야기 슬라이드"""
    if not background_story:
        return None
    slide = _blank_slide(prs)
    y = Inches(0.4)
    if background_story.get("kicker"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), background_story["kicker"],
                  size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += Inches(0.4)
    if background_story.get("title"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.5), background_story["title"],
                  size=20, bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.55)
    content = background_story.get("content", "")
    h = estimate_text_height(content, 12, CONTENT_W)
    add_text(slide, MARGIN, y, CONTENT_W, h, content, size=12, align=PP_ALIGN.CENTER)
    return slide


def build_reasons_slide(prs, why_reasons):
    """'왜 사천성인가' 같은 이유 N가지 슬라이드"""
    if not why_reasons:
        return None
    slide = _blank_slide(prs)
    y = Inches(0.4)
    for reason in why_reasons:
        title_h = estimate_text_height(reason.get("title", ""), 15, CONTENT_W, bold=True)
        add_text(slide, MARGIN, y, CONTENT_W, title_h, reason.get("title", ""),
                  size=15, bold=True, color=ACCENT_COLOR, align=PP_ALIGN.CENTER)
        y += title_h + Inches(0.1)
        content_h = estimate_text_height(reason.get("content", ""), 12, CONTENT_W)
        add_text(slide, MARGIN, y, CONTENT_W, content_h, reason.get("content", ""),
                  size=12, align=PP_ALIGN.CENTER)
        y += content_h + Inches(0.35)
    return slide


def build_route_compare_slide(prs, route_compare):
    """두 노선/코스를 비교하는 표 슬라이드"""
    if not route_compare or not route_compare.get("routes"):
        return None
    slide = _blank_slide(prs)
    y = Inches(0.4)
    if route_compare.get("title"):
        add_section_bar(slide, y, route_compare["title"])
        y += Inches(0.6)
    routes = route_compare["routes"]
    col_w = CONTENT_W / len(routes)
    criteria = ["course", "scenery", "appeal", "summary"]
    criteria_label = {"course": "코스", "scenery": "풍경", "appeal": "매력", "summary": "한줄 요약"}
    row_h = Inches(0.9)
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
    return slide


def build_experience_slide(prs, brand_tagline, brand_points, experience_points):
    """브랜드 소구 + 경험 포인트(아이콘 카드 N개)"""
    slide = _blank_slide(prs)
    y = Inches(0.4)
    if brand_tagline:
        h = estimate_text_height(brand_tagline, 16, CONTENT_W, bold=True)
        add_text(slide, MARGIN, y, CONTENT_W, h, brand_tagline, size=16, bold=True,
                  align=PP_ALIGN.CENTER)
        y += h + Inches(0.2)
    for point in (brand_points or []):
        h = estimate_text_height(point, 13, CONTENT_W)
        add_text(slide, MARGIN, y, CONTENT_W, h, f"· {point}", size=13)
        y += h + Inches(0.1)
    y += Inches(0.3)
    if experience_points:
        col_w = CONTENT_W / len(experience_points)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            add_image_placeholder(slide, x + Inches(0.05), y, col_w - Inches(0.1), Inches(0.7), "아이콘")
        y += Inches(0.85)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            th = estimate_text_height(ep.get("title", ""), 12, col_w - Inches(0.1), bold=True)
            add_text(slide, x, y, col_w - Inches(0.1), th, ep.get("title", ""), size=12,
                      bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.4)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            dh = estimate_text_height(ep.get("description", ""), 10, col_w - Inches(0.1))
            add_text(slide, x, y, col_w - Inches(0.1), dh, ep.get("description", ""), size=10,
                      color=MUTED_COLOR, align=PP_ALIGN.CENTER)
    return slide


def build_highlights_slides(prs, highlights, heading=None):
    """번호 매긴 여정 하이라이트 카드 (destinations와 별개 — 더 큰 테마 단위)"""
    if not highlights:
        return []
    heading = heading or "여정 하이라이트"  # AI가 빠뜨려도 타이틀 없는 슬라이드가 나가지 않도록 기본값
    slides = []
    bottom_limit = SLIDE_H - Inches(0.3)
    idx = 0
    first = True
    while idx < len(highlights):
        slide = _blank_slide(prs)
        y = Inches(0.4)
        if first and heading:
            add_section_bar(slide, y, heading)
            y += Inches(0.6)
            first = False
        placed_any = False
        while idx < len(highlights):
            item = highlights[idx]
            num_label = f"{idx + 1:02d}"
            title_h = estimate_text_height(item.get("title", ""), 14, CONTENT_W, bold=True)
            image_h = Inches(1.5)
            desc_h = estimate_text_height(item.get("description", ""), 11, CONTENT_W)
            block_h = Inches(0.3) + title_h + Inches(0.1) + image_h + Inches(0.15) + desc_h + Inches(0.3)
            if placed_any and y + block_h > bottom_limit:
                break
            add_text(slide, MARGIN, y, Inches(0.6), Inches(0.3), num_label, size=13, bold=True,
                      color=ACCENT_COLOR)
            y += Inches(0.35)
            add_text(slide, MARGIN, y, CONTENT_W, title_h, item.get("title", ""), size=14, bold=True)
            y += title_h + Inches(0.1)
            add_image_placeholder(slide, MARGIN, y, CONTENT_W, image_h, "이미지")
            y += image_h + Inches(0.15)
            add_text(slide, MARGIN, y, CONTENT_W, desc_h, item.get("description", ""), size=11,
                      color=MUTED_COLOR)
            y += desc_h + Inches(0.3)
            placed_any = True
            idx += 1
        slides.append(slide)
    return slides


def build_safety_slide(prs, altitude_profile, safety_note):
    """경유지 고도 프로필(있는 경우) + 안전/난이도 관련 표준 안내.
    고산 트레킹의 '고산증', 도보순례의 '체력/보험', 일반 하이킹의 '난이도' 등
    카테고리에 따라 톤이 다른 표준 안내문을 담는 범용 슬라이드."""
    if not altitude_profile and not safety_note:
        return None
    slide = _blank_slide(prs)
    y = Inches(0.4)
    if safety_note and safety_note.get("question"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), safety_note["question"], size=15,
                  bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.5)
        ans_h = estimate_text_height(safety_note.get("answer", ""), 12, CONTENT_W)
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
            add_text(slide, x + gap, y, col_w - gap * 2, Inches(0.5), label, size=9, align=PP_ALIGN.CENTER)
    return slide


def build_season_slide(prs, season, season_table=None):
    slide = _blank_slide(prs)
    y = Inches(0.4)
    add_section_bar(slide, y, season.get("title", "언제 가면 좋을까?"))
    y += Inches(0.6)
    if season.get("stat_line"):
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, y, CONTENT_W, Inches(0.4))
        bar.fill.solid()
        bar.fill.fore_color.rgb = ACCENT_COLOR
        bar.line.fill.background()
        _tf_setup(bar.text_frame, season["stat_line"], 13, WHITE, bold=True, align=PP_ALIGN.CENTER)
        bar.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        y += Inches(0.5)
    content_h = estimate_text_height(season.get("content", ""), 12, CONTENT_W)
    add_text(slide, MARGIN, y, CONTENT_W, content_h, season.get("content", ""), size=12)
    y += content_h + Inches(0.3)
    if season_table:
        add_image_placeholder(slide, MARGIN, y, CONTENT_W, Inches(1.8),
                               "월별 기온 차트 (자리표시 — 실제 그래픽은 디자이너 작업)")
        y += Inches(1.95)
        header = "  ".join(f"{row.get('month','')}" for row in season_table)
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.3), header, size=10, color=MUTED_COLOR,
                  align=PP_ALIGN.CENTER)
    return slide


def build_banner_request_slide(prs, cover):
    """배너 기획 페이지 — 실제 회사 배너제작 템플릿(배너제작.pptx)의 레이아웃을
    그대로 재현. 4개 배너 슬롯(메인 와이드/서브메인 띠/서브메인 2단/지역 리스트)에
    각각 이미지 자리와 '태그라인+타이틀' 텍스트를 넣는다. 스펙 라벨/안내 문구는
    회사 표준이라 고정값이며, 지역/상품과 무관하게 항상 그대로 포함."""
    slide = _blank_slide(prs)
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


def build(content_json, out_path, per_slide=3):
    """content_json(정현지 스키마) -> 새 PPTX 파일 생성"""
    prs = new_presentation()
    cover = content_json.get("cover", {})
    build_cover_slide(prs, cover, content_json.get("watermark_label", ""))
    build_background_slide(prs, content_json.get("background_story"))
    build_reasons_slide(prs, content_json.get("why_reasons"))
    build_destination_slides(
        prs,
        content_json.get("destinations", []),
        section_title=content_json.get("destinations_heading"),
        theme_line=None,
        per_slide=per_slide,
    )
    build_route_compare_slide(prs, content_json.get("route_compare"))
    build_experience_slide(
        prs,
        content_json.get("brand_tagline", ""),
        content_json.get("brand_points", []),
        content_json.get("experience_points"),
    )
    build_highlights_slides(
        prs,
        content_json.get("highlights"),
        heading=content_json.get("highlights_heading"),
    )
    build_season_slide(prs, content_json.get("season", {}), content_json.get("season_table"))
    build_safety_slide(prs, content_json.get("altitude_profile"), content_json.get("safety_note"))
    build_banner_request_slide(prs, cover)
    prs.save(out_path)
    return prs
