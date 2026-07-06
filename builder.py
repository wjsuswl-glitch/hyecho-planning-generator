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


def estimate_text_height(text, size_pt, width_emu, line_spacing=1.35):
    """글자 수 기반으로 텍스트가 실제로 차지할 높이를 대략 추정한다.
    고정 간격 대신 이걸 써야 설명 길이에 따라 다음 요소와 안 겹친다."""
    if not text:
        return Inches(0.15)
    width_in = Emu(width_emu).inches
    # 한글 기준 12pt 글자 폭 대략치 (완전 정확하진 않지만 겹침 방지엔 충분)
    char_w_in = (size_pt / 72) * 0.95
    chars_per_line = max(1, int(width_in / char_w_in))
    total_lines = 0
    for line in str(text).split("\n"):
        total_lines += max(1, -(-len(line) // chars_per_line))  # ceil
    line_h_in = (size_pt / 72) * line_spacing
    return Inches(total_lines * line_h_in + 0.05)


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
        add_text(slide, SLIDE_W - Inches(1.5), Inches(0.3), Inches(1.1), Inches(0.3),
                  watermark_label, size=11, bold=True, color=RGBColor(0xCC, 0xB0, 0x00),
                  align=PP_ALIGN.RIGHT)
    return slide


def build_brand_slide(prs, brand_tagline, brand_points, why_hyecho):
    slide = _blank_slide(prs)
    y = Inches(0.4)
    if brand_tagline:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.5), brand_tagline,
                  size=16, bold=True, align=PP_ALIGN.CENTER)
        y += Inches(0.6)
    for point in (brand_points or []):
        h = estimate_text_height(point, 13, CONTENT_W)
        add_text(slide, MARGIN, y, CONTENT_W, h, f"· {point}", size=13)
        y += h + Inches(0.1)
    y += Inches(0.2)
    if why_hyecho:
        add_section_bar(slide, y, why_hyecho.get("section_title", "혜초와 함께라면"))
        y += Inches(0.6)
        if why_hyecho.get("subtitle1"):
            add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), why_hyecho["subtitle1"], size=13)
            y += Inches(0.4)
        if why_hyecho.get("subtitle2"):
            add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), why_hyecho["subtitle2"], size=13)
            y += Inches(0.4)
        if why_hyecho.get("badge"):
            add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), why_hyecho["badge"],
                      size=13, bold=True, color=ACCENT_COLOR)
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
            title_h = estimate_text_height(dest.get("title", ""), 15, CONTENT_W)
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


def build_season_slide(prs, season):
    slide = _blank_slide(prs)
    y = Inches(0.4)
    add_section_bar(slide, y, season.get("title", "언제 가면 좋을까?"))
    y += Inches(0.6)
    if season.get("stat_line"):
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), season["stat_line"],
                  size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bar = slide.shapes[-1]
        bar.fill.solid()
        bar.fill.fore_color.rgb = ACCENT_COLOR
        y += Inches(0.5)
    add_text(slide, MARGIN, y, CONTENT_W, estimate_text_height(season.get("content",""), 12, CONTENT_W),
              season.get("content", ""), size=12)
    return slide


def build_banner_request_slide(prs, product_name):
    """배너 기획 페이지 — 지역/상품과 무관하게 항상 고정 포함.
    디자이너에게 필요한 배너 소재 요청 목록만 담고, 톤앤매너 코멘트 등은
    넣지 않는다 (기획자가 결과물을 보고 직접 남길 예정)."""
    slide = _blank_slide(prs)
    add_section_bar(slide, Inches(0.4), "배너 기획")
    y = Inches(1.0)
    add_text(slide, MARGIN, y, CONTENT_W, Inches(0.6), product_name, size=20, bold=True,
              align=PP_ALIGN.CENTER)
    y += Inches(0.8)
    for label in ["메인 배너 (가로형)", "서브 배너 (정사각형)", "썸네일 배너"]:
        add_image_placeholder(slide, MARGIN, y, CONTENT_W, Inches(1.3), label)
        y += Inches(1.5)
    return slide


def build(content_json, out_path, per_slide=3):
    """content_json(정현지 스키마) -> 새 PPTX 파일 생성"""
    prs = new_presentation()
    cover = content_json.get("cover", {})
    build_cover_slide(prs, cover, content_json.get("watermark_label", ""))
    build_brand_slide(
        prs,
        content_json.get("brand_tagline", ""),
        content_json.get("brand_points", []),
        content_json.get("why_hyecho", {}),
    )
    build_destination_slides(
        prs,
        content_json.get("destinations", []),
        section_title=content_json.get("why_hyecho", {}).get("section_title"),
        theme_line=content_json.get("why_hyecho", {}).get("theme_line"),
        per_slide=per_slide,
    )
    build_season_slide(prs, content_json.get("season", {}))
    build_banner_request_slide(prs, cover.get("product_name", ""))
    prs.save(out_path)
    return prs
