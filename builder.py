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
    add_text(slide, MARGIN, y, CONTENT_W, Inches(1.0), cover.get("intro_copy", ""),
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
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), f"· {point}", size=13)
        y += Inches(0.45)
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


def build_destination_slides(prs, destinations, section_title=None, theme_line=None, per_slide=3):
    """목적지 개수만큼만 슬라이드를 만든다. per_slide개씩 묶어서 한 슬라이드에 배치."""
    if not destinations:
        return []
    slides = []
    chunks = [destinations[i:i + per_slide] for i in range(0, len(destinations), per_slide)]
    for ci, chunk in enumerate(chunks):
        slide = _blank_slide(prs)
        y = Inches(0.4)
        if ci == 0:
            if section_title:
                add_section_bar(slide, y, section_title)
                y += Inches(0.6)
            if theme_line:
                add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), theme_line,
                          size=14, bold=True, align=PP_ALIGN.CENTER)
                y += Inches(0.5)
        for dest in chunk:
            region_tag = dest.get("region_tag")
            if region_tag:
                add_text(slide, MARGIN, y, Inches(1.2), Inches(0.3), region_tag,
                          size=10, bold=True, color=WHITE)
                y += Inches(0.05)
            add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), dest.get("title", ""),
                      size=15, bold=True)
            y += Inches(0.45)
            add_image_placeholder(slide, MARGIN, y, CONTENT_W, Inches(1.6), "이미지")
            y += Inches(1.75)
            add_text(slide, MARGIN, y, CONTENT_W, Inches(0.6), dest.get("description", ""),
                      size=12, color=MUTED_COLOR)
            y += Inches(0.85)
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
    add_text(slide, MARGIN, y, CONTENT_W, Inches(1.2), season.get("content", ""), size=12)
    return slide


def build_banner_request_slide(prs, product_name):
    """배너 기획 페이지 — 지역/상품과 무관하게 항상 고정 포함.
    디자이너에게 필요한 배너 소재 요청 목록만 담고, 톤앤매너 코멘트 등은
    넣지 않는다 (기획자가 결과물을 보고 직접 남길 예정)."""
    slide = _blank_slide(prs)
    add_section_bar(slide, Inches(0.4), "배너 기획")
    y = Inches(1.0)
    add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4), product_name, size=14, bold=True,
              align=PP_ALIGN.CENTER)
    y += Inches(0.6)
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
