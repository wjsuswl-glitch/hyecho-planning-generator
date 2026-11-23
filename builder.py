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
GROUP_BAR_COLOR = RGBColor(0x5C, 0x8A, 0x9E)  # ACCENT_COLOR보다 밝은 티얼. 방문지 그룹(상위
                                                # 카테고리) 바에 쓴다 — 섹션 바(진한 색)보다
                                                # 밝게 해서 "섹션 > 그룹 > 방문지" 위계가
                                                # 색 대비로도 바로 보이게 한다.
TEXT_COLOR = RGBColor(0x22, 0x22, 0x22)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED_COLOR = RGBColor(0x66, 0x66, 0x66)
MARGIN = Inches(0.4)
CONTENT_W = SLIDE_W - MARGIN * 2

# ---- 모바일 최적화 기획안 전용 스타일 상수 ----
# "모바일 최적화 콘텐츠 가이드라인_20260702.pdf"(마케팅실 최정인) + 실제 완성
# 샘플(17.모바일최적화/기존이미지 vs 신규이미지, 돌로미테·그랜드서클 등) 대조로
# 확인한 방향을 PPT 기획안 스케일로 반영한 값. 실제 최종 이미지는 1150px 폭
# 기준 32pt 본문/22pt 최소 캡션/160~180% 줄간격/80~100px 여백을 쓰지만, 이
# 앱의 산출물은 (디자이너가 참고해 최종 그래픽을 만드는) "기획안"이지 최종
# 이미지 자체가 아니므로, 그 픽셀 스펙을 그대로 옮기는 대신 PC용 기존 폰트
# 크기 대비 확대 비율로 환산해 적용한다. 실제 Before/After 대조 결과, 두
# 버전의 텍스트 내용은 100% 동일하고 폰트/여백/배지 스타일/배경색만 달랐다
# (요약·축약이 아니라 "재배치") — 그래서 스키마는 PC용과 동일한 걸 그대로
# 쓰고, 이 스타일 상수 + build_mobile_* 함수들로 시각 표현만 바꾼다.
MOBILE_BODY_SIZE = 16        # PC 본문 12pt -> 16pt
MOBILE_CAPTION_SIZE = 13     # PC 캡션 10~11pt -> 13pt
MOBILE_TITLE_SIZE = 19       # PC 소제목 14~15pt -> 19pt
MOBILE_LINE_SPACING = 1.7    # 가이드라인 160~180% 구간의 중간값 — 실제 렌더링(add_text의
                              # line_spacing 인자)에 이 값을 그대로 쓴다.
MOBILE_LINE_SPACING_EST = 2.1  # estimate_text_height 전용 보정값. python-pptx/PowerPoint의
                                # line_spacing(spcPct)은 "폰트 크기의 N배"가 아니라 "기본
                                # 한 줄 간격(이미 폰트 크기보다 넉넉함, 이 파일 다른 곳의
                                # 기본값 1.22가 그 기준)의 N배"로 적용된다 — 그래서 실제
                                # 렌더링에 MOBILE_LINE_SPACING(1.7)을 쓰면 실제 줄 높이는
                                # estimate_text_height(line_spacing=1.7)로 추정한 것보다 더
                                # 크다. 모바일 카드 그리드(build_mobile_experience_slide)에서
                                # 이 오차 때문에 텍스트가 카드 밖으로 넘치는 걸 실제 렌더링
                                # 대조로 확인 후 보정(1.7*1.22≈2.07, 여유 포함 2.1로 반올림).
MOBILE_BG_COLOR = RGBColor(0xEE, 0xF5, 0xF8)  # 완전 흰색 배경 대신 옅은 브랜드톤 배경
                                                # (가이드라인: "플랫한 흰 배경 지양")
MOBILE_BADGE_COLOR = RGBColor(0xE0, 0x7A, 0x3E)  # 배지(태그)용 강조색 — ACCENT_COLOR와
                                                   # 구분되는 톤으로, 배지가 "포인트
                                                   # 표시"라는 게 한눈에 보이게 한다
MOBILE_CARD_COLOR = RGBColor(0xFF, 0xFF, 0xFF)    # 특별함 카드 배경(옅은 바탕 위 흰 카드)


def _tf_setup(tf, text, size, color, bold=False, align=PP_ALIGN.LEFT, font=FONT_NAME,
              line_spacing=None):
    tf.word_wrap = True
    lines = str(text).split("\n")
    tf.text = lines[0]
    p0 = tf.paragraphs[0]
    p0.alignment = align
    if line_spacing is not None:
        p0.line_spacing = line_spacing
    for run in p0.runs or [p0.add_run()]:
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.text = line
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
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
             bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.vertical_anchor = anchor
    _tf_setup(tf, text, size, color, bold, align, line_spacing=line_spacing)
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


def group_bar_height(text, size=13):
    """add_group_bar가 실제로 그릴 높이를 미리 계산한다 — 배치 전에 다음 요소와
    겹치는지 판단하려면 그리기 전에 높이를 알아야 해서 별도 함수로 분리."""
    return estimate_text_height(text, size, CONTENT_W - Inches(0.3), bold=True) + Inches(0.16)


def add_group_bar(slide, top, text, height=None, size=13):
    """방문지를 지역/테마로 묶을 때 쓰는 "상위 카테고리" 바. 예전엔 이 그룹명을
    배경 없는 굵은 글씨 한 줄로만 표시했는데, 그러면 카테고리라기보다 방문지
    항목 하나처럼 보여서 위계가 안 느껴진다는 피드백이 있었다(멕시코 문명기행
    재테스트) — "카라코람 트레킹 하이라이트", "함께 즐기는 관광 코스" 같은
    섹션 바(add_section_bar)와 같은 배경색 바 형태로 통일하되, 섹션 바보다
    밝은 GROUP_BAR_COLOR를 써서 섹션(진한 색) 아래 놓인 하위 카테고리라는
    걸 색으로도 구분한다. 왼쪽 정렬 + 들여쓰기로 섹션 바(가운데 정렬)와도
    구분한다."""
    if height is None:
        height = group_bar_height(text, size)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, top, CONTENT_W, height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = GROUP_BAR_COLOR
    bar.line.fill.background()
    tf = bar.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    _tf_setup(tf, text, size, WHITE, bold=True, align=PP_ALIGN.LEFT)
    return bar


def add_slide_bg(slide, color=MOBILE_BG_COLOR):
    """슬라이드 전체를 옅은 톤으로 채우는 배경 도형을 맨 뒤로 보내 깔아준다.
    python-pptx는 slide.background로 그라데이션/이미지 배경을 직접 설정하기
    까다로워서(테마 XML 조작 필요), 슬라이드 크기와 같은 사각형 도형을 하나
    그려 맨 뒤로 보내는 방식을 쓴다 — 이후 add_text/add_image_placeholder 등이
    이 위에 그대로 얹힌다."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    spTree = slide.shapes._spTree
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)  # spTree의 맨 앞(0,1은 nvGrpSpPr/grpSpPr)으로 보내 최하단에 깔기
    return bg


def badge_pill_width(text, size=13, bold=True):
    """add_badge_pill이 그릴 알약형 배지의 폭을 텍스트 길이로 대략 추정한다
    (estimate_text_height와 같은 글자폭 근사 방식 + 좌우 여백)."""
    char_w_in = (size / 72) * (1.05 if bold else 0.95)
    return Inches(len(str(text)) * char_w_in + 0.5)


def add_badge_pill(slide, left, top, text, size=13, fill=MOBILE_BADGE_COLOR, text_color=WHITE,
                    width=None):
    """알약형(rounded) 배지 — 모바일 가이드라인의 "핵심 셀링포인트는 배지·인디케이터·
    커스텀 아이콘으로 컴포넌트화" 요구를 반영한 헬퍼. 코스 섹션 위 짧은 태그
    (예: "당일 산행으로 부담없이!")나 표지의 region_tag처럼, 평문이 아니라 눈에 띄는
    태그로 보여주고 싶은 짧은 문구에 쓴다."""
    if width is None:
        width = badge_pill_width(text, size)
    height = Inches(size / 72 * 1.9 + 0.14)
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    pill.fill.solid()
    pill.fill.fore_color.rgb = fill
    pill.line.fill.background()
    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    _tf_setup(tf, text, size, text_color, bold=True, align=PP_ALIGN.CENTER)
    return pill, width, height


def add_image_placeholder(slide, left, top, width, height, label="이미지", label_bottom=False):
    """label_bottom=True면 라벨을 박스 가운데가 아니라 위쪽 가운데에 둔다 — 배너 슬라이드처럼
    박스 위에 카피 텍스트가 겹쳐 올라가는 경우, 가운데 라벨이 카피와 겹쳐 읽기 어려웠다
    (2026-10-27 비교 테스트 배너 4칸 모두에서 확인)."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.background()
    box.line.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    tf = box.text_frame
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM if label_bottom else MSO_ANCHOR.MIDDLE
    _tf_setup(tf, label, 11 if not label_bottom else 9, MUTED_COLOR, align=PP_ALIGN.CENTER)
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


class MobileSlideFlow(SlideFlow):
    """SlideFlow와 동일하지만, 새 슬라이드를 시작할 때마다 옅은 브랜드톤 배경을
    자동으로 깔아준다(가이드라인: "플랫한 흰 배경 지양, 브랜드 톤 배경/텍스처/이미지
    중첩 활용"). 배경만 다르고 페이지 흐름 로직(ensure 등)은 SlideFlow와 완전히
    같아서, 기존 build_* 함수들을 그대로 재사용해도(예: build_route_compare_slide)
    이 배경 위에 자연스럽게 얹힌다."""

    def new_slide(self):
        slide = super().new_slide()
        add_slide_bg(slide)
        return slide


# ---------------------------------------------------------------------------
# 슬라이드 빌더 함수 (정현지 스타일) — 모두 SlideFlow를 받아 가능하면 같은
# 슬라이드에 이어 그리고, 공간이 없을 때만 새 슬라이드를 시작한다.
# ---------------------------------------------------------------------------

def cover_option_text(cover, options_key, single_key):
    """표지 후보안(tagline_options/subtitle_options) 렌더링용 텍스트를 만든다.

    2026-09-15 추가 — 담당자 요청: "신규 상품소개 제작에서 타이틀의 태그라인 등은
    기획안 안에서도 몇 가지 안을 만들어서 보여줬으면 좋겠다(3가지 안)". AI가
    tagline_options/subtitle_options(각 3개짜리 배열)를 채워주면 "안1) ...\\n안2)
    ...\\n안3) ..." 형태로 표지에 나란히 보여주고, 그 필드가 없으면(예: 모바일
    최적화 모드처럼 새 카피를 만들지 않고 원본을 그대로 옮기기만 하는 경우) 예전
    그대로 단일 값(tagline/subtitle)을 보여준다 — 완전히 하위 호환."""
    options = cover.get(options_key)
    if isinstance(options, list):
        options = [str(o).strip() for o in options if str(o).strip()]
    else:
        options = []
    if len(options) >= 2:
        return "\n".join(f"안{i + 1}) {opt}" for i, opt in enumerate(options))
    return cover.get(single_key, "")


def build_cover_slide(flow, cover, watermark_label=""):
    """표지 슬라이드. tagline/product_name/subtitle 모두 고정 높이를 확보해두고
    있었는데, 이 셋 다 AI가 채우는 가변 길이 텍스트라 길어지면 예상보다 줄바꿈이
    늘어 바로 아래 요소와 겹칠 수 있다(background_story에서 실제로 발생한 것과
    같은 문제 — 표지는 모든 상품에서 항상 렌더링되는 슬라이드라 잠재 위험이 가장
    크다). estimate_text_height로 실제 줄 수를 추정해 안전하게 확보한다."""
    slide = flow.new_slide()
    y = flow.y
    tagline = cover_option_text(cover, "tagline_options", "tagline")
    tagline_h = estimate_text_height(tagline, 13, CONTENT_W) if tagline else Inches(0.5)
    add_text(slide, MARGIN, y, CONTENT_W, tagline_h, tagline,
              size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
    y += tagline_h + Inches(0.05)

    product_name = cover.get("product_name", "")
    product_h = estimate_text_height(product_name, 26, CONTENT_W, bold=True) if product_name else Inches(0.9)
    add_text(slide, MARGIN, y, CONTENT_W, product_h, product_name,
              size=26, bold=True, align=PP_ALIGN.CENTER)
    y += product_h + Inches(0.05)

    if cover.get("region_tag"):
        region_h = estimate_text_height(cover["region_tag"], 13, CONTENT_W)
        add_text(slide, MARGIN, y, CONTENT_W, region_h, cover["region_tag"],
                  size=13, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += region_h + Inches(0.05)
    y += Inches(0.1)
    add_small_image_placeholder(slide, y, Inches(2.3), Inches(0.9), "메인 이미지")
    y += Inches(1.0)
    subtitle = cover_option_text(cover, "subtitle_options", "subtitle")
    if subtitle:
        subtitle_h = estimate_text_height(subtitle, 14, CONTENT_W, bold=True)
        add_text(slide, MARGIN, y, CONTENT_W, subtitle_h, subtitle,
                  size=14, bold=True, align=PP_ALIGN.CENTER)
        y += subtitle_h + Inches(0.05)
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


def build_mobile_cover_slide(flow, cover, watermark_label=""):
    """모바일 버전 표지. build_cover_slide와 텍스트 구조·내용은 완전히 동일하고,
    폰트를 키우고 여백을 넓히며(가이드라인 32pt 본문/80~100px 여백 취지), region_tag를
    평문 대신 배지(add_badge_pill)로 컴포넌트화한 점만 다르다 — 실제 돌로미테 샘플도
    표지 문구 자체는 PC/모바일이 동일했다."""
    slide = flow.new_slide()
    y = flow.y
    tagline = cover_option_text(cover, "tagline_options", "tagline")
    tagline_h = estimate_text_height(tagline, MOBILE_CAPTION_SIZE, CONTENT_W,
                                      line_spacing=MOBILE_LINE_SPACING_EST) if tagline else Inches(0.5)
    add_text(slide, MARGIN, y, CONTENT_W, tagline_h, tagline,
              size=MOBILE_CAPTION_SIZE, color=MUTED_COLOR, align=PP_ALIGN.CENTER,
              line_spacing=MOBILE_LINE_SPACING)
    y += tagline_h + Inches(0.08)

    product_name = cover.get("product_name", "")
    title_size = MOBILE_TITLE_SIZE + 8
    product_h = estimate_text_height(product_name, title_size, CONTENT_W, bold=True) if product_name else Inches(1.0)
    add_text(slide, MARGIN, y, CONTENT_W, product_h, product_name,
              size=title_size, bold=True, align=PP_ALIGN.CENTER, color=ACCENT_COLOR)
    y += product_h + Inches(0.12)

    if cover.get("region_tag"):
        pill_w = badge_pill_width(cover["region_tag"], size=MOBILE_CAPTION_SIZE)
        pill_left = MARGIN + (CONTENT_W - pill_w) / 2
        _, _, pill_h = add_badge_pill(
            slide, pill_left, y, cover["region_tag"], size=MOBILE_CAPTION_SIZE, width=pill_w)
        y += pill_h + Inches(0.14)
    y += Inches(0.15)
    add_small_image_placeholder(slide, y, Inches(2.7), Inches(1.1), "메인 이미지 (기존 이미지 재사용)")
    y += Inches(1.2)
    subtitle = cover_option_text(cover, "subtitle_options", "subtitle")
    if subtitle:
        subtitle_h = estimate_text_height(subtitle, MOBILE_TITLE_SIZE, CONTENT_W, bold=True)
        add_text(slide, MARGIN, y, CONTENT_W, subtitle_h, subtitle,
                  size=MOBILE_TITLE_SIZE, bold=True, align=PP_ALIGN.CENTER)
        y += subtitle_h + Inches(0.08)
    intro_h = estimate_text_height(cover.get("intro_copy", ""), MOBILE_BODY_SIZE, CONTENT_W,
                                    line_spacing=MOBILE_LINE_SPACING_EST)
    add_text(slide, MARGIN, y, CONTENT_W, intro_h, cover.get("intro_copy", ""),
              size=MOBILE_BODY_SIZE, color=MUTED_COLOR, align=PP_ALIGN.CENTER,
              line_spacing=MOBILE_LINE_SPACING)
    y += intro_h
    if watermark_label:
        add_text(slide, SLIDE_W - Inches(1.6), Inches(0.15), Inches(1.2), Inches(0.3),
                  watermark_label, size=11, bold=True, color=RGBColor(0xCC, 0xB0, 0x00),
                  align=PP_ALIGN.RIGHT)
    flow.y = y
    return slide


def build_destination_slides(flow, destinations, section_title=None, theme_line=None):
    """목적지 개수만큼만 슬라이드를 만든다. 고정 개수로 나누지 않고, 실제 텍스트
    길이를 추정해서 한 슬라이드에 들어갈 수 있는 만큼만 채우고 넘치면 다음 슬라이드로.
    첫 슬라이드는 이전 섹션이 남긴 여백에 이어 붙일 수 있으면 이어 붙인다.

    region_tag가 있는 방문지는 같은 그룹(성격이 비슷하거나 같은 지역)끼리 묶어서
    그룹명을 소제목으로 한 번만 보여준다 — 예전엔 방문지마다 작은 라벨을 하나씩
    달았는데, 그 라벨이 흰 글씨(WHITE)를 배경 없이 그냥 텍스트로 찍어서 화면에서
    보이지도 않았고(멕시코 문명기행 테스트로 확인), 설령 보였어도 방문지 하나하나에
    태그만 붙을 뿐 "묶어서 보여주는" 느낌은 아니었다. 사업부 자료에 이미 있는 그룹
    구조(예: "[ 사포테카 문명 | 와하카 ]")를 살려 지역/테마별로 소제목 아래 방문지를
    들여쓰기해서 묶어 보여주도록 다시 짰다 — region_tag가 없는 방문지는 예전처럼
    그룹 소제목 없이 그냥 나열된다(하위 호환).

    그 다음 라운드에서 그룹명을 작은 굵은 글씨(add_text)로만 표시했더니, 이번엔
    글씨는 잘 보이지만 "상위 카테고리"라는 느낌이 안 나고 방문지 제목과 비슷한
    무게로 읽힌다는 피드백을 받았다(멕시코 재확인 — "그룹핑한 방문지의 상위
    카테고리가 있어야 해", "카라코람 트레킹 하이라이트"/"함께 즐기는 관광 코스"
    같은 섹션 바 스타일 참고 요청). 그래서 그룹명도 section_title과 같은
    "색 배경 바" 형태(add_group_bar)로 승격했다 — 다만 section_title(섹션 전체를
    가르는 최상위 구분, 예: 트레킹/관광 코스)보다는 한 단계 아래라는 걸 보여주기
    위해 더 밝은 색(GROUP_BAR_COLOR)과 왼쪽 정렬을 쓴다. 즉 위계는
    section_title(진한 색, 가운데 정렬) > region_tag 그룹 바(밝은 색, 좌측 정렬)
    > 개별 방문지 순."""
    if not destinations:
        return []
    slides = []
    idx = 0
    first_slide = True
    header_h = (Inches(0.55) if section_title else Inches(0)) + \
               (Inches(0.4) if theme_line else Inches(0))
    current_group = None  # 화면에 마지막으로 그린 그룹명 — 슬라이드가 넘어가도 유지

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
            # 그룹이 슬라이드 경계를 넘어 이어지면, 새 슬라이드 맨 위에서도 그룹명을
            # 다시 보여준다 — 안 그러면 방문지가 갑자기 그룹 없이 나온 것처럼 보인다.
            # 다음 방문지가 같은 그룹일 때만 그룹 바를 다시 그린다 — 안 그러면 방문지 없이
            # 앞 그룹 바만 덩그러니 남는다(일차별로 region_tag가 1개씩일 때 발견, 2026-10-27).
            if current_group and (destinations[idx].get("region_tag") or None) == current_group:
                bar_h = group_bar_height(current_group)
                add_group_bar(slide, y, current_group, height=bar_h)
                y += bar_h + Inches(0.15)

        indent = Inches(0.15)  # 그룹에 속한 방문지는 소제목 아래 있다는 느낌을 주기 위해 살짝 들여씀
        placed_any = False
        while idx < len(destinations):
            dest = destinations[idx]
            group = dest.get("region_tag") or None
            is_new_group = bool(group) and group != current_group
            title_w = CONTENT_W - (indent if group else Inches(0))
            title_h = estimate_text_height(dest.get("title", ""), 15, title_w, bold=True)
            image_h = Inches(0.45)
            desc_h = estimate_text_height(dest.get("description", ""), 12, title_w)
            # stats: 코스별 거리/소요시간/최고고도/상행·하행 한 줄 요약(없으면 생략).
            # 2026-10-27 — 사람이 만든 기획안엔 코스마다 이 줄이 있는데 자동 생성본엔 없던 문제.
            stats_text = (dest.get("stats") or "").strip()
            stats_h = estimate_text_height(stats_text, 11, title_w, bold=True) if stats_text else Inches(0)
            group_label_h = Inches(0)
            if is_new_group:
                group_label_h = group_bar_height(group) + Inches(0.15)
            block_h = group_label_h + title_h + Inches(0.06) + image_h + Inches(0.1) \
                + (stats_h + Inches(0.06) if stats_text else Inches(0)) \
                + desc_h + Inches(0.18)

            if placed_any and y + block_h > flow.bottom_limit:
                break  # 이 슬라이드엔 더 안 들어감 -> 다음 슬라이드로

            if is_new_group:
                bar_h = group_label_h - Inches(0.15)
                add_group_bar(slide, y, group, height=bar_h)
                y += group_label_h
                current_group = group

            x = MARGIN + (indent if group else Inches(0))
            add_text(slide, x, y, title_w, title_h, dest.get("title", ""),
                      size=15, bold=True)
            y += title_h + Inches(0.06)
            add_small_image_placeholder(slide, y, Inches(1.6), image_h, "이미지")
            y += image_h + Inches(0.1)
            if stats_text:
                add_text(slide, x, y, title_w, stats_h, stats_text, size=11, bold=True,
                          color=ACCENT_COLOR)
                y += stats_h + Inches(0.06)
            add_text(slide, x, y, title_w, desc_h, dest.get("description", ""),
                      size=12, color=MUTED_COLOR)
            y += desc_h + Inches(0.18)

            placed_any = True
            idx += 1

        flow.y = y
        slides.append(slide)
    return slides


def build_mobile_destination_slides(flow, destinations, section_title=None, theme_line=None):
    """build_destination_slides의 모바일 버전. 방문지 제목/설명 내용과 그룹핑 로직은
    완전히 동일하게 재사용하고(원문 그대로 재배치이지 요약이 아니므로), 폰트만
    MOBILE_* 크기로 키우고 줄간격을 넓힌다. section_title은 배지형이 아니라
    기존과 같은 색 바(add_section_bar)를 유지한다 — 실제 샘플에서도 "돌로미테에서
    놓쳐선 안되는 트레킹 코스 Best 6" 같은 대제목은 색 바 스타일 그대로였고,
    그 위에 짧은 보조 태그("당일 산행으로 부담없이!")만 배지로 추가돼 있었다."""
    if not destinations:
        return []
    slides = []
    idx = 0
    first_slide = True
    header_h = (Inches(0.55) if section_title else Inches(0)) + \
               (Inches(0.4) if theme_line else Inches(0))
    current_group = None

    while idx < len(destinations):
        if first_slide:
            y = flow.ensure(header_h + Inches(0.6))
            slide = flow.slide
            if section_title:
                add_section_bar(slide, y, section_title, size=MOBILE_TITLE_SIZE - 2)
                y += Inches(0.55)
            if theme_line:
                add_text(slide, MARGIN, y, CONTENT_W, Inches(0.35), theme_line,
                          size=MOBILE_TITLE_SIZE - 3, bold=True, align=PP_ALIGN.CENTER)
                y += Inches(0.4)
            first_slide = False
        else:
            slide = flow.new_slide()
            y = flow.y
            # 다음 방문지가 같은 그룹일 때만 그룹 바를 다시 그린다 — 안 그러면 방문지 없이
            # 앞 그룹 바만 덩그러니 남는다(일차별로 region_tag가 1개씩일 때 발견, 2026-10-27).
            if current_group and (destinations[idx].get("region_tag") or None) == current_group:
                bar_h = group_bar_height(current_group, size=MOBILE_CAPTION_SIZE + 1)
                add_group_bar(slide, y, current_group, height=bar_h, size=MOBILE_CAPTION_SIZE + 1)
                y += bar_h + Inches(0.15)

        indent = Inches(0.15)
        placed_any = False
        while idx < len(destinations):
            dest = destinations[idx]
            group = dest.get("region_tag") or None
            is_new_group = bool(group) and group != current_group
            title_w = CONTENT_W - (indent if group else Inches(0))
            title_size = MOBILE_TITLE_SIZE - 1
            title_h = estimate_text_height(dest.get("title", ""), title_size, title_w, bold=True)
            image_h = Inches(0.55)
            desc_h = estimate_text_height(dest.get("description", ""), MOBILE_BODY_SIZE, title_w,
                                           line_spacing=MOBILE_LINE_SPACING_EST)
            group_label_h = Inches(0)
            if is_new_group:
                group_label_h = group_bar_height(group, size=MOBILE_CAPTION_SIZE + 1) + Inches(0.15)
            block_h = group_label_h + title_h + Inches(0.08) + image_h + Inches(0.12) \
                + desc_h + Inches(0.22)

            if placed_any and y + block_h > flow.bottom_limit:
                break

            if is_new_group:
                bar_h = group_label_h - Inches(0.15)
                add_group_bar(slide, y, group, height=bar_h, size=MOBILE_CAPTION_SIZE + 1)
                y += group_label_h
                current_group = group

            x = MARGIN + (indent if group else Inches(0))
            add_text(slide, x, y, title_w, title_h, dest.get("title", ""),
                      size=title_size, bold=True, color=ACCENT_COLOR)
            y += title_h + Inches(0.08)
            add_small_image_placeholder(slide, y, Inches(1.9), image_h, "이미지 (기존 이미지 재사용)")
            y += image_h + Inches(0.12)
            add_text(slide, x, y, title_w, desc_h, dest.get("description", ""),
                      size=MOBILE_BODY_SIZE, color=MUTED_COLOR, line_spacing=MOBILE_LINE_SPACING)
            y += desc_h + Inches(0.22)

            placed_any = True
            idx += 1

        flow.y = y
        slides.append(slide)
    return slides


def build_background_slide(flow, background_story):
    """배경 이야기 섹션. 예전엔 "OOO란?" 형태의 kicker 소제목을 title 위에 따로
    붙였는데, 모든 상품 기획안에 기계적으로 반복되는 상투적 표현이라 제거함
    (title만으로 바로 임팩트 있게 시작 — prompt_builder.py도 함께 수정됨).

    title은 20pt 굵은 글씨라 길면 2줄로 줄바꿈되는데, 예전엔 고정 Inches(0.55)만
    확보해두고 그 아래 content를 바로 이어 그려서, title이 2줄이 되는 순간 content
    첫 줄과 겹치는 버그가 있었다(카라코람 가을 버전 테스트에서 확인됨 — "실크로드의
    마지막 관문, 카라코람이 품은 가을의 황금빛"이 2줄로 줄바꿈되며 바로 아래 문단과
    겹쳤음). build_experience_slide/build_safety_slide와 같은 문제라 같은 방식
    (estimate_text_height로 실제 높이 추정)으로 고친다."""
    if not background_story:
        return None
    title = background_story.get("title", "")
    content = background_story.get("content", "")
    title_h = estimate_text_height(title, 20, CONTENT_W, bold=True) if title else Inches(0)
    content_h = estimate_text_height(content, 12, CONTENT_W)
    total_h = (title_h + Inches(0.1) if title else Inches(0)) + content_h

    y = flow.ensure(total_h)
    slide = flow.slide
    if title:
        add_text(slide, MARGIN, y, CONTENT_W, title_h, title,
                  size=20, bold=True, align=PP_ALIGN.CENTER)
        y += title_h + Inches(0.1)
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
    question_h = estimate_text_height(meal_info["question"], 15, CONTENT_W, bold=True)
    ans_h = estimate_text_height(meal_info.get("answer", ""), 12, CONTENT_W)
    total_h = question_h + Inches(0.1) + ans_h

    y = flow.ensure(total_h)
    slide = flow.slide
    add_text(slide, MARGIN, y, CONTENT_W, question_h, meal_info["question"], size=15,
              bold=True, align=PP_ALIGN.CENTER)
    y += question_h + Inches(0.1)
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
    title_h = Inches(0)
    desc_h = Inches(0)
    if experience_points:
        title_h = max(
            estimate_text_height(ep.get("title", ""), 12, col_w - Inches(0.1), bold=True)
            for ep in experience_points
        )
        desc_h = max(
            estimate_text_height(ep.get("description", ""), 10, col_w - Inches(0.1))
            for ep in experience_points
        )
    # 카드 제목이 1줄일 때만 맞는 고정 간격(예전 Inches(0.3))을 쓰면, 제목이 2줄로
    # 줄바꿈되는 순간 바로 아래 설명 텍스트와 겹치는 버그가 있었음 — title_h를 실제
    # 추정 높이로 계산해서 다음 요소를 그만큼 아래로 밀어내도록 수정.
    total_h = (tagline_h + Inches(0.2) if brand_tagline else Inches(0)) \
        + (Inches(0.65) + title_h + Inches(0.1) + desc_h if experience_points else Inches(0))

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
            add_text(slide, x, y, col_w - Inches(0.1), title_h, ep.get("title", ""), size=12,
                      bold=True, align=PP_ALIGN.CENTER)
        y += title_h + Inches(0.1)
        for i, ep in enumerate(experience_points):
            x = MARGIN + col_w * i
            dh = estimate_text_height(ep.get("description", ""), 10, col_w - Inches(0.1))
            add_text(slide, x, y, col_w - Inches(0.1), dh, ep.get("description", ""), size=10,
                      color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += desc_h
    flow.y = y
    return slide


def build_mobile_experience_slide(flow, brand_tagline, experience_points):
    """모바일 버전 "특별함" 카드 섹션. PC 버전(build_experience_slide)은 항목 수만큼
    한 줄에 나란히 배치하는데(예: 5개면 5열), 실제 돌로미테 샘플을 보면 모바일에서는
    한 줄에 최대 3개까지만 두고(5개면 3+2), 아이콘·카드를 훨씬 크게 키우며, 흰
    배경(MOBILE_BG_COLOR로 깔린 옅은 톤 위) 카드로 구획을 나눠 "컴포넌트화"한다.
    텍스트 내용(title/description)은 PC와 동일 — 배치/크기만 바뀐다."""
    if not brand_tagline and not experience_points:
        return None
    tagline_h = estimate_text_height(brand_tagline, MOBILE_TITLE_SIZE, CONTENT_W, bold=True) \
        if brand_tagline else Inches(0)

    MAX_COLS = 3
    points = experience_points or []
    rows = [points[i:i + MAX_COLS] for i in range(0, len(points), MAX_COLS)] if points else []

    gap = Inches(0.12)
    icon_h = Inches(0.75)

    def _row_height(row):
        col_w = CONTENT_W / len(row) - gap
        text_w = col_w - Inches(0.2)  # add_text가 실제로 쓰는 폭(카드 좌우 0.1in 여백 뺀 값)과
                                        # 반드시 같아야 함 — 안 그러면 추정 높이가 실제보다 좁은
                                        # 폭 기준 줄바꿈보다 작게 나와 텍스트가 카드 밖으로 넘친다.
        title_h = max(estimate_text_height(ep.get("title", ""), MOBILE_CAPTION_SIZE + 1, text_w,
                                            bold=True) for ep in row)
        desc_h = max(estimate_text_height(ep.get("description", ""), MOBILE_CAPTION_SIZE, text_w,
                                           line_spacing=MOBILE_LINE_SPACING_EST) for ep in row)
        return Inches(0.15) + icon_h + Inches(0.1) + title_h + Inches(0.06) + desc_h + Inches(0.15), title_h, desc_h

    row_metrics = [_row_height(row) for row in rows]
    total_h = (tagline_h + Inches(0.2) if brand_tagline else Inches(0)) \
        + sum(rh for rh, _, _ in row_metrics) + Inches(0.1) * max(len(rows) - 1, 0)

    y = flow.ensure(total_h)
    slide = flow.slide
    if brand_tagline:
        add_text(slide, MARGIN, y, CONTENT_W, tagline_h, brand_tagline, size=MOBILE_TITLE_SIZE,
                  bold=True, align=PP_ALIGN.CENTER, color=ACCENT_COLOR)
        y += tagline_h + Inches(0.2)

    for row, (row_h, title_h, desc_h) in zip(rows, row_metrics):
        col_w = CONTENT_W / len(row) - gap
        for i, ep in enumerate(row):
            x = MARGIN + (col_w + gap) * i
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, row_h)
            card.fill.solid()
            card.fill.fore_color.rgb = MOBILE_CARD_COLOR
            card.line.color.rgb = RGBColor(0xDD, 0xE6, 0xEA)
            card.shadow.inherit = False
            card.text_frame.word_wrap = True
            card.text_frame.margin_top = Inches(0.1)
        for i, ep in enumerate(row):
            x = MARGIN + (col_w + gap) * i
            add_image_placeholder(slide, x + Inches(0.1), y + Inches(0.15), col_w - Inches(0.2),
                                   icon_h, "아이콘")
        cy = y + Inches(0.15) + icon_h + Inches(0.1)
        for i, ep in enumerate(row):
            x = MARGIN + (col_w + gap) * i
            add_text(slide, x + Inches(0.1), cy, col_w - Inches(0.2), title_h, ep.get("title", ""),
                      size=MOBILE_CAPTION_SIZE + 1, bold=True, align=PP_ALIGN.CENTER)
        cy += title_h + Inches(0.06)
        for i, ep in enumerate(row):
            x = MARGIN + (col_w + gap) * i
            add_text(slide, x + Inches(0.1), cy, col_w - Inches(0.2), desc_h, ep.get("description", ""),
                      size=MOBILE_CAPTION_SIZE, color=MUTED_COLOR, align=PP_ALIGN.CENTER,
                      line_spacing=MOBILE_LINE_SPACING)
        y += row_h + Inches(0.1)

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
    # 기온/강수량 값 표 — 행 수(최저/최고 + 강수량이 있으면 1행) × 행 높이 + 헤더 1행
    has_rain = bool(season_table) and any((r.get("rain") or "").strip() for r in season_table)
    n_rows = (3 if has_rain else 2) + 1
    table_h = (Inches(0.85) + Inches(0.26) * n_rows + Inches(0.1)) if season_table else Inches(0)
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
        # 예전엔 월 이름만 한 줄로 찍고 기온/강수량 값은 JSON에만 있어서 디자이너가 값을
        # 볼 수 없었다(2026-10-27 비교 테스트에서 확인) — 값이 보이는 실제 표를 넣는다.
        row_h = Inches(0.26)
        tbl_shape = slide.shapes.add_table(n_rows, len(season_table) + 1, MARGIN, y,
                                           CONTENT_W, row_h * n_rows)
        tbl = tbl_shape.table
        labels = ["구분", "최저기온 (°C)", "최고기온 (°C)"] + (["강수량 (mm)"] if has_rain else [])
        for r_i, label in enumerate(labels):
            tbl.cell(r_i, 0).text = label
        for c_i, row in enumerate(season_table, start=1):
            tbl.cell(0, c_i).text = str(row.get("month", ""))
            tbl.cell(1, c_i).text = str(row.get("low", ""))
            tbl.cell(2, c_i).text = str(row.get("high", ""))
            if has_rain:
                tbl.cell(3, c_i).text = str(row.get("rain", ""))
        for r_i in range(n_rows):
            tbl.rows[r_i].height = row_h
            for c_i in range(len(season_table) + 1):
                cell = tbl.cell(r_i, c_i)
                cell.margin_top = cell.margin_bottom = Inches(0.02)
                for para in cell.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER
                    for run in para.runs:
                        run.font.size = Pt(9)
                        run.font.name = FONT_NAME
                        run.font.bold = (r_i == 0 or c_i == 0)
        y += row_h * n_rows + Inches(0.1)
    flow.y = y
    return slide


# 구간별 고도 프로필(자리표시 박스 + 코스명/고도 라벨) 슬라이드 부분 노출 여부.
# 2026-10-27 담당자(정현지) 요청: 항상 들어가는데 필요 없는 부분이라 기본값을 끔(False).
# 다시 켜려면 True로 바꾸고 prompt_builder.py 스키마에 altitude_profile을 되살리면 된다.
SHOW_ALTITUDE_PROFILE = False


def build_safety_slide(flow, altitude_profile, safety_note):
    """경유지 고도 프로필(있는 경우) + 안전/난이도 관련 표준 안내.
    고산 트레킹의 '고산증', 도보순례의 '체력/보험', 일반 하이킹의 '난이도' 등
    카테고리에 따라 톤이 다른 표준 안내문을 담는 범용 섹션."""
    if not SHOW_ALTITUDE_PROFILE:
        altitude_profile = None
    if not altitude_profile and not safety_note:
        return None
    ans_h = Inches(0)
    question_h = Inches(0)
    qa_h = Inches(0)
    if safety_note and safety_note.get("question"):
        question_h = estimate_text_height(safety_note["question"], 15, CONTENT_W, bold=True)
        ans_h = estimate_text_height(safety_note.get("answer", ""), 12, CONTENT_W)
        qa_h = question_h + Inches(0.1) + ans_h + Inches(0.35)

    n = len(altitude_profile) if altitude_profile else 0
    col_w = CONTENT_W / max(n, 1)
    gap = Inches(0.06)
    labels = []
    if altitude_profile:
        for stop in altitude_profile:
            # 자리표시 박스는 "숙박"이 아니라 "지도"여야 함 — 예전 코드가 숙박/롯지
            # 소개 슬라이드에서 라벨만 안 바꾼 채 복붙된 흔적으로 보임.
            label = f"{stop.get('name','')}\n{stop.get('altitude','')}"
            extra = " / ".join(v for v in (stop.get("distance"), stop.get("duration")) if v)
            if extra:
                label += f"\n{extra}"
            if stop.get("highlight"):
                label += f"\n{stop['highlight']}"
            labels.append(label)
    # 라벨 줄 수가 highlight 유무에 따라 2~4줄로 달라지므로, 고정 Inches(0.65) 대신
    # 실제 텍스트 높이를 추정해서 다음 요소와 겹치지 않게 함(build_experience_slide와
    # 동일한 문제였음).
    label_h = max((estimate_text_height(l, 9, col_w - gap * 2) for l in labels), default=Inches(0.5))
    profile_h = (Inches(0.4) + Inches(0.6) + label_h) if altitude_profile else Inches(0)
    total_h = qa_h + profile_h

    y = flow.ensure(total_h)
    slide = flow.slide
    if safety_note and safety_note.get("question"):
        add_text(slide, MARGIN, y, CONTENT_W, question_h, safety_note["question"], size=15,
                  bold=True, align=PP_ALIGN.CENTER)
        y += question_h + Inches(0.1)
        add_text(slide, MARGIN, y, CONTENT_W, ans_h, safety_note.get("answer", ""), size=12,
                  align=PP_ALIGN.CENTER)
        y += ans_h + Inches(0.35)
    if altitude_profile:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.3),
                  "구간별 고도 프로필 (자리표시 — 실제 그래픽은 디자이너 작업)",
                  size=10, color=MUTED_COLOR, align=PP_ALIGN.CENTER)
        y += Inches(0.4)
        for i in range(n):
            x = MARGIN + col_w * i
            add_image_placeholder(slide, x + gap, y, col_w - gap * 2, Inches(0.5), "지도")
        y += Inches(0.6)
        for i, label in enumerate(labels):
            x = MARGIN + col_w * i
            add_text(slide, x + gap, y, col_w - gap * 2, label_h, label, size=9, align=PP_ALIGN.CENTER)
        y += label_h
    flow.y = y
    return slide


def build_mobile_qna_slide(flow, qna_pairs, heading="궁금증을 풀어드려요"):
    """meal_info/safety_note(각각 question/answer 한 쌍)를 모바일 버전에서는
    돌로미테 샘플의 "돌로미테 트레킹에 대한 궁금증을 풀어드려요!" 섹션처럼 아이콘
    카드형 Q&A 리스트로 묶어서 보여준다. PC 버전은 이 둘을 각각 별도 섹션(question
    문장을 큰 소제목처럼)으로 보여주는데, 실제 샘플은 여러 Q&A를 카드 리스트 하나로
    묶어 보여줘서 그 구조를 따른다 — question/answer 텍스트 내용 자체는 그대로 둔다."""
    qna_pairs = [qa for qa in (qna_pairs or []) if qa and qa.get("question")]
    if not qna_pairs:
        return None
    header_h = Inches(0.55)
    tw = CONTENT_W - Inches(1.0)

    def _card_height(qa):
        q_h = estimate_text_height(qa["question"], MOBILE_CAPTION_SIZE + 2, tw, bold=True)
        a_h = estimate_text_height(qa.get("answer", ""), MOBILE_BODY_SIZE, tw,
                                    line_spacing=MOBILE_LINE_SPACING_EST)
        return Inches(0.2) + q_h + Inches(0.06) + a_h + Inches(0.2), q_h, a_h

    card_metrics = [_card_height(qa) for qa in qna_pairs]
    # 헤더만 그려놓고 카드 첫 개도 못 들어갈 자리에서 끝나면(예전 버그) 소제목이
    # 다음 슬라이드의 카드들과 떨어져 혼자 남는다 — 헤더 + 첫 카드 높이를 함께
    # 확보해야 헤더가 반드시 카드 하나 이상과 같은 슬라이드에 붙는다.
    y = flow.ensure(header_h + Inches(0.1) + card_metrics[0][0])
    slide = flow.slide
    add_section_bar(slide, y, heading, size=MOBILE_TITLE_SIZE - 2)
    y += header_h + Inches(0.1)

    for qa, (card_h, q_h, a_h) in zip(qna_pairs, card_metrics):
        if y + card_h > flow.bottom_limit:
            slide = flow.new_slide()
            y = flow.y
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, MARGIN, y, CONTENT_W, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = MOBILE_CARD_COLOR
        card.line.color.rgb = RGBColor(0xDD, 0xE6, 0xEA)
        card.shadow.inherit = False
        add_image_placeholder(slide, MARGIN + Inches(0.15), y + Inches(0.2), Inches(0.55),
                               Inches(0.55), "아이콘")
        tx = MARGIN + Inches(0.85)
        add_text(slide, tx, y + Inches(0.2), tw, q_h, qa["question"], size=MOBILE_CAPTION_SIZE + 2,
                  bold=True, color=ACCENT_COLOR)
        add_text(slide, tx, y + Inches(0.2) + q_h + Inches(0.06), tw, a_h, qa.get("answer", ""),
                  size=MOBILE_BODY_SIZE, color=MUTED_COLOR, line_spacing=MOBILE_LINE_SPACING)
        y += card_h + Inches(0.15)

    flow.y = y
    return slide


def build_banner_request_slide(flow, cover, banner_copy=None, bar_color=None):
    """배너 기획 페이지 — 실제 회사 배너제작 템플릿(배너제작.pptx)의 레이아웃을
    그대로 재현. 4개 배너 슬롯(메인 와이드/서브메인 띠/서브메인 2단/지역 리스트)에
    각각 이미지 자리와 '태그라인+타이틀' 텍스트를 넣는다. 스펙 라벨/안내 문구는
    회사 표준이라 고정값이며, 지역/상품과 무관하게 항상 그대로 포함. 절대 위치로
    실제 배너 템플릿과 맞춰야 해서 다른 섹션과 공유하지 않고 항상 전용 슬라이드로 만든다.

    예전엔 cover.tagline+product_name(표지 전체 카피)을 4개 슬롯에 그대로 복붙해서
    문구가 길고 밋밋했음 — 실제 회사 배너는 표지 카피와 다르게, 상품에서 가장
    눈에 띄는 포인트만 뽑아 훨씬 짧고 강렬하게 압축한 별도 카피를 쓴다(첨부 배너
    예시: "봄으로 물든 차마고도를 걷다 / 카피할 수 없는 오리지널의 위엄" +
    "세계 3대 트레킹 / 호도협·옥룡설산"). banner_copy가 있으면 그걸 쓰고, AI가
    누락했을 때만 cover 필드로 대체(하위 호환).

    4개 슬롯 모두 "부제(후킹 카피) + 상품명(타이틀)" 두 요소를 함께 넣는다. 처음엔
    상품명을 부제보다 작게 눌러서 부제(후킹 문구)가 시선을 먼저 받게 했었는데,
    실제 배너에서는 반대로 상품명(타이틀)이 메인 카피이고 부제는 그걸 보조하는
    작은 문구라는 피드백을 받아 위계를 뒤집었다 — 부제는 작고 옅게, 상품명은
    크고 굵게 렌더링한다. 부제와 상품명을 하나의 멀티라인 텍스트박스에 같은
    크기로 합쳐 넣으면(예전 방식) 위계가 전혀 안 생기고, 심지어 메인 와이드
    배너는 상품명 자체가 빠지는 버그가 있었다(카라코람 테스트에서 확인됨) —
    부제와 상품명을 별도 텍스트박스로 쌓아 크기를 다르게 주고, 4개 슬롯 전부에
    상품명을 반드시 포함시킨다."""
    slide = flow.new_slide()
    banner_copy = banner_copy or {}
    kicker = banner_copy.get("kicker") or cover.get("tagline", "")
    title = banner_copy.get("title") or cover.get("product_name", "")
    product_name = cover.get("product_name", "")

    def _add_banner_copy(left, top, width, subtitle_text, subtitle_size, name_size):
        """부제(작고 옅은 글씨)를 먼저 쌓고, 그 아래 상품명/타이틀(더 크고 굵은
        글씨)을 이어 쌓는다 — 실제 들어간 줄 수에 맞춰 높이를 추정해서 다음
        요소랑 안 겹치게 한다."""
        y = top
        if subtitle_text:
            h = estimate_text_height(subtitle_text, subtitle_size, width)
            add_text(slide, left, y, width, h, subtitle_text, size=subtitle_size,
                      color=MUTED_COLOR, align=PP_ALIGN.CENTER)
            y += h
        if product_name:
            h = estimate_text_height(product_name, name_size, width, bold=True)
            add_text(slide, left, y, width, h, product_name, size=name_size,
                      bold=True, align=PP_ALIGN.CENTER)
            y += h
        return y

    if bar_color is None:
        add_section_bar(slide, Inches(0), "배너 기획", height=Inches(0.47), size=14)
    else:  # PC 정돈형: 페이지 폭 전체 네이비 바
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.47))
        bar.fill.solid()
        bar.fill.fore_color.rgb = bar_color
        bar.line.fill.background()
        bar.shadow.inherit = False
        bar.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        _tf_setup(bar.text_frame, "배너 기획", 14, WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.106), Inches(0.675), Inches(1.717), Inches(0.404),
              "배너제작요청", size=13, bold=True)

    # 메인 와이드 배너 — 부제 2줄(kicker+title, 후킹 카피, size 11) + 상품명(size 18, 굵게)
    add_text(slide, Inches(0.106), Inches(1.271), Inches(5.701), Inches(0.303),
              "메인 와이드 배너 (PC: 1920x700, MO: 750x510), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.217), Inches(1.630), Inches(5.475), Inches(1.817), "이미지", label_bottom=True)
    _add_banner_copy(Inches(0.388), Inches(2.158), Inches(5.133),
                      f"{kicker}\n{title}", subtitle_size=11, name_size=18)

    # 서브메인 띠배너 — 부제 1줄(kicker, size 11) + 상품명(size 16, 굵게)
    add_text(slide, Inches(0.081), Inches(3.581), Inches(5.642), Inches(0.303),
              "서브메인 띠배너 (PC: 1920x200, MO: 750x200), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.136), Inches(4.021), Inches(7.028), Inches(0.979), "이미지", label_bottom=True)
    _add_banner_copy(Inches(1.184), Inches(4.173), Inches(5.133),
                      kicker, subtitle_size=11, name_size=16)

    # 서브메인 2단배너 — 부제 1줄(kicker, size 9) + 상품명(size 13, 굵게)
    add_text(slide, Inches(0.136), Inches(5.137), Inches(5.642), Inches(0.303),
              "서브메인 2단배너 (PC: 590x370, MO: 585x670), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.221), Inches(5.521), Inches(2.835), Inches(1.771), "이미지", label_bottom=True)
    _add_banner_copy(Inches(0.288), Inches(6.138), Inches(2.835),
                      kicker, subtitle_size=9, name_size=13)

    # 지역 리스트 배너 — 부제 1줄(kicker, size 11) + 상품명(size 16, 굵게)
    add_text(slide, Inches(0.205), Inches(7.481), Inches(5.701), Inches(0.303),
              "지역 리스트 배너 (PC: 1200x207, MO: 750x207), 가이드라인에 맞춰 제작", size=9, color=MUTED_COLOR)
    add_image_placeholder(slide, Inches(0.316), Inches(7.989), Inches(6.871), Inches(1.336), "이미지", label_bottom=True)
    _add_banner_copy(Inches(1.184), Inches(8.280), Inches(5.133),
                      kicker, subtitle_size=11, name_size=16)

    return slide


def build_review_notes_slide(flow, review):
    """Gemini 교차검수 결과를 PPTX 마지막 페이지에 그대로 남긴다 — 예전엔 Streamlit
    화면에만 표시되고 다운로드한 PPTX에는 안 남아서, 파일만 디자이너/다른 담당자에게
    넘기면 검수 내역이 통째로 사라지는 문제가 있었다. review가 없거나 검수 자체를
    건너뛴 경우(_dry_run — GEMINI_API_KEY 미설정)는 내용 없는 안내 페이지를 억지로
    넣지 않기 위해 슬라이드를 만들지 않는다."""
    if not review or review.get("_dry_run"):
        return []
    issues = review.get("issues") or []

    slide = flow.new_slide()
    y = flow.y
    add_section_bar(slide, y, "Gemini 교차검수 결과 (내부 참고용)")
    y += Inches(0.55)
    slides = [slide]

    if not issues:
        add_text(slide, MARGIN, y, CONTENT_W, Inches(0.4),
                  "✅ 검수 통과 — 왜곡/날조·저작권/표절·사실확인·비문/맞춤법 이슈가 "
                  "발견되지 않았습니다.", size=12, align=PP_ALIGN.CENTER)
        flow.y = y + Inches(0.4)
        return slides

    if review.get("summary"):
        summary_h = estimate_text_height(review["summary"], 11, CONTENT_W)
        add_text(slide, MARGIN, y, CONTENT_W, summary_h, review["summary"], size=11,
                  color=MUTED_COLOR)
        y += summary_h + Inches(0.2)

    for issue in issues:
        header_text = f"[{issue.get('category', '')} · {issue.get('severity', '')}] {issue.get('field', '')}"
        header_h = estimate_text_height(header_text, 12, CONTENT_W, bold=True)
        quote = issue.get("quote", "")
        quote_h = estimate_text_height(quote, 10, CONTENT_W - Inches(0.2)) if quote else Inches(0)
        explanation_h = estimate_text_height(issue.get("explanation", ""), 11, CONTENT_W)
        block_h = header_h + Inches(0.05) \
            + (quote_h + Inches(0.05) if quote else Inches(0)) \
            + explanation_h + Inches(0.22)

        if y + block_h > flow.bottom_limit:
            slide = flow.new_slide()
            y = flow.y
            slides.append(slide)

        add_text(slide, MARGIN, y, CONTENT_W, header_h, header_text, size=12, bold=True,
                  color=ACCENT_COLOR)
        y += header_h + Inches(0.05)
        if quote:
            add_text(slide, MARGIN + Inches(0.1), y, CONTENT_W - Inches(0.2), quote_h,
                      f"“{quote}”", size=10, color=MUTED_COLOR)
            y += quote_h + Inches(0.05)
        add_text(slide, MARGIN, y, CONTENT_W, explanation_h, issue.get("explanation", ""), size=11)
        y += explanation_h + Inches(0.22)

    flow.y = y
    return slides


# ===========================================================================
# PC 정돈형 레이아웃 (2026-11) — build()가 사용한다. (모바일 build_mobile은 위쪽의 기존 함수 그대로)
#
# 담당자 피드백: "제작되는 기획안이 디자이너가 보기도 힘들고 정돈되지 않은 느낌" →
#   (1) 글자 크기·가독성  (2) 이미지 자리 구분(번호)  (3) 제목 스타일 불일치  (4) 여백·균형
#   (5) 포인트 부분은 아이콘 활용
# 사람이 만든 기획안(돌로미테 Part.2, 호쿠리쿠)의 구조를 기준으로 다시 짰다:
#   - 제목 위계 통일: H1 = 모서리를 자른 네이비 바(섹션), H2 = 번호 탭(포인트 항목),
#     그룹 바 = 옅은 파란 바 + 네이비 세로선, 코스 항목 = 번호 원 + 제목
#   - 이미지 자리는 문서 전체에서 "이미지 1, 이미지 2 …" 연번 + 항목당 장수(image_count, 기본 2)
#   - 본문은 왼쪽 정렬 11pt 진한 글씨(예전: 가운데 정렬 12pt 옅은 회색), 문단(\n) 사이 간격
#   - 섹션 통째로 다음 장으로 넘기지 않고 항목 단위로 이어 붙여 장 하단 빈 공간을 줄임
#     (제목만 장 끝에 홀로 남지 않도록 "제목 + 첫 항목"을 함께 확보)
#   - 아이콘: icons_data.py(lucide ISC 아이콘 마스크, 실행 시 네이비/흰색 입힘). AI가 icon 키를
#     주면 쓰고, 없으면 제목·설명 키워드로 추정(pick_icon)해서 기존 JSON에도 아이콘이 붙는다.
# ===========================================================================
import os
import re

from pptx.enum.dml import MSO_LINE

PC_NAVY = RGBColor(0x22, 0x2A, 0x35)       # 사람 기획안 섹션 탭 색(tx2 50%)
PC_LINE_BLUE = RGBColor(0x41, 0x71, 0x9C)   # 사람 기획안 이미지 자리 테두리 색
PC_PALE = RGBColor(0xEE, 0xF3, 0xF8)        # 이미지 자리/칩/그룹 바 배경
PC_PALE_BAR = RGBColor(0xE3, 0xEB, 0xF3)
PC_BODY = RGBColor(0x2B, 0x2B, 0x2B)
PC_SUB = RGBColor(0x5A, 0x5A, 0x5A)
PC_PAGE_MARGIN = Inches(0.45)
PC_CW = SLIDE_W - PC_PAGE_MARGIN * 2
PC_BODY_SIZE = 11
PC_LS = 1.08                                 # 본문 줄간격 배수
PC_LINE_H = 1.38                             # 한 줄 높이 ≈ 글자크기 × 1.38 × 줄간격 (맑은 고딕+대체폰트 여유)
PC_PARA_GAP_PT = 5

ICON_KEYS = [
    "mountain", "mountain-snow", "trees", "flower-2", "leaf", "snowflake", "sun", "cloud-sun", "waves",
    "droplets", "hotel", "bed-double", "utensils", "wine", "coffee", "soup", "train-front", "bus", "plane",
    "ship", "footprints", "map", "map-pin", "compass", "camera", "star", "heart", "shield-check", "users",
    "clock", "calendar-days", "landmark", "castle", "church", "building-2", "thermometer", "ticket",
    "wallet", "sparkles", "gem", "sunrise", "tent", "backpack", "route", "binoculars", "umbrella", "car",
    "sailboat", "flame", "bath", "hand-heart", "badge-check", "circle-check", "image", "circle-help",
    "thumbs-up", "ban", "gift",
]

# (키워드들, 아이콘) — 위에서부터 먼저 맞는 규칙을 쓴다. 제목을 먼저, 없으면 설명을 본다.
_ICON_RULES = [
    (("노쇼핑", "노옵션", "쇼핑", "옵션"), "badge-check"),
    (("인솔", "동행", "가이드", "전문가", "인솔자"), "users"),
    (("온천", "힐링", "스파"), "bath"),
    (("호텔", "숙박", "리조트", "롯지", "휴양", "휴식", "연박"), "hotel"),
    (("와인", "맥주", "양조"), "wine"),
    (("카페", "커피"), "coffee"),
    (("식사", "미식", "음식", "요리", "맛집", "레스토랑", "산해진미", "별미"), "utensils"),
    (("기차", "열차", "철도", "트램", "빙하특급"), "train-front"),
    (("케이블카", "곤돌라", "리프트"), "route"),
    (("버스", "전용차", "차량"), "bus"),
    (("항공", "직항", "비행"), "plane"),
    (("크루즈", "선박", "페리", "유람선"), "ship"),
    (("설산", "빙하", "설경", "만년설", "눈"), "mountain-snow"),
    (("호수", "폭포", "계곡", "강", "해변", "바다"), "waves"),
    (("꽃", "야생화", "단풍", "벚꽃", "정원", "튤립"), "flower-2"),
    (("숲", "초원", "자연", "나무"), "trees"),
    (("일출", "일몰", "노을", "전망", "파노라마", "뷰"), "sunrise"),
    (("사진", "포토", "촬영"), "camera"),
    (("유적", "문화", "역사", "사원", "성당", "교회", "박물관", "유네스코", "세계유산"), "landmark"),
    (("안전", "보험", "안심"), "shield-check"),
    (("산", "암봉", "봉우리", "트레킹", "능선", "고원", "등반", "산행"), "mountain"),
    (("당일", "시간", "일정", "여유", "소요"), "clock"),
    (("지도", "루트", "코스", "노선"), "route"),
    (("가족", "친구", "커플", "함께"), "heart"),
    (("특별", "프리미엄", "단독", "최초", "유일", "완성"), "sparkles"),
    (("텐트", "캠핑", "야영"), "tent"),
    # 개요(핵심 정보 칸)용 보강 규칙 — 앞 규칙에 안 걸린 라벨만 여기로 온다
    (("기간", "박", "출발", "일수"), "calendar-days"),
    (("난이도", "체력", "강도"), "thermometer"),
    (("지역", "도시", "핵심", "방문", "나라", "국가"), "map-pin"),
    (("이동", "교통", "수단"), "bus"),
    (("인원", "규모", "정원"), "users"),
    (("가격", "비용", "요금", "경비"), "wallet"),
    (("체험", "활동", "액티비티"), "sparkles"),
]


def pick_icon(title="", description="", hint=None):
    """AI가 준 icon 키(hint)가 목록에 있으면 그걸, 아니면 제목→설명 순으로 키워드를 훑어 추정."""
    if hint and str(hint).strip() in ICON_KEYS:
        return str(hint).strip()
    for text in (title or "", description or ""):
        for words, key in _ICON_RULES:
            if any(w in text for w in words):
                return key
    return "star"


_ICON_CACHE = {}


def _icon_png(key, white=False):
    """icons_data.ICON_MASKS의 알파 마스크에 색(네이비/흰색)을 입혀 PNG 바이트로 돌려준다(캐시)."""
    ck = (key, white)
    if ck in _ICON_CACHE:
        return _ICON_CACHE[ck]
    data = None
    try:
        import base64
        import io as _io
        from PIL import Image
        from icons_data import ICON_MASKS
        mask_b64 = ICON_MASKS.get(key)
        if mask_b64:
            mask = Image.open(_io.BytesIO(base64.b64decode(mask_b64))).convert("L")
            color = (255, 255, 255, 255) if white else (0x22, 0x2A, 0x35, 255)
            img = Image.new("RGBA", mask.size, color)
            img.putalpha(mask)
            buf = _io.BytesIO()
            img.save(buf, "PNG")
            data = buf.getvalue()
    except Exception:
        data = None  # 아이콘은 장식 요소 — 실패해도 기획안 생성은 계속한다
    _ICON_CACHE[ck] = data
    return data


def add_icon(slide, key, left, top, size, white=False):
    data = _icon_png(key, white)
    if not data:
        return None
    import io as _io
    return slide.shapes.add_picture(_io.BytesIO(data), left, top, size, size)


# ---- 텍스트 측정(글자 종류별 폭 + 단어 단위 줄바꿈 시뮬레이션) -------------------------

def _char_w(ch, bold):
    o = ord(ch)
    if ch == " ":
        w = 0.30
    elif (0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F or 0x4E00 <= o <= 0x9FFF
          or 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFFEF or 0x2460 <= o <= 0x27BF):
        w = 0.97
    elif ch.isdigit():
        w = 0.56
    elif ch.isascii() and ch.isalpha():
        w = 0.62 if ch.isupper() else 0.54
    elif ch in ".,:;'!|":
        w = 0.32
    elif ch in "·()[]-~/":
        w = 0.42
    elif ch in "‘’“”":
        w = 0.5
    else:
        w = 0.7
    return w * (1.04 if bold else 1.0)


def _word_w(word, bold):
    return sum(_char_w(c, bold) for c in word)


def count_lines(text, size_pt, width_emu, bold=False):
    """줄바꿈 후 총 줄 수(문단별 합). 공백 기준으로 단어를 채우고, 한 단어가 한 줄보다
    길면 글자 단위로 끊는다. 약간(4%) 보수적으로 잡는다."""
    width_pt = Emu(width_emu).pt / 1.02
    total = 0
    for para in str(text).split("\n"):
        if not para.strip():
            total += 1
            continue
        lines, cur = 1, 0.0
        for word in re.findall(r"\S+\s*|\s+", para):
            ww = _word_w(word, bold) * size_pt
            ww_trim = _word_w(word.rstrip(), bold) * size_pt
            if cur + ww_trim <= width_pt:
                cur += ww
                continue
            if cur > 0:
                lines += 1
                cur = 0.0
            if ww_trim > width_pt:  # 한 단어가 한 줄보다 길다 -> 글자 단위
                for ch in word:
                    cw = _char_w(ch, bold) * size_pt
                    if cur + cw > width_pt:
                        lines += 1
                        cur = 0.0
                    cur += cw
            else:
                cur = ww
        total += lines
    return total


def text_h(text, size_pt, width_emu, bold=False, ls=PC_LS, para_gap_pt=PC_PARA_GAP_PT):
    if not text:
        return Emu(0)
    n_para = len(str(text).split("\n"))
    lines = count_lines(text, size_pt, width_emu, bold)
    pt = lines * size_pt * PC_LINE_H * ls + (n_para - 1) * para_gap_pt + 2
    return Pt(pt)


def pc_text(slide, left, top, width, text, size=PC_BODY_SIZE, color=PC_BODY, bold=False,
            align=PP_ALIGN.LEFT, ls=PC_LS, para_gap_pt=PC_PARA_GAP_PT, anchor=MSO_ANCHOR.TOP):
    """여백 0인 텍스트 박스를 만들고, 측정한 높이를 돌려준다(박스 높이 = 측정 높이)."""
    h = text_h(text, size, width, bold=bold, ls=ls, para_gap_pt=para_gap_pt)
    box = slide.shapes.add_textbox(left, top, width, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = ls
        if i < len(lines) - 1:
            p.space_after = Pt(para_gap_pt)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = FONT_NAME
    return h


def _rect(slide, shape, left, top, width, height, fill=None, line=None, line_w=0.75, dash=False):
    s = slide.shapes.add_shape(shape, left, top, width, height)
    s.shadow.inherit = False
    for er in s._element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef"):
        er.set("idx", "0")  # 테마 그림자 참조 제거(LibreOffice 등 일부 뷰어가 effectLst를 무시하는 경우 대비)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
        if dash:
            s.line.dash_style = MSO_LINE.DASH
    return s


def _shape_text(shape, text, size, color, bold=False, align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE, margins=(0.08, 0.02, 0.08, 0.02)):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margins]
    _tf_setup(tf, text, size, color, bold=bold, align=align)


# ---- 페이지 흐름 -------------------------------------------------------------------

class PCFlow:
    """항목 단위로 이어 붙이는 페이지 커서. 새 페이지마다 하단 푸터(상품명 · 페이지 번호)를 깐다."""

    def __init__(self, prs, footer_label=""):
        self.prs = prs
        self.footer_label = footer_label
        self.slide = None
        self.page_no = 0
        self.top = Inches(0.45)
        self.y = self.top
        self.bottom_limit = SLIDE_H - Inches(0.62)
        self.img_no = 0

    def new_slide(self, footer=True):
        self.slide = _blank_slide(self.prs)
        self.page_no += 1
        self.y = self.top
        if footer:
            self._footer()
        return self.slide

    new_page = new_slide

    def _footer(self):
        s = self.slide
        y = SLIDE_H - Inches(0.42)
        line = s.shapes.add_connector(1, PC_PAGE_MARGIN, y, SLIDE_W - PC_PAGE_MARGIN, y)
        line.line.color.rgb = RGBColor(0xC9, 0xD3, 0xDE)
        line.line.width = Pt(0.5)
        if self.footer_label:
            pc_text(s, PC_PAGE_MARGIN, y + Inches(0.07), Inches(4.5), self.footer_label, size=8,
                    color=RGBColor(0x8A, 0x94, 0xA0), ls=1.0)
        pc_text(s, SLIDE_W - PC_PAGE_MARGIN - Inches(1.0), y + Inches(0.07), Inches(1.0),
                f"{self.page_no}", size=8, color=RGBColor(0x8A, 0x94, 0xA0), align=PP_ALIGN.RIGHT, ls=1.0)

    def ensure(self, height, gap_before=Inches(0.28)):
        """height만큼 이어 붙일 수 있으면 그 y를, 아니면 새 페이지의 y를 돌려준다.
        페이지 맨 위에서는 gap을 주지 않는다."""
        if self.slide is None:
            self.new_slide()
            return self.y
        at_top = abs(self.y - self.top) < Inches(0.01)
        cand = self.y if at_top else self.y + gap_before
        if cand + height > self.bottom_limit and not at_top:
            self.new_slide()
            return self.y
        self.y = cand
        return self.y

    def take_images(self, n):
        nums = list(range(self.img_no + 1, self.img_no + n + 1))
        self.img_no += n
        return nums


# ---- 공통 부품 ---------------------------------------------------------------------

PC_H1_H = Inches(0.42)


def pc_h1(flow, y, text, icon=None):
    """섹션 제목: 모서리를 자른 네이비 바 + (선택) 흰색 아이콘. 높이 PC_H1_H."""
    s = flow.slide
    bar = _rect(s, MSO_SHAPE.SNIP_1_RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, PC_H1_H, fill=PC_NAVY)
    bar.adjustments[0] = 0.35
    tx = PC_PAGE_MARGIN + Inches(0.18)
    if icon:
        add_icon(s, icon, tx, y + Inches(0.08), Inches(0.26), white=True)
        tx += Inches(0.38)
    pc_text(s, tx, y + Inches(0.07), PC_CW - (tx - PC_PAGE_MARGIN) - Inches(0.4), text, size=13.5,
            color=WHITE, bold=True, ls=1.0, anchor=MSO_ANCHOR.MIDDLE)
    return PC_H1_H


def _tab_height(text):
    size = 12
    one_line_w = Emu(int(_word_w(text, True) * size * 12700 * 1.06)) + Inches(0.95)
    w = min(PC_CW, max(Inches(2.6), one_line_w))
    th = text_h(text, size, w - Inches(0.9), bold=True, ls=1.0, para_gap_pt=0)
    return w, max(Inches(0.36), th + Inches(0.16))


def pc_tab(flow, y, number_label, text):
    s = flow.slide
    w, h = _tab_height(text)
    tab = _rect(s, MSO_SHAPE.SNIP_1_RECTANGLE, PC_PAGE_MARGIN, y, w, h, fill=PC_NAVY)
    tab.adjustments[0] = 0.3
    # 번호와 제목을 하나의 도형 텍스트로 넣으면 세로 중앙 정렬이 자동으로 맞는다
    tf = tab.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left, tf.margin_right = Inches(0.14), Inches(0.14)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r1 = p.add_run()
    r1.text = f"{number_label}   "
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x9F, 0xC0, 0xE0)
    r1.font.name = FONT_NAME
    r2 = p.add_run()
    r2.text = text
    r2.font.size = Pt(12)
    r2.font.bold = True
    r2.font.color.rgb = WHITE
    r2.font.name = FONT_NAME
    return h


def pc_group_bar_h(text):
    return max(Inches(0.34), text_h(text, 11.5, PC_CW - Inches(0.4), bold=True, ls=1.0, para_gap_pt=0) + Inches(0.14))


def pc_group_bar(flow, y, text):
    """방문지 묶음(지역/일차) 바: 옅은 파란 배경 + 왼쪽 네이비 세로선."""
    s = flow.slide
    h = pc_group_bar_h(text)
    _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, h, fill=PC_PALE_BAR)
    _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, Inches(0.07), h, fill=PC_NAVY)
    tb_h = text_h(text, 11.5, PC_CW - Inches(0.4), bold=True, ls=1.0, para_gap_pt=0)
    pc_text(s, PC_PAGE_MARGIN + Inches(0.22), y + (h - tb_h) / 2, PC_CW - Inches(0.4), text, size=11.5,
            color=PC_NAVY, bold=True, ls=1.0)
    return h


PC_IMG_H = Inches(0.82)


def pc_image_row_h(n):
    return PC_IMG_H


def pc_image_row(flow, y, numbers, label_prefix="이미지", x=None, width=None):
    """번호 붙은 이미지 자리 n칸을 (x, width) 영역 안에서 가운데 정렬로 한 줄에 놓는다(최대 3칸)."""
    s = flow.slide
    n = len(numbers)
    x = PC_PAGE_MARGIN if x is None else x
    width = width or PC_CW
    gap = Inches(0.14)
    w = min(Inches(2.15), int((width - gap * (n - 1)) / n))
    total = w * n + gap * (n - 1)
    x0 = x + int((width - total) / 2)
    for i, no in enumerate(numbers):
        x = x0 + (w + gap) * i
        box = _rect(s, MSO_SHAPE.RECTANGLE, x, y, w, PC_IMG_H, fill=PC_PALE, line=PC_LINE_BLUE,
                    line_w=1.0, dash=True)
        add_icon(s, "image", x + int((w - Inches(0.28)) / 2), y + Inches(0.14), Inches(0.28))
        pc_text(s, x, y + Inches(0.5), w, f"{label_prefix} {no}", size=10, color=PC_LINE_BLUE, bold=True,
                align=PP_ALIGN.CENTER, ls=1.0)
    return PC_IMG_H


def _image_count(item, default=2):
    try:
        n = int(item.get("image_count", default))
    except (TypeError, ValueError):
        n = default
    return max(1, min(3, n))


def pc_chip_h(text, size=10, width=None, with_icon=True):
    width = width or PC_CW
    inner = width - Inches(0.28) - (Inches(0.3) if with_icon else 0)
    return max(Inches(0.32), text_h(text, size, inner, bold=True, ls=1.0, para_gap_pt=0) + Inches(0.14))


def pc_chip(flow, y, text, icon="footprints", size=10, x=None, width=None):
    """옅은 파란 칩(코스 스펙 한 줄 등): [아이콘] 굵은 네이비 글씨."""
    s = flow.slide
    x = PC_PAGE_MARGIN if x is None else x
    width = width or PC_CW
    h = pc_chip_h(text, size, width, with_icon=bool(icon))
    _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, width, h, fill=PC_PALE, line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.25
    tx = x + Inches(0.14)
    if icon:
        add_icon(s, icon, tx, y + (h - Inches(0.2)) / 2, Inches(0.2))
        tx += Inches(0.3)
    tb_h = text_h(text, size, width - (tx - x) - Inches(0.14), bold=True, ls=1.0, para_gap_pt=0)
    pc_text(s, tx, y + (h - tb_h) / 2, width - (tx - x) - Inches(0.14), text, size=size, color=PC_NAVY,
            bold=True, ls=1.0)
    return h


def pc_body(flow, y, text, x=None, width=None, size=PC_BODY_SIZE, color=PC_BODY, align=PP_ALIGN.LEFT, bold=False):
    x = PC_PAGE_MARGIN if x is None else x
    width = width or PC_CW
    return pc_text(flow.slide, x, y, width, text, size=size, color=color, align=align, bold=bold)


def pc_body_h(text, width=None, size=PC_BODY_SIZE, bold=False):
    return text_h(text, size, width or PC_CW, bold=bold)


# ---- 섹션 ---------------------------------------------------------------------------

def _option_lines(cover, options_key, single_key):
    text = cover_option_text(cover, options_key, single_key)
    return [ln for ln in str(text or "").split("\n") if ln.strip()]


def _has_options(cover, key):
    opts = cover.get(key)
    return isinstance(opts, list) and len([o for o in opts if str(o).strip()]) >= 2


def _option_items(cover, options_key, types_key):
    """후보안 [(라벨, 문구), ...]. 라벨 = "안1 · 감성형 · 추천"(첫 번째 안이 추천안).
    types 배열이 없거나 짧으면 해당 안은 "안N"만 표시(하위 호환)."""
    opts = [str(o).strip() for o in (cover.get(options_key) or []) if str(o).strip()]
    types = cover.get(types_key) or []
    items = []
    for i, text in enumerate(opts):
        text = re.sub(r"^\s*안\s*\d+\s*[\)\.]\s*", "", text)  # AI가 "안1)"을 직접 붙인 경우 제거(중복 방지)
        label = f"안{i + 1}"
        t = str(types[i]).strip() if i < len(types) and types[i] else ""
        if t:
            label += f" · {t}"
        if i == 0:
            label += " · 추천"
        items.append((label, text))
    return items


_PILL_W = Inches(1.55)


def _candidate_card_h(items, size, caption=True, bold=False):
    inner_w = PC_CW - Inches(0.2) - Inches(0.3) - _PILL_W - Inches(0.14)
    h = Inches(0.12) + (Inches(0.22) if caption else 0)
    for _, text in items:
        h += max(Inches(0.26), text_h(text, size, inner_w, bold=bold, ls=1.0, para_gap_pt=0)) + Inches(0.09)
    return h + Inches(0.04)


def _candidate_card(flow, y, caption, items, size, bold=False):
    """후보안 카드: 한 줄에 [라벨 알약] + 문구. 첫 번째(추천안)는 네이비로 채워 구분한다."""
    s = flow.slide
    h = _candidate_card_h(items, size, caption=bool(caption), bold=bold)
    cx = PC_PAGE_MARGIN + Inches(0.1)
    _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, y, PC_CW - Inches(0.2), h,
          fill=PC_PALE, line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.1
    cy = y + Inches(0.1)
    if caption:
        pc_text(s, PC_PAGE_MARGIN, cy, PC_CW, caption, size=8, color=PC_LINE_BLUE, bold=True,
                align=PP_ALIGN.CENTER, ls=1.0)
        cy += Inches(0.22)
    text_x = cx + Inches(0.15) + _PILL_W + Inches(0.14)
    text_w = PC_CW - Inches(0.2) - Inches(0.3) - _PILL_W - Inches(0.14)
    for i, (label, text) in enumerate(items):
        th = text_h(text, size, text_w, bold=bold, ls=1.0, para_gap_pt=0)
        row_h = max(Inches(0.26), th)
        pill_h = Inches(0.24)
        pill = _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.15), cy + (row_h - pill_h) / 2, _PILL_W, pill_h,
                     fill=PC_NAVY if i == 0 else WHITE, line=PC_NAVY, line_w=0.75)
        pill.adjustments[0] = 0.5
        _shape_text(pill, label, 8.5, WHITE if i == 0 else PC_NAVY, bold=True, margins=(0.04, 0, 0.04, 0))
        pc_text(s, text_x, cy + (row_h - th) / 2, text_w, text, size=size, color=PC_BODY, bold=bold, ls=1.0,
                para_gap_pt=0)
        cy += row_h + Inches(0.09)
    return h


def balance_lines(text, size_pt, width_emu, bold=False):
    """한 줄에 안 들어가는 짧은 제목은 공백 기준으로 가운데에서 두 줄로 나눠 한 글자만 남는 줄을 막는다."""
    text = str(text or "")
    if "\n" in text or count_lines(text, size_pt, width_emu, bold) != 2 or " " not in text:
        return text
    mid = len(text) / 2
    cands = [i for i, c in enumerate(text) if c == " "]
    cut = min(cands, key=lambda i: abs(i - mid))
    return text[:cut] + "\n" + text[cut + 1:]


def build_pc_cover(flow, cover, watermark_label=""):
    slide = flow.new_slide()
    y = flow.y
    # 표지가 한 장을 넘지 않도록 메인 이미지 자리 높이를 남는 공간에 맞춰 줄인다(최대 2.0in, 최소 1.1in)
    _tl = _option_lines(cover, "tagline_options", "tagline")
    _sl = _option_lines(cover, "subtitle_options", "subtitle")
    _pn = balance_lines(cover.get("product_name", ""), 27, PC_CW, bold=True)
    fixed = Inches(0.1) + Inches(0.26)
    if _tl:
        fixed += (_candidate_card_h(_option_items(cover, "tagline_options", "tagline_option_types"), 12) + Inches(0.22)
                  if _has_options(cover, "tagline_options")
                  else text_h("\n".join(_tl), 13, PC_CW) + Inches(0.12))
    if _pn:
        fixed += text_h(_pn, 27, PC_CW, bold=True, ls=1.0, para_gap_pt=0) + Inches(0.3)
    if cover.get("region_tag"):
        fixed += Inches(0.52)
    if _sl:
        fixed += (_candidate_card_h(_option_items(cover, "subtitle_options", "subtitle_option_types"), 12.5, bold=True)
                  + Inches(0.2) if _has_options(cover, "subtitle_options")
                  else text_h("\n".join(_sl), 15, PC_CW, bold=True) + Inches(0.14))
    if cover.get("intro_copy"):
        fixed += text_h(cover["intro_copy"], 11, PC_CW - Inches(0.6)) + Inches(0.05)
    img_h = int(max(Inches(1.1), min(Inches(2.0), flow.bottom_limit - flow.top - fixed)))
    if watermark_label:
        pc_text(slide, SLIDE_W - PC_PAGE_MARGIN - Inches(1.6), Inches(0.18), Inches(1.6), watermark_label,
                size=10, bold=True, color=RGBColor(0xB8, 0x9B, 0x00), align=PP_ALIGN.RIGHT, ls=1.0)
    y += Inches(0.1)

    tag_lines = _option_lines(cover, "tagline_options", "tagline")
    multi_tag = _has_options(cover, "tagline_options")
    if tag_lines:
        if multi_tag:
            y += _candidate_card(flow, y, "태그라인 후보 (유형별 3안 · 안1이 추천안)",
                                 _option_items(cover, "tagline_options", "tagline_option_types"), 12) + Inches(0.22)
        else:
            y += pc_text(slide, PC_PAGE_MARGIN, y, PC_CW, "\n".join(tag_lines), size=13, color=PC_SUB,
                         align=PP_ALIGN.CENTER) + Inches(0.12)

    product_name = balance_lines(cover.get("product_name", ""), 27, PC_CW, bold=True)
    if product_name:
        y += pc_text(slide, PC_PAGE_MARGIN, y, PC_CW, product_name, size=27, bold=True, color=PC_NAVY,
                     align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0) + Inches(0.1)
        _rect(slide, MSO_SHAPE.RECTANGLE, int((SLIDE_W - Inches(0.9)) / 2), y, Inches(0.9), Pt(3), fill=PC_LINE_BLUE)
        y += Inches(0.2)

    if cover.get("region_tag"):
        tag = cover["region_tag"]
        pw = min(PC_CW, Emu(int(_word_w(tag, True) * 11 * 12700)) + Inches(0.6))
        pill = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, int((SLIDE_W - pw) / 2), y, pw, Inches(0.32),
                     fill=None, line=PC_NAVY, line_w=1.0)
        pill.adjustments[0] = 0.5
        _shape_text(pill, tag, 11, PC_NAVY, bold=True)
        y += Inches(0.32) + Inches(0.2)

    # 메인 이미지 자리 (번호 체계 밖 — 표지 대표 이미지)
    box = _rect(slide, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, img_h, fill=PC_PALE,
                line=PC_LINE_BLUE, line_w=1.0, dash=True)
    add_icon(slide, "image", PC_PAGE_MARGIN + int((PC_CW - Inches(0.4)) / 2), y + int(img_h / 2) - Inches(0.38), Inches(0.4))
    pc_text(slide, PC_PAGE_MARGIN, y + int(img_h / 2) + Inches(0.1), PC_CW, "메인 이미지", size=11, bold=True,
            color=PC_LINE_BLUE, align=PP_ALIGN.CENTER, ls=1.0)
    y += img_h + Inches(0.26)

    sub_lines = _option_lines(cover, "subtitle_options", "subtitle")
    if sub_lines:
        if _has_options(cover, "subtitle_options"):
            y += _candidate_card(flow, y, "서브카피 후보 (유형별 3안 · 안1이 추천안)",
                                 _option_items(cover, "subtitle_options", "subtitle_option_types"), 12.5, bold=True) + Inches(0.2)
        else:
            y += pc_text(slide, PC_PAGE_MARGIN, y, PC_CW, "\n".join(sub_lines), size=15, bold=True, color=PC_NAVY,
                         align=PP_ALIGN.CENTER) + Inches(0.14)
    if cover.get("intro_copy"):
        y += pc_text(slide, PC_PAGE_MARGIN + Inches(0.3), y, PC_CW - Inches(0.6), cover["intro_copy"],
                     size=11, color=PC_SUB, align=PP_ALIGN.CENTER) + Inches(0.05)
    flow.y = y
    return slide


def build_pc_background(flow, background_story):
    if not background_story:
        return None
    title = background_story.get("title", "")
    content = background_story.get("content", "")
    if not (title or content):
        return None
    tw = PC_CW - Inches(0.4)
    title_h = text_h(title, 16, tw, bold=True, ls=1.0, para_gap_pt=0) if title else Emu(0)
    content_h = text_h(content, PC_BODY_SIZE, PC_CW)
    total = (title_h + Inches(0.3) if title else Emu(0)) + content_h
    y = flow.ensure(total, gap_before=Inches(0.3))
    s = flow.slide
    if title:
        pc_text(s, PC_PAGE_MARGIN + Inches(0.2), y, tw, title, size=16, bold=True, color=PC_NAVY,
                align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
        y += title_h + Inches(0.1)
        _rect(s, MSO_SHAPE.RECTANGLE, int((SLIDE_W - Inches(0.6)) / 2), y, Inches(0.6), Pt(2.5), fill=PC_LINE_BLUE)
        y += Inches(0.2)
    pc_body(flow, y, content)
    flow.y = y + content_h
    return s


def build_pc_reasons(flow, why_reasons, product_name="", heading=None):
    """'{상품명} 포인트 0N' — H1 + 번호 탭 항목(이미지 자리 + 본문)."""
    if not why_reasons:
        return None
    heading = heading or f"{product_name} 포인트 {len(why_reasons):02d}".strip()
    metrics = []
    for i, r in enumerate(why_reasons):
        n_img = _image_count(r)
        tab_h = _tab_height(r.get("title", ""))[1]
        body_h = pc_body_h(r.get("content", ""))
        metrics.append((n_img, tab_h, body_h, tab_h + Inches(0.12) + PC_IMG_H + Inches(0.12) + body_h))

    first = True
    for i, (r, (n_img, tab_h, body_h, blk_h)) in enumerate(zip(why_reasons, metrics)):
        need = blk_h + (PC_H1_H + Inches(0.2) if first else Inches(0))
        y = flow.ensure(need)
        slide = flow.slide
        if first:
            pc_h1(flow, y, heading, icon="sparkles")
            y += PC_H1_H + Inches(0.2)
            first = False
        pc_tab(flow, y, f"{i + 1:02d}", r.get("title", ""))
        y += tab_h + Inches(0.12)
        pc_image_row(flow, y, flow.take_images(n_img))
        y += PC_IMG_H + Inches(0.12)
        pc_body(flow, y, r.get("content", ""))
        flow.y = y + body_h
    return flow.slide


def _is_day_tag(tag):
    return bool(re.match(r"^\s*\d+\s*일\s*차", tag or ""))


def build_pc_destinations(flow, destinations, section_title=None, icon="route", default_title="코스 하이라이트"):
    """코스/방문지 목록.
    - region_tag가 짧은 "N일차"면: 제목 앞에 일차 배지(네이비 알약) — 사람 기획안의 일차 배지와 같은 역할
    - 그 외 region_tag(지역/테마): 그룹 바(옅은 파랑 + 네이비 세로선) 아래에 번호 원 + 제목
    - 이어서 스펙 칩(stats) > 번호 붙은 이미지 자리 > 본문."""
    if not destinations:
        return None
    heading = section_title or default_title
    current_group = None
    drew_heading = False
    circle = Inches(0.36)

    for idx, dest in enumerate(destinations):
        title = dest.get("title", "")
        group = dest.get("region_tag") or None
        day_badge = group if (group and _is_day_tag(group) and len(group) <= 8) else None
        bar_group = None if day_badge else group
        new_group = bool(bar_group) and bar_group != current_group
        badge_w = Emu(int(_word_w(day_badge, True) * 10.5 * 12700)) + Inches(0.36) if day_badge else circle
        badge_w = max(badge_w, Inches(0.62)) if day_badge else circle
        off = badge_w + Inches(0.14)
        tw = PC_CW - off
        title_h = max(circle, text_h(title, 13, tw, bold=True, ls=1.0, para_gap_pt=0))
        stats = (dest.get("stats") or "").strip()
        stats_h = pc_chip_h(stats, 10, PC_CW - off) if stats else Emu(0)
        n_img = _image_count(dest)
        desc = dest.get("description", "")
        desc_h = pc_body_h(desc, PC_CW - off)
        group_h = (pc_group_bar_h(bar_group) + Inches(0.14)) if new_group else Emu(0)
        block_h = title_h + Inches(0.1) + (stats_h + Inches(0.1) if stats else 0) + PC_IMG_H + Inches(0.1) + desc_h
        need = group_h + block_h + (PC_H1_H + Inches(0.18) if not drew_heading else 0)

        y = flow.ensure(need, gap_before=Inches(0.26))
        slide = flow.slide
        if not drew_heading:
            pc_h1(flow, y, heading, icon=icon)
            y += PC_H1_H + Inches(0.18)
            drew_heading = True
        # 그룹이 페이지를 넘어 이어질 때(새 그룹이 아닌데 페이지 맨 위) 그룹 바를 다시 보여준다
        if not new_group and bar_group and abs(y - flow.top) < Inches(0.01):
            y += pc_group_bar(flow, y, bar_group) + Inches(0.14)
        if new_group:
            y += pc_group_bar(flow, y, bar_group) + Inches(0.14)
            current_group = bar_group
        elif not bar_group:
            current_group = None

        by = y + (title_h - circle) / 2
        if day_badge:
            pill = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, PC_PAGE_MARGIN, by, badge_w, circle, fill=PC_NAVY)
            pill.adjustments[0] = 0.5
            _shape_text(pill, day_badge, 10.5, WHITE, bold=True, margins=(0, 0, 0, 0))
        else:
            oval = _rect(slide, MSO_SHAPE.OVAL, PC_PAGE_MARGIN, by, circle, circle, fill=PC_NAVY)
            _shape_text(oval, f"{idx + 1}", 12, WHITE, bold=True, margins=(0, 0, 0, 0))
        t_h = text_h(title, 13, tw, bold=True, ls=1.0, para_gap_pt=0)
        pc_text(slide, PC_PAGE_MARGIN + off, y + (title_h - t_h) / 2, tw, title, size=13, bold=True,
                color=PC_BODY, ls=1.0, para_gap_pt=0)
        y += title_h + Inches(0.1)
        if stats:
            pc_chip(flow, y, stats, icon="footprints", x=PC_PAGE_MARGIN + off, width=PC_CW - off)
            y += stats_h + Inches(0.1)
        pc_image_row(flow, y, flow.take_images(n_img), x=PC_PAGE_MARGIN + off, width=PC_CW - off)
        y += PC_IMG_H + Inches(0.1)
        pc_body(flow, y, desc, x=PC_PAGE_MARGIN + off, width=PC_CW - off)
        flow.y = y + desc_h
    return flow.slide


def build_pc_route_compare(flow, route_compare):
    if not route_compare or not route_compare.get("routes"):
        return None
    routes = route_compare["routes"]
    crit = [("course", "코스"), ("scenery", "풍경"), ("appeal", "매력"), ("summary", "한줄 요약")]
    label_w = Inches(0.9)
    col_w = int((PC_CW - label_w) / len(routes))
    cell_w = col_w - Inches(0.16)
    row_hs = []
    for key, _ in crit:
        row_hs.append(max(text_h(str(r.get(key, "")), 10, cell_w, ls=1.0, para_gap_pt=0) for r in routes) + Inches(0.16))
    head_h = Inches(0.36)
    total = PC_H1_H + Inches(0.18) + head_h + sum(row_hs)
    y = flow.ensure(total)
    s = flow.slide
    pc_h1(flow, y, route_compare.get("title") or "노선 비교", icon="route")
    y += PC_H1_H + Inches(0.18)
    for ri, route in enumerate(routes):
        x = PC_PAGE_MARGIN + label_w + col_w * ri
        c = _rect(s, MSO_SHAPE.RECTANGLE, x, y, col_w, head_h, fill=PC_NAVY, line=WHITE)
        _shape_text(c, route.get("name", ""), 11, WHITE, bold=True)
    y += head_h
    for (key, label), rh in zip(crit, row_hs):
        c = _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, label_w, rh, fill=PC_PALE_BAR, line=WHITE)
        _shape_text(c, label, 10, PC_NAVY, bold=True)
        for ri, route in enumerate(routes):
            x = PC_PAGE_MARGIN + label_w + col_w * ri
            c = _rect(s, MSO_SHAPE.RECTANGLE, x, y, col_w, rh, fill=PC_PALE if key != "summary" else WHITE,
                      line=RGBColor(0xD5, 0xDF, 0xEA))
            _shape_text(c, str(route.get(key, "")), 10, PC_BODY, align=PP_ALIGN.LEFT,
                        margins=(0.08, 0.05, 0.08, 0.05))
        y += rh
    flow.y = y
    return s


def build_pc_transport(flow, transport_spec, image_label="이동수단 이미지", with_image=True, icon="train-front"):
    if not transport_spec or not transport_spec.get("specs"):
        return None
    label_w = Inches(1.5)
    rows = []
    total = PC_H1_H + Inches(0.18) + ((PC_IMG_H + Inches(0.14)) if with_image else 0)
    for spec in transport_spec["specs"]:
        v = str(spec.get("value", ""))
        rh = max(Inches(0.34), text_h(v, 10.5, PC_CW - label_w - Inches(0.2), ls=1.0, para_gap_pt=0) + Inches(0.14))
        rows.append((spec.get("label", ""), v, rh))
        total += rh
    y = flow.ensure(total)
    s = flow.slide
    pc_h1(flow, y, transport_spec.get("title") or "이동수단 안내", icon=icon)
    y += PC_H1_H + Inches(0.18)
    if with_image:
        pc_image_row(flow, y, flow.take_images(1), label_prefix=image_label)
        y += PC_IMG_H + Inches(0.14)
    for label, v, rh in rows:
        c = _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, label_w, rh, fill=PC_PALE_BAR, line=WHITE)
        _shape_text(c, label, 10.5, PC_NAVY, bold=True)
        c2 = _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN + label_w, y, PC_CW - label_w, rh, fill=None,
                   line=RGBColor(0xD5, 0xDF, 0xEA))
        _shape_text(c2, v, 10.5, PC_BODY, align=PP_ALIGN.LEFT, margins=(0.1, 0.04, 0.1, 0.04))
        y += rh
    flow.y = y
    return s


def build_pc_guide(flow, guide_profile, heading="함께하는 사람들"):
    if not guide_profile:
        return None
    first = True
    pic = Inches(1.05)
    for guide in guide_profile:
        name_title = guide.get("name", "")
        if guide.get("title"):
            name_title = f"{name_title}  ({guide['title']})" if name_title else guide["title"]
        bio = guide.get("bio", "")
        tx_w = PC_CW - pic - Inches(0.2)
        name_h = text_h(name_title, 12.5, tx_w, bold=True, ls=1.0, para_gap_pt=0)
        bio_h = pc_body_h(bio, tx_w, size=10.5)
        block_h = max(pic, name_h + Inches(0.08) + bio_h)
        need = block_h + (PC_H1_H + Inches(0.18) if first else 0)
        y = flow.ensure(need)
        s = flow.slide
        if first:
            pc_h1(flow, y, heading, icon="users")
            y += PC_H1_H + Inches(0.18)
            first = False
        box = _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, pic, pic, fill=PC_PALE, line=PC_LINE_BLUE, line_w=1.0, dash=True)
        _shape_text(box, "프로필 사진", 9, PC_LINE_BLUE, bold=True)
        pc_text(s, PC_PAGE_MARGIN + pic + Inches(0.2), y, tx_w, name_title, size=12.5, bold=True, color=PC_NAVY, ls=1.0, para_gap_pt=0)
        pc_body(flow, y + name_h + Inches(0.08), bio, x=PC_PAGE_MARGIN + pic + Inches(0.2), width=tx_w, size=10.5)
        flow.y = y + block_h
    return flow.slide


def build_pc_experience(flow, brand_tagline, experience_points, heading="혜초와 함께라면 더 특별한 이유"):
    """브랜드 한 줄 + 경험 포인트 아이콘 카드(최대 4열)."""
    if not brand_tagline and not experience_points:
        return None
    pts = []
    for p in (experience_points or []):
        p = dict(p)
        # "혜초와 함께하면, ○○" 같은 앞머리는 카드 제목을 흐리게 하므로(섹션 제목이 이미 그 뜻) 떼고 이점만 남긴다
        cleaned = re.sub(r"^\s*혜초와\s*함께\s*(하면|라면|하는|한다면)\s*[,，:]?\s*", "", str(p.get("title", "")))
        if cleaned.strip():
            p["title"] = cleaned.strip()
        pts.append(p)
    n = len(pts)
    tag_h = text_h(brand_tagline, 14, PC_CW - Inches(0.6), bold=True, ls=1.0, para_gap_pt=0) if brand_tagline else Emu(0)
    gap = Inches(0.12)
    cols = min(max(n, 1), 4)
    card_w = int((PC_CW - gap * (cols - 1)) / cols)
    inner_w = card_w - Inches(0.24)
    rows_n = -(-n // cols) if n else 0
    title_hs = [text_h(p.get("title", ""), 11.5, inner_w, bold=True, ls=1.0, para_gap_pt=0) for p in pts]
    desc_hs = [text_h(p.get("description", ""), 10, inner_w, ls=1.0, para_gap_pt=0) for p in pts]
    row_card_hs = []
    for r in range(rows_n):
        sl = slice(r * cols, (r + 1) * cols)
        row_card_hs.append(Inches(0.2) + Inches(0.52) + Inches(0.12) + max(title_hs[sl]) + Inches(0.08) + max(desc_hs[sl]) + Inches(0.2))
    total = (PC_H1_H + Inches(0.2) if True else 0) + (tag_h + Inches(0.2) if brand_tagline else 0) \
        + sum(row_card_hs) + gap * max(rows_n - 1, 0)
    y = flow.ensure(total)
    s = flow.slide
    pc_h1(flow, y, heading, icon="badge-check")
    y += PC_H1_H + Inches(0.2)
    if brand_tagline:
        pc_text(s, PC_PAGE_MARGIN + Inches(0.3), y, PC_CW - Inches(0.6), brand_tagline, size=14, bold=True,
                color=PC_NAVY, align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
        y += tag_h + Inches(0.2)
    for r in range(rows_n):
        ch = row_card_hs[r]
        for c in range(cols):
            i = r * cols + c
            if i >= n:
                break
            p = pts[i]
            x = PC_PAGE_MARGIN + (card_w + gap) * c
            _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, ch, fill=PC_PALE,
                  line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.07
            cx = x + int((card_w - Inches(0.52)) / 2)
            _rect(s, MSO_SHAPE.OVAL, cx, y + Inches(0.2), Inches(0.52), Inches(0.52), fill=WHITE, line=PC_LINE_BLUE)
            key = pick_icon(p.get("title", ""), p.get("description", ""), p.get("icon"))
            add_icon(s, key, cx + Inches(0.1), y + Inches(0.3), Inches(0.32))
            ty = y + Inches(0.2) + Inches(0.52) + Inches(0.12)
            pc_text(s, x + Inches(0.12), ty, inner_w, p.get("title", ""), size=11.5, bold=True, color=PC_NAVY,
                    align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
            pc_text(s, x + Inches(0.12), ty + max(title_hs[r * cols:(r + 1) * cols]) + Inches(0.08), inner_w,
                    p.get("description", ""), size=10, color=PC_SUB, align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
        y += ch + gap
    flow.y = y - gap
    return s


def _pc_table(slide, x, y, widths, rows, row_h, header=True, size=9, align=PP_ALIGN.CENTER):
    """rows[0]이 머리글. 머리글 네이비, 첫 열 옅은 파랑, 나머지 흰/옅은 줄무늬."""
    n_rows, n_cols = len(rows), len(rows[0])
    row_hs = list(row_h) if isinstance(row_h, (list, tuple)) else [row_h] * n_rows
    shape = slide.shapes.add_table(n_rows, n_cols, x, y, sum(widths), sum(row_hs))
    tbl = shape.table
    for ci, w in enumerate(widths):
        tbl.columns[ci].width = w
    for ri in range(n_rows):
        tbl.rows[ri].height = row_hs[ri]
        for ci in range(n_cols):
            cell = tbl.cell(ri, ci)
            cell.text = str(rows[ri][ci])
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.margin_left = cell.margin_right = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0 and header:
                cell.fill.fore_color.rgb = PC_NAVY
                fc = WHITE
            elif ci == 0:
                cell.fill.fore_color.rgb = PC_PALE_BAR
                fc = PC_NAVY
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 else PC_PALE
                fc = PC_BODY
            for para in cell.text_frame.paragraphs:
                para.alignment = align if ((ri > 0 or not header) and ci > 0) else PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.size = Pt(size)
                    run.font.name = FONT_NAME
                    run.font.bold = (ri == 0 or ci == 0)
                    run.font.color.rgb = fc
    return shape


def build_pc_season(flow, season, season_table=None):
    if not season or (not season.get("content") and not season_table):
        return None
    stat = season.get("stat_line")
    content = season.get("content", "")
    stat_h = pc_chip_h(stat, 11, PC_CW, with_icon=True) if stat else Emu(0)
    content_h = pc_body_h(content)
    has_rain = bool(season_table) and any((r.get("rain") or "").strip() for r in season_table)
    n_rows = (3 if has_rain else 2) + 1
    row_h = Inches(0.27)
    chart_h = Inches(0.55)
    table_h = chart_h + Inches(0.12) + row_h * n_rows if season_table else Emu(0)
    total = PC_H1_H + Inches(0.18) + (stat_h + Inches(0.12) if stat else 0) + content_h \
        + (Inches(0.2) + table_h if season_table else 0)
    y = flow.ensure(total)
    s = flow.slide
    pc_h1(flow, y, season.get("title", "언제 가면 좋을까?"), icon="sun")
    y += PC_H1_H + Inches(0.18)
    if stat:
        pc_chip(flow, y, stat, icon="calendar-days", size=11)
        y += stat_h + Inches(0.12)
    if content:
        pc_body(flow, y, content)
        y += content_h
    if season_table:
        y += Inches(0.2)
        box = _rect(s, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, chart_h, fill=PC_PALE, line=PC_LINE_BLUE,
                    line_w=1.0, dash=True)
        _shape_text(box, "월별 기온 차트 (자리표시 — 실제 그래픽은 디자이너 작업)", 9.5, PC_LINE_BLUE, bold=True)
        y += chart_h + Inches(0.12)
        label_w = Inches(1.0)
        n_m = len(season_table)
        col_w = int((PC_CW - label_w) / n_m)
        widths = [label_w] + [col_w] * n_m
        labels = ["구분", "최저기온 (°C)", "최고기온 (°C)"] + (["강수량 (mm)"] if has_rain else [])
        keys = ["month", "low", "high"] + (["rain"] if has_rain else [])
        rows = [[labels[i]] + [str(m.get(k, "")) for m in season_table] for i, k in enumerate(keys)]
        _pc_table(s, PC_PAGE_MARGIN, y, widths, rows, row_h, size=9)
        y += row_h * n_rows
    flow.y = y
    return s


def build_pc_qna(flow, qna_items, heading="궁금증을 풀어드려요"):
    """qna_items = [(icon_key, qa_dict), ...] — 식사/난이도·안전 안내를 아이콘 카드 목록으로."""
    items = [(ic, qa) for ic, qa in qna_items if qa and qa.get("question")]
    if not items:
        return None
    ic_col = Inches(0.62)
    tw = PC_CW - ic_col - Inches(0.1)
    first = True
    for ic, qa in items:
        q_h = text_h(qa["question"], 12.5, tw - Inches(0.3), bold=True, ls=1.0, para_gap_pt=0)
        a_h = pc_body_h(qa.get("answer", ""), tw - Inches(0.3))
        card_h = Inches(0.16) + q_h + Inches(0.08) + a_h + Inches(0.18)
        need = card_h + (PC_H1_H + Inches(0.18) if first else 0)
        y = flow.ensure(need)
        s = flow.slide
        if first:
            pc_h1(flow, y, heading, icon="circle-help")
            y += PC_H1_H + Inches(0.18)
            first = False
        _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, card_h, fill=PC_PALE,
              line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.06
        _rect(s, MSO_SHAPE.OVAL, PC_PAGE_MARGIN + Inches(0.12), y + Inches(0.14), Inches(0.46), Inches(0.46), fill=WHITE, line=PC_LINE_BLUE)
        add_icon(s, ic, PC_PAGE_MARGIN + Inches(0.2), y + Inches(0.22), Inches(0.3))
        tx = PC_PAGE_MARGIN + ic_col + Inches(0.15)
        pc_text(s, tx, y + Inches(0.14), tw - Inches(0.3), qa["question"], size=12.5, bold=True, color=PC_NAVY,
                ls=1.0, para_gap_pt=0)
        pc_body(flow, y + Inches(0.14) + q_h + Inches(0.08), qa.get("answer", ""), x=tx, width=tw - Inches(0.3))
        flow.y = y + card_h
    return flow.slide


def build_pc_review_notes(flow, review):
    """Gemini 교차검수 결과(내부 참고용) — 새 스타일의 제목/본문으로."""
    if not review or review.get("_dry_run"):
        return []
    issues = review.get("issues") or []
    slide = flow.new_slide()
    y = flow.y
    pc_h1(flow, y, "Gemini 교차검수 결과 (내부 참고용)", icon="shield-check")
    y += PC_H1_H + Inches(0.2)
    slides = [slide]
    if not issues:
        flow.y = y + pc_body(flow, y, "검수 통과 — 왜곡·날조 / 저작권·표절 / 사실확인 / 비문·맞춤법 이슈가 발견되지 않았습니다.",
                              color=PC_SUB)
        return slides
    if review.get("summary"):
        y += pc_body(flow, y, review["summary"], color=PC_SUB, size=10.5) + Inches(0.18)
    for issue in issues:
        head = f"[{issue.get('category', '')} · {issue.get('severity', '')}] {issue.get('field', '')}"
        head_h = text_h(head, 11, PC_CW, bold=True, ls=1.0, para_gap_pt=0)
        quote = issue.get("quote", "")
        quote_h = text_h(quote, 10, PC_CW - Inches(0.3)) if quote else Emu(0)
        exp_h = pc_body_h(issue.get("explanation", ""), size=10.5)
        block_h = head_h + Inches(0.05) + (quote_h + Inches(0.05) if quote else 0) + exp_h + Inches(0.2)
        if y + block_h > flow.bottom_limit:
            slide = flow.new_slide()
            slides.append(slide)
            y = flow.y
        pc_text(flow.slide, PC_PAGE_MARGIN, y, PC_CW, head, size=11, bold=True, color=PC_NAVY, ls=1.0, para_gap_pt=0)
        y += head_h + Inches(0.05)
        if quote:
            pc_text(flow.slide, PC_PAGE_MARGIN + Inches(0.15), y, PC_CW - Inches(0.3), f"“{quote}”", size=10, color=PC_SUB)
            y += quote_h + Inches(0.05)
        pc_body(flow, y, issue.get("explanation", ""), size=10.5)
        y += exp_h + Inches(0.2)
    flow.y = y
    return slides


# ==== PC 상품별 동적 구성(블록 라이브러리, 2026-11) =========================================
# content_json["sections"] = [{"type": 블록종류, "title": 섹션 제목, ...}, ...] 를 위에서부터
# 차례로 그린다. 표지/배너/검수 페이지는 고정이고, 그 사이 본문 구성만 상품마다 AI가 고른다.

def _s(v):
    return "" if v is None else str(v).strip()


def _items(sec, *keys):
    for k in keys + ("items",):
        v = sec.get(k)
        if isinstance(v, list) and v:
            return v
    return []


def _item_dict(it, *title_keys):
    """리스트 항목이 문자열이어도 {"title": ...}처럼 다룰 수 있게 정규화."""
    if isinstance(it, dict):
        return it
    return {"title": _s(it)}


def _first(d, *keys, default=""):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return _s(v)
    return default


def _begin(flow, heading, icon, need_after, gap=Inches(0.3)):
    """섹션 제목 바를 그리고, 제목만 페이지 끝에 남지 않도록 첫 항목 높이까지 함께 확보한다."""
    y = flow.ensure(PC_H1_H + Inches(0.18) + need_after, gap_before=gap)
    pc_h1(flow, y, heading, icon=icon)
    flow.y = y + PC_H1_H + Inches(0.18)
    return flow.y


def _img_box(slide, x, y, w, h, label):
    _rect(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=PC_PALE, line=PC_LINE_BLUE, line_w=1.0, dash=True)
    add_icon(slide, "image", x + int((w - Inches(0.28)) / 2), y + int(h / 2) - Inches(0.26), Inches(0.28))
    pc_text(slide, x, y + int(h / 2) + Inches(0.04), w, label, size=10, color=PC_LINE_BLUE, bold=True,
            align=PP_ALIGN.CENTER, ls=1.0)


# ---- 개요(핵심 정보 칸 + 지도 자리) ----

def build_pc_overview(flow, sec):
    heading = _s(sec.get("title")) or "여행 한눈에 보기"
    facts = [_item_dict(f) for f in _items(sec, "facts")]
    facts = [f for f in facts if _first(f, "value", "title") or _first(f, "label")]
    intro = _s(sec.get("intro") or sec.get("content"))
    want_map = sec.get("map", True) not in (False, "false", "False", 0)
    if not facts and not intro:
        return None
    n = len(facts)
    cols = 0 if n == 0 else (n if n <= 4 else (3 if n <= 6 else 4))
    gap = Inches(0.12)
    card_w = int((PC_CW - gap * (cols - 1)) / cols) if cols else 0
    inner = card_w - Inches(0.2)
    rows_n = -(-n // cols) if cols else 0
    row_hs = []
    for r in range(rows_n):
        chunk = facts[r * cols:(r + 1) * cols]
        lab = max(text_h(_first(f, "label"), 9.5, inner, ls=1.0, para_gap_pt=0) for f in chunk)
        val = max(text_h(_first(f, "value", "title"), 11.5, inner, bold=True, ls=1.0, para_gap_pt=0) for f in chunk)
        row_hs.append(Inches(0.16) + Inches(0.34) + Inches(0.08) + lab + Inches(0.05) + val + Inches(0.16))
    intro_h = pc_body_h(intro) if intro else Emu(0)
    map_h = Inches(1.55) if want_map else Emu(0)
    first_need = (intro_h + Inches(0.14)) if intro else (row_hs[0] if row_hs else map_h)
    y = _begin(flow, heading, sec.get("icon") or "map", first_need)
    s = flow.slide
    if intro:
        pc_body(flow, y, intro)
        y += intro_h + Inches(0.14)
        flow.y = y
    for r in range(rows_n):
        y = flow.ensure(row_hs[r], gap_before=Inches(0.0) if r == 0 else Inches(0.12))
        s = flow.slide
        chunk = facts[r * cols:(r + 1) * cols]
        for c, f in enumerate(chunk):
            x = PC_PAGE_MARGIN + (card_w + gap) * c
            _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, row_hs[r], fill=PC_PALE,
                  line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.08
            ic = pick_icon(_first(f, "label"), _first(f, "value", "title"), f.get("icon"))
            add_icon(s, ic, x + int((card_w - Inches(0.34)) / 2), y + Inches(0.16), Inches(0.34))
            ty = y + Inches(0.16) + Inches(0.34) + Inches(0.08)
            lab_h = max(text_h(_first(g, "label"), 9.5, inner, ls=1.0, para_gap_pt=0) for g in chunk)
            pc_text(s, x + Inches(0.1), ty, inner, _first(f, "label"), size=9.5, color=PC_SUB,
                    align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
            pc_text(s, x + Inches(0.1), ty + lab_h + Inches(0.05), inner, _first(f, "value", "title"), size=11.5,
                    bold=True, color=PC_NAVY, align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
        y += row_hs[r]
        flow.y = y
    if want_map:
        y = flow.ensure(map_h, gap_before=Inches(0.14))
        box = _rect(flow.slide, MSO_SHAPE.RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, map_h, fill=PC_PALE,
                    line=PC_LINE_BLUE, line_w=1.0, dash=True)
        _shape_text(box, _s(sec.get("map_label")) or "여정 지도 (자리표시 — 실제 지도는 디자이너 작업)", 10,
                    PC_LINE_BLUE, bold=True)
        flow.y = y + map_h
    return flow.slide


# ---- 아이콘 카드/행(추천 대상, 체크 리스트 등) ----

def _icon_rows(flow, heading, heading_icon, items, default_icon, card=True):
    items = [_item_dict(i) for i in items]
    items = [i for i in items if _first(i, "title", "question", "name") or _first(i, "description", "answer", "content")]
    if not items:
        return None
    off = Inches(0.77) if card else Inches(0.44)
    tw = PC_CW - off - (Inches(0.15) if card else 0)
    first = True
    for it in items:
        t = _first(it, "title", "question", "name")
        d = _first(it, "description", "answer", "content")
        t_h = text_h(t, 12 if card else 11.5, tw, bold=True, ls=1.0, para_gap_pt=0) if t else Emu(0)
        d_h = pc_body_h(d, tw, size=10.5) if d else Emu(0)
        body_h = t_h + (Inches(0.06) if (t and d) else 0) + d_h
        h = (Inches(0.14) + max(body_h, Inches(0.46)) + Inches(0.14)) if card else (max(body_h, Inches(0.26)) + Inches(0.06))
        if first:
            y = _begin(flow, heading, heading_icon, h)
            first = False
        else:
            y = flow.ensure(h, gap_before=Inches(0.1) if card else Inches(0.06))
        s = flow.slide
        if card:
            _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, h, fill=PC_PALE,
                  line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.06
            _rect(s, MSO_SHAPE.OVAL, PC_PAGE_MARGIN + Inches(0.12), y + int((h - Inches(0.46)) / 2), Inches(0.46),
                  Inches(0.46), fill=WHITE, line=PC_LINE_BLUE)
            ik = it.get("icon") if it.get("icon") in ICON_KEYS else None
            if not ik:
                guess = pick_icon(t, d, None)
                ik = default_icon if guess == "star" else guess
            add_icon(s, ik, PC_PAGE_MARGIN + Inches(0.2), y + int((h - Inches(0.3)) / 2), Inches(0.3))
            ty = y + Inches(0.14) + (max(body_h, Inches(0.46)) - body_h) / 2
        else:
            add_icon(s, it.get("icon") if it.get("icon") in ICON_KEYS else default_icon, PC_PAGE_MARGIN + Inches(0.04), y + Inches(0.02), Inches(0.22))
            ty = y
        tx = PC_PAGE_MARGIN + off
        if t:
            pc_text(s, tx, ty, tw, t, size=12 if card else 11.5, bold=True, color=PC_NAVY if card else PC_BODY,
                    ls=1.0, para_gap_pt=0)
            ty += t_h + (Inches(0.06) if d else 0)
        if d:
            pc_text(s, tx, ty, tw, d, size=10.5, color=PC_BODY if card else PC_SUB)
        flow.y = y + h
    return flow.slide


def build_pc_audience(flow, sec):
    return _icon_rows(flow, _s(sec.get("title")) or "이런 분께 추천해요", sec.get("icon") or "users",
                      _items(sec, "audiences"), "heart", card=True)


def build_pc_list(flow, sec):
    heading = _s(sec.get("title")) or "안내"
    return _icon_rows(flow, heading, sec.get("icon") or "circle-check", _items(sec), "circle-check", card=False)


# ---- 하이라이트(소제목 + 핵심 키워드 + 이미지 + 본문) ----

def build_pc_highlights(flow, sec):
    heading = _s(sec.get("title")) or "하이라이트"
    items = [_item_dict(i) for i in _items(sec, "highlights")]
    items = [i for i in items if _first(i, "keyword", "title") or _first(i, "description", "content")]
    if not items:
        return None
    first = True
    for it in items:
        kicker = _first(it, "kicker", "subtitle")
        kw = _first(it, "keyword", "title")
        desc = _first(it, "description", "content")
        n_img = _image_count(it)
        k_h = text_h(kicker, 10.5, PC_CW - Inches(0.4), ls=1.0, para_gap_pt=0) if kicker else Emu(0)
        w_h = text_h(kw, 15, PC_CW - Inches(0.4), bold=True, ls=1.0, para_gap_pt=0) if kw else Emu(0)
        d_h = pc_body_h(desc)
        h = (k_h + Inches(0.05) if kicker else 0) + w_h + Inches(0.12) + PC_IMG_H + Inches(0.12) + d_h
        if first:
            y = _begin(flow, heading, sec.get("icon") or "gem", h)
            first = False
        else:
            y = flow.ensure(h, gap_before=Inches(0.28))
        s = flow.slide
        if kicker:
            pc_text(s, PC_PAGE_MARGIN + Inches(0.2), y, PC_CW - Inches(0.4), kicker, size=10.5, color=PC_LINE_BLUE,
                    align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
            y += k_h + Inches(0.05)
        if kw:
            pc_text(s, PC_PAGE_MARGIN + Inches(0.2), y, PC_CW - Inches(0.4), kw, size=15, bold=True, color=PC_NAVY,
                    align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
            y += w_h + Inches(0.12)
        pc_image_row(flow, y, flow.take_images(n_img))
        y += PC_IMG_H + Inches(0.12)
        pc_body(flow, y, desc)
        flow.y = y + d_h
    return flow.slide


# ---- 사진 카드(메뉴·맛집·숙소·객실·체험 등 2열) ----

def build_pc_cards(flow, sec):
    heading = _s(sec.get("title")) or "하나하나 소개해요"
    items = [_item_dict(i) for i in _items(sec, "cards")]
    items = [i for i in items if _first(i, "title", "name")]
    if not items:
        return None
    try:
        cols = int(sec.get("columns", 2))
    except (TypeError, ValueError):
        cols = 2
    cols = max(1, min(3, cols))
    if len(items) == 1:
        cols = 1
    gap = Inches(0.14)
    card_w = int((PC_CW - gap * (cols - 1)) / cols)
    inner = card_w - Inches(0.24)
    img_h = Inches(1.05) if cols >= 2 else Inches(1.3)
    rows = [items[i:i + cols] for i in range(0, len(items), cols)]

    def card_h(it):
        t = _first(it, "title", "name")
        sub = _first(it, "subtitle", "tag")
        d = _first(it, "description", "content")
        return (Inches(0.12) + img_h + Inches(0.1) + text_h(t, 12, inner, bold=True, ls=1.0, para_gap_pt=0)
                + ((text_h(sub, 9.5, inner, bold=True, ls=1.0, para_gap_pt=0) + Inches(0.05)) if sub else 0)
                + ((text_h(d, 10, inner, ls=1.0, para_gap_pt=0) + Inches(0.06)) if d else 0) + Inches(0.14))

    first = True
    for chunk in rows:
        rh = max(card_h(i) for i in chunk)
        if first:
            y = _begin(flow, heading, sec.get("icon") or "utensils", rh)
            first = False
        else:
            y = flow.ensure(rh, gap_before=Inches(0.14))
        s = flow.slide
        for c, it in enumerate(chunk):
            x = PC_PAGE_MARGIN + (card_w + gap) * c
            _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, rh, fill=WHITE,
                  line=RGBColor(0xC9, 0xD8, 0xE8)).adjustments[0] = 0.04
            _img_box(s, x + Inches(0.12), y + Inches(0.12), card_w - Inches(0.24), img_h,
                     f"이미지 {flow.take_images(1)[0]}")
            ty = y + Inches(0.12) + img_h + Inches(0.1)
            t = _first(it, "title", "name")
            pc_text(s, x + Inches(0.12), ty, inner, t, size=12, bold=True, color=PC_NAVY, ls=1.0, para_gap_pt=0)
            ty += text_h(t, 12, inner, bold=True, ls=1.0, para_gap_pt=0)
            sub = _first(it, "subtitle", "tag")
            if sub:
                ty += Inches(0.05)
                pc_text(s, x + Inches(0.12), ty, inner, sub, size=9.5, bold=True, color=PC_LINE_BLUE, ls=1.0,
                        para_gap_pt=0)
                ty += text_h(sub, 9.5, inner, bold=True, ls=1.0, para_gap_pt=0)
            d = _first(it, "description", "content")
            if d:
                ty += Inches(0.06)
                pc_text(s, x + Inches(0.12), ty, inner, d, size=10, color=PC_SUB, ls=1.0, para_gap_pt=0)
        flow.y = y + rh
    return flow.slide


# ---- 일정 한눈에(이미지 없는 일차별 요약 행) ----

def build_pc_schedule(flow, sec):
    heading = _s(sec.get("title")) or "일정 한눈에 보기"
    items = [_item_dict(i) for i in _items(sec, "days", "schedule")]
    items = [i for i in items if _first(i, "title", "description", "content")]
    if not items:
        return None
    badge_w = Inches(0.78)
    off = badge_w + Inches(0.16)
    tw = PC_CW - off
    first = True
    for it in items:
        day = _first(it, "day", "tag", "region_tag")
        t = _first(it, "title")
        d = _first(it, "description", "content")
        t_h = text_h(t, 11.5, tw, bold=True, ls=1.0, para_gap_pt=0) if t else Emu(0)
        d_h = pc_body_h(d, tw, size=10.5) if d else Emu(0)
        h = max(Inches(0.34), t_h + (Inches(0.05) if (t and d) else 0) + d_h) + Inches(0.12)
        if first:
            y = _begin(flow, heading, sec.get("icon") or "calendar-days", h)
            first = False
        else:
            y = flow.ensure(h, gap_before=Inches(0.0))
        s = flow.slide
        if day:
            pill = _rect(s, MSO_SHAPE.ROUNDED_RECTANGLE, PC_PAGE_MARGIN, y + Inches(0.02), badge_w, Inches(0.3),
                         fill=PC_NAVY)
            pill.adjustments[0] = 0.5
            _shape_text(pill, day, 10, WHITE, bold=True, margins=(0, 0, 0, 0))
        ty = y
        if t:
            pc_text(s, PC_PAGE_MARGIN + off, ty + Inches(0.02), tw, t, size=11.5, bold=True, color=PC_BODY, ls=1.0,
                    para_gap_pt=0)
            ty += t_h + Inches(0.05)
        if d:
            pc_text(s, PC_PAGE_MARGIN + off, ty + Inches(0.02), tw, d, size=10.5, color=PC_SUB)
        ln = s.shapes.add_connector(1, PC_PAGE_MARGIN, y + h - Inches(0.03), SLIDE_W - PC_PAGE_MARGIN,
                                    y + h - Inches(0.03))
        ln.line.color.rgb = RGBColor(0xD5, 0xDF, 0xEA)
        ln.line.width = Pt(0.5)
        flow.y = y + h
    return flow.slide


# ---- 마무리 한 줄(CTA 밴드) ----

def build_pc_closing(flow, sec):
    msg = _s(sec.get("message") or sec.get("content") or sec.get("title"))
    kicker = _s(sec.get("kicker")) if sec.get("message") else ""
    sub = _s(sec.get("sub") or sec.get("description"))
    if not msg:
        return None
    inner = PC_CW - Inches(0.8)
    k_h = text_h(kicker, 10.5, inner, ls=1.0, para_gap_pt=0) if kicker else Emu(0)
    m_h = text_h(msg, 16, inner, bold=True, ls=1.05, para_gap_pt=0)
    s_h = text_h(sub, 10.5, inner, ls=1.0, para_gap_pt=0) if sub else Emu(0)
    h = Inches(0.32) + (k_h + Inches(0.08) if kicker else 0) + m_h + (Inches(0.1) + s_h if sub else 0) + Inches(0.32)
    y = flow.ensure(h, gap_before=Inches(0.4))
    s = flow.slide
    band = _rect(s, MSO_SHAPE.SNIP_1_RECTANGLE, PC_PAGE_MARGIN, y, PC_CW, h, fill=PC_NAVY)
    band.adjustments[0] = 0.15
    ty = y + Inches(0.32)
    if kicker:
        pc_text(s, PC_PAGE_MARGIN + Inches(0.4), ty, inner, kicker, size=10.5, color=RGBColor(0x9F, 0xC0, 0xE0),
                align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
        ty += k_h + Inches(0.08)
    pc_text(s, PC_PAGE_MARGIN + Inches(0.4), ty, inner, msg, size=16, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, ls=1.05, para_gap_pt=0)
    ty += m_h
    if sub:
        pc_text(s, PC_PAGE_MARGIN + Inches(0.4), ty + Inches(0.1), inner, sub, size=10.5,
                color=RGBColor(0xD5, 0xE2, 0xF0), align=PP_ALIGN.CENTER, ls=1.0, para_gap_pt=0)
    flow.y = y + h
    return s


# ---- 일반 텍스트 / 표 ----

def build_pc_text(flow, sec):
    heading = _s(sec.get("title"))
    content = _s(sec.get("content") or sec.get("description") or sec.get("message"))
    if not (heading or content):
        return None
    n_img = 0
    if sec.get("image_count") not in (None, "", 0, "0"):
        n_img = _image_count(sec)
    c_h = pc_body_h(content) if content else Emu(0)
    need = (PC_IMG_H + Inches(0.12) if n_img else 0) + c_h
    y = _begin(flow, heading or "안내", sec.get("icon") or "star", need if need else Inches(0.3))
    if n_img:
        pc_image_row(flow, y, flow.take_images(n_img))
        y += PC_IMG_H + Inches(0.12)
    if content:
        pc_body(flow, y, content)
        y += c_h
    flow.y = y
    return flow.slide


def build_pc_table_block(flow, sec):
    heading = _s(sec.get("title")) or "표로 정리했어요"
    cols = [_s(c) for c in (sec.get("columns") or [])]
    raw_rows = sec.get("rows") or []
    rows = []
    for r in raw_rows:
        if isinstance(r, dict):
            r = [r.get(c, "") for c in cols] if cols else list(r.values())
        if isinstance(r, (list, tuple)):
            rows.append([_s(c) for c in r])
    if not rows:
        return None
    n_cols = max(len(cols), max(len(r) for r in rows))
    rows = [r + [""] * (n_cols - len(r)) for r in rows]
    header = bool(cols)
    if header:
        cols = cols + [""] * (n_cols - len(cols))
        all_rows = [cols] + rows
    else:
        all_rows = rows
    first_w = Inches(1.4) if n_cols > 2 else Inches(1.7)
    if n_cols == 1:
        widths = [PC_CW]
    else:
        rest = int((PC_CW - first_w) / (n_cols - 1))
        widths = [PC_CW - rest * (n_cols - 1)] + [rest] * (n_cols - 1)
    size = 9.5 if n_cols <= 4 else 9
    heights = []
    for ri, r in enumerate(all_rows):
        mh = max(text_h(c, size, widths[ci] - Inches(0.14), bold=(ci == 0 or (header and ri == 0)), ls=1.0,
                        para_gap_pt=0) for ci, c in enumerate(r))
        heights.append(max(Inches(0.3), mh + Inches(0.14)))
    note = _s(sec.get("note"))
    note_h = text_h(note, 9.5, PC_CW, ls=1.0, para_gap_pt=0) if note else Emu(0)
    # 페이지에 다 안 들어가면 머리글을 반복하며 나눠 그린다
    i0 = 1 if header else 0
    body_idx = list(range(i0, len(all_rows)))
    pos = 0
    first_chunk = True
    while pos < len(body_idx):
        head_h = heights[0] if header else Emu(0)
        avail_base = flow.bottom_limit - flow.top - (PC_H1_H + Inches(0.18) if first_chunk else 0) - head_h
        chunk, used = [], Emu(0)
        for ix in body_idx[pos:]:
            if used + heights[ix] > avail_base and chunk:
                break
            chunk.append(ix)
            used += heights[ix]
        # 현재 페이지 남은 공간에 첫 행들이 안 들어가면 새 페이지로
        need_first = head_h + heights[chunk[0]]
        if first_chunk:
            y = _begin(flow, heading, sec.get("icon") or "list-checks", need_first)
        else:
            y = flow.ensure(need_first, gap_before=Inches(0.0))
        # 현재 페이지 남은 높이만큼만 이번 조각에 담는다
        room = flow.bottom_limit - y - head_h
        take, used = [], Emu(0)
        for ix in chunk:
            if used + heights[ix] > room and take:
                break
            take.append(ix)
            used += heights[ix]
        rows_draw = ([all_rows[0]] if header else []) + [all_rows[ix] for ix in take]
        hs = ([heights[0]] if header else []) + [heights[ix] for ix in take]
        _pc_table(flow.slide, PC_PAGE_MARGIN, y, widths, rows_draw, hs, header=header, size=size,
                  align=PP_ALIGN.LEFT)
        flow.y = y + sum(hs)
        pos += len(take)
        first_chunk = False
        if pos < len(body_idx):
            flow.new_slide()
    if note:
        y = flow.ensure(note_h, gap_before=Inches(0.08))
        pc_text(flow.slide, PC_PAGE_MARGIN, y, PC_CW, note, size=9.5, color=PC_SUB, ls=1.0, para_gap_pt=0)
        flow.y = y + note_h
    return flow.slide


# ---- 기존 렌더러 어댑터 ----

def _clean_item(it):
    """어댑터가 넘기는 항목의 None 값을 빈 문자열로(렌더러가 "None"을 그리지 않도록)."""
    d = dict(_item_dict(it))
    for k, v in list(d.items()):
        if v is None:
            d[k] = ""
    return d


def _blk_background(flow, sec, ctx):
    return build_pc_background(flow, {"title": _s(sec.get("title")), "content": _s(sec.get("content"))})


def _blk_points(flow, sec, ctx):
    items = [_clean_item(i) for i in _items(sec, "reasons")]
    for it in items:
        if "content" not in it and it.get("description"):
            it["content"] = it["description"]
    return build_pc_reasons(flow, items, product_name=ctx.get("product_name", ""),
                            heading=_s(sec.get("title")) or None)


def _blk_course(flow, sec, ctx):
    items = []
    for i in _items(sec, "destinations", "courses", "spots"):
        d = _clean_item(i)
        if "region_tag" not in d and d.get("tag"):
            d["region_tag"] = d["tag"]
        if "description" not in d and d.get("content"):
            d["description"] = d["content"]
        items.append(d)
    return build_pc_destinations(flow, items, section_title=_s(sec.get("title")) or None,
                                 icon=sec.get("icon") or "route", default_title="코스 하이라이트")


def _blk_route_compare(flow, sec, ctx):
    routes = sec.get("routes") or sec.get("items") or []
    return build_pc_route_compare(flow, {"title": _s(sec.get("title")), "routes": routes})


def _blk_spec_table(flow, sec, ctx):
    specs = []
    for i in _items(sec, "specs", "rows"):
        if isinstance(i, dict):
            specs.append({"label": _first(i, "label", "title", "name"), "value": _first(i, "value", "description")})
        elif isinstance(i, (list, tuple)) and len(i) >= 2:
            specs.append({"label": _s(i[0]), "value": _s(i[1])})
    with_image = sec.get("image_count") not in (0, "0", False)
    return build_pc_transport(flow, {"title": _s(sec.get("title")) or "스펙 한눈에 보기", "specs": specs},
                              image_label="이미지", with_image=with_image, icon=sec.get("icon") or "list-checks")


def _blk_service_cards(flow, sec, ctx):
    items = []
    for i in _items(sec, "points"):
        items.append(_clean_item(i))
    return build_pc_experience(flow, _s(sec.get("tagline")), items,
                               heading=_s(sec.get("title")) or "혜초와 함께라면 더 특별한 이유")


def _blk_guides(flow, sec, ctx):
    items = []
    for i in _items(sec, "guides"):
        d = _clean_item(i)
        if "name" not in d and d.get("title"):
            d["name"] = d.pop("title")
        if "bio" not in d and d.get("description"):
            d["bio"] = d["description"]
        items.append(d)
    return build_pc_guide(flow, items, heading=_s(sec.get("title")) or "함께하는 사람들")


def _blk_season(flow, sec, ctx):
    table = sec.get("table") or sec.get("season_table") or []
    season = {"title": _s(sec.get("title")) or "언제 가면 좋을까?", "content": _s(sec.get("content")),
              "stat_line": _s(sec.get("stat_line"))}
    return build_pc_season(flow, season, table)


def _blk_qna(flow, sec, ctx):
    pairs = []
    for i in _items(sec, "qna", "faq"):
        d = _item_dict(i)
        q = _first(d, "question", "title")
        a = _first(d, "answer", "description", "content")
        if q:
            pairs.append((d.get("icon") if d.get("icon") in ICON_KEYS else (pick_icon(q, a, None)), {"question": q, "answer": a}))
    return build_pc_qna(flow, pairs, heading=_s(sec.get("title")) or "궁금증을 풀어드려요")


_BLOCK_ALIASES = {
    "background": "background", "story": "background", "intro": "background",
    "points": "points", "reasons": "points", "why": "points", "why_reasons": "points",
    "course": "course", "courses": "course", "destinations": "course", "spots": "course",
    "tour_spots": "course", "trekking": "course",
    "route_compare": "route_compare", "compare": "route_compare",
    "spec_table": "spec_table", "specs": "spec_table", "transport": "spec_table", "transport_spec": "spec_table",
    "service_cards": "service_cards", "services": "service_cards", "service": "service_cards",
    "experience": "service_cards", "experience_points": "service_cards",
    "guides": "guides", "guide": "guides", "guide_profile": "guides",
    "season": "season", "qna": "qna", "faq": "qna",
    "overview": "overview", "summary": "overview", "facts": "overview",
    "audience": "audience", "recommend": "audience",
    "highlights": "highlights", "highlight": "highlights",
    "cards": "cards", "gallery": "cards",
    "schedule": "schedule", "days": "schedule", "timeline": "schedule", "itinerary": "schedule",
    "closing": "closing", "cta": "closing",
    "text": "text", "paragraph": "text",
    "list": "list", "checklist": "list", "bullets": "list",
    "table": "table",
}

_BLOCK_RENDERERS = {
    "background": _blk_background, "points": _blk_points, "course": _blk_course,
    "route_compare": _blk_route_compare, "spec_table": _blk_spec_table,
    "service_cards": _blk_service_cards, "guides": _blk_guides, "season": _blk_season, "qna": _blk_qna,
    "overview": lambda f, s, c: build_pc_overview(f, s),
    "audience": lambda f, s, c: build_pc_audience(f, s),
    "highlights": lambda f, s, c: build_pc_highlights(f, s),
    "cards": lambda f, s, c: build_pc_cards(f, s),
    "schedule": lambda f, s, c: build_pc_schedule(f, s),
    "closing": lambda f, s, c: build_pc_closing(f, s),
    "text": lambda f, s, c: build_pc_text(f, s),
    "list": lambda f, s, c: build_pc_list(f, s),
    "table": lambda f, s, c: build_pc_table_block(f, s),
}

BLOCK_LABELS = {
    "background": "배경 이야기", "points": "포인트(차별점)", "course": "코스/방문지", "route_compare": "노선 비교",
    "spec_table": "스펙 표", "service_cards": "혜초 서비스 카드", "guides": "가이드 소개",
    "season": "시즌 안내", "qna": "Q&A", "overview": "개요(핵심 정보)", "audience": "추천 대상",
    "highlights": "하이라이트", "cards": "사진 카드", "schedule": "일정 한눈에", "closing": "마무리 문구",
    "text": "텍스트", "list": "체크 리스트", "table": "표",
}


def normalize_block_type(t):
    key = re.sub(r"[\s\-]+", "_", _s(t).lower())
    return _BLOCK_ALIASES.get(key)


def _build_generic_section(flow, sec):
    """알 수 없는 블록 종류: 제목 + (본문 | 항목 리스트)로 최대한 그대로 보여준다."""
    items = _items(sec)
    if items:
        return build_pc_list(flow, {"title": sec.get("title") or sec.get("type") or "안내", "items": items})
    return build_pc_text(flow, sec)


def build_pc_sections(flow, sections, ctx):
    """sections를 순서대로 그린다. 한 블록이 오류를 내도 나머지는 계속 그리고, 실패한 블록은
    일반 텍스트/리스트로 한 번 더 시도한다. 반환값: 그린 블록 수."""
    drawn = 0
    for sec in sections:
        if not isinstance(sec, dict):
            continue
        kind = normalize_block_type(sec.get("type"))
        fn = _BLOCK_RENDERERS.get(kind)
        try:
            res = fn(flow, sec, ctx) if fn else _build_generic_section(flow, sec)
        except Exception:
            try:
                res = _build_generic_section(flow, sec)
            except Exception:
                res = None
        if res:
            drawn += 1
    return drawn


_AI_MARKUP_RE = re.compile(r"(?:</?|\()\s*cite\b[^>\n]*>|</\s*cite\s*>", re.I)


def strip_ai_markup(obj):
    """웹 검색 결과를 쓴 AI 응답에 섞여 나오는 인용 태그(<cite index="...">...</cite>)를 제거하고
    안쪽 글만 남긴다. 딕셔너리/리스트 안의 모든 문자열에 적용(원본은 건드리지 않고 사본을 돌려줌)."""
    if isinstance(obj, str):
        return _AI_MARKUP_RE.sub("", obj)
    if isinstance(obj, list):
        return [strip_ai_markup(x) for x in obj]
    if isinstance(obj, dict):
        return {k: strip_ai_markup(v) for k, v in obj.items()}
    return obj


def build_legacy(content_json, out_path, review=None):
    """content_json(정현지 스키마) -> 새 PPTX 파일 생성.
    review(reviewer.review_content()의 반환값)를 넘기면 마지막에 검수 결과 페이지를
    덧붙인다 — 넘기지 않으면(기본값 None) 예전과 동일하게 검수 페이지 없이 생성된다."""
    content_json = strip_ai_markup(content_json)
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
    # tour_spots: 트레킹 상품 중 트레킹 코스 + 관광(도시/유적/박물관) 코스가 함께 있는
    # 경우에만 채워진다 — 카라코람 테스트에서 훈자 마을/이슬라마바드/탁실라 박물관 같은
    # "관광" 코스가 트레킹 코스(destinations)와 구분 없이 한 섹션에 섞여 나온 문제를
    # 고치기 위해 별도 섹션으로 분리했다. build_destination_slides와 카드 레이아웃은
    # 동일하고 헤딩만 다르다.
    build_destination_slides(
        flow,
        content_json.get("tour_spots", []),
        section_title=content_json.get("tour_spots_heading"),
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
    # highlights/highlights_heading 필드는 제거됨 — why_reasons("포인트 0N" 섹션)와
    # 지시문이 실질적으로 같은 "상품 차별점/테마 하이라이트" 내용을 요구해, AI가 같은
    # 내용을 문구만 바꿔 두 번 반복하는 문제가 반복 발생했다(예: 멕시코 문명기행
    # 테스트에서 "칸쿤 없이 완성하는 내륙 문명 루트" 등 4개 항목이 포인트 섹션과
    # 하이라이트 섹션에 그대로 중복). 스키마에서 highlights를 없애 why_reasons
    # 하나로 통일한다(build_highlights_slides 함수 자체는 향후 필요시를 위해 남겨둠).
    build_season_slide(flow, content_json.get("season", {}), content_json.get("season_table"))
    build_meal_slide(flow, content_json.get("meal_info"))
    build_safety_slide(flow, content_json.get("altitude_profile"), content_json.get("safety_note"))
    build_banner_request_slide(flow, cover, content_json.get("banner_copy"))
    build_review_notes_slide(flow, review)
    prs.save(out_path)
    return prs


def build(content_json, out_path, review=None):
    """PC 상품소개 기획안 생성(정돈형 레이아웃, 2026-11). 예전 방식은 build_legacy()에 남겨둠.
    content_json 스키마는 그대로 — icon/image_count는 선택 필드(없으면 자동 추정/기본 2장)."""
    content_json = strip_ai_markup(content_json)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    cover = content_json.get("cover", {})
    flow = PCFlow(prs, footer_label=" ".join(str(cover.get("product_name", "")).split()))

    build_pc_cover(flow, cover, content_json.get("watermark_label", ""))
    sections = content_json.get("sections")
    if isinstance(sections, list) and sections:
        # 상품별 동적 구성: AI가 고른 블록 순서대로 본문을 그린다.
        ctx = {"product_name": cover.get("product_name", "")}
        drawn = build_pc_sections(flow, sections, ctx)
        if drawn:
            build_banner_request_slide(flow, cover, content_json.get("banner_copy"), bar_color=PC_NAVY)
            build_pc_review_notes(flow, review)
            prs.save(out_path)
            return prs
    # sections가 없거나 하나도 못 그렸으면 예전 고정 구성으로 그린다(하위 호환)
    build_pc_background(flow, content_json.get("background_story"))
    build_pc_reasons(flow, content_json.get("why_reasons"), product_name=cover.get("product_name", ""))
    build_pc_destinations(flow, content_json.get("destinations", []),
                          section_title=content_json.get("destinations_heading"),
                          icon="route", default_title="코스 하이라이트")
    build_pc_destinations(flow, content_json.get("tour_spots", []),
                          section_title=content_json.get("tour_spots_heading"),
                          icon="landmark", default_title="함께 즐기는 관광 코스")
    build_pc_route_compare(flow, content_json.get("route_compare"))
    build_pc_transport(flow, content_json.get("transport_spec"))
    build_pc_experience(flow, content_json.get("brand_tagline", ""), content_json.get("experience_points"))
    build_pc_guide(flow, content_json.get("guide_profile"))
    build_pc_season(flow, content_json.get("season", {}), content_json.get("season_table"))
    build_pc_qna(flow, [("utensils", content_json.get("meal_info")),
                        ("shield-check", content_json.get("safety_note"))])
    build_banner_request_slide(flow, cover, content_json.get("banner_copy"), bar_color=PC_NAVY)
    build_pc_review_notes(flow, review)
    prs.save(out_path)
    return prs



def build_mobile(content_json, out_path, review=None):
    """모바일 최적화 기획안 빌더. content_json은 build()과 완전히 같은 "정현지
    스키마"를 그대로 받는다 — 새 스키마를 만들지 않은 이유는, 실제 회사 가이드라인
    (모바일 최적화 콘텐츠 가이드라인_20260702.pdf)과 실제 완성 Before/After 샘플
    (17.모바일최적화 폴더, 돌로미테·그랜드서클 등)을 대조해본 결과 모바일 버전은
    PC 버전과 텍스트 내용이 100% 동일하고 폰트 크기·여백·배지(태그) 스타일·배경색만
    달랐기 때문이다 — "축약/요약"이 아니라 "재배치". 그래서 콘텐츠 스키마는
    재사용하고, 시각 표현만 build_mobile_* 함수들로 갈아끼운다.

    1차 버전 범위: 실제 샘플에서 PC 대비 가장 뚜렷하게 스타일이 바뀌어 있던 4개
    섹션(표지, 특별함 카드, 방문지/투어스팟, Q&A)만 모바일 전용 함수로 새로 그린다.
    나머지 섹션(노선비교/교통정보/가이드소개 등)은 실제 샘플에 대응 사례가 없어
    확인이 안 되므로, 1차 버전에서는 PC용 함수를 그대로 재사용한다(MobileSlideFlow가
    깔아주는 배경 위에 얹히긴 하지만, 폰트 크기는 PC 그대로 나온다는 뜻 — 필요하면
    다음 라운드에서 같은 패턴으로 확장 가능)."""
    content_json = strip_ai_markup(content_json)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    flow = MobileSlideFlow(prs)

    cover = content_json.get("cover", {})
    build_mobile_cover_slide(flow, cover, content_json.get("watermark_label", ""))
    build_background_slide(flow, content_json.get("background_story"))
    build_reasons_slide(flow, content_json.get("why_reasons"), product_name=cover.get("product_name", ""))
    build_mobile_destination_slides(
        flow,
        content_json.get("destinations", []),
        section_title=content_json.get("destinations_heading"),
    )
    build_mobile_destination_slides(
        flow,
        content_json.get("tour_spots", []),
        section_title=content_json.get("tour_spots_heading"),
    )
    build_route_compare_slide(flow, content_json.get("route_compare"))
    build_transport_slide(flow, content_json.get("transport_spec"))
    build_mobile_experience_slide(
        flow,
        content_json.get("brand_tagline", ""),
        content_json.get("experience_points"),
    )
    build_guide_slide(flow, content_json.get("guide_profile"))
    build_season_slide(flow, content_json.get("season", {}), content_json.get("season_table"))
    # altitude_profile(고도 프로필)은 safety_note 없이 build_safety_slide로 그대로
    # 그리고, meal_info/safety_note의 question/answer는 build_mobile_qna_slide로
    # 묶어서 카드 리스트로 재배치한다(실제 샘플의 "궁금증을 풀어드려요" 섹션과 동일한 구조).
    build_safety_slide(flow, content_json.get("altitude_profile"), None)
    qna_pairs = [content_json.get("meal_info"), content_json.get("safety_note")]
    build_mobile_qna_slide(flow, qna_pairs)
    build_banner_request_slide(flow, cover, content_json.get("banner_copy"))
    build_review_notes_slide(flow, review)
    prs.save(out_path)
    return prs
