"""사업부 원본자료(docx/pptx/이미지) 파싱 모듈 — 2단계 계층 구조 지원"""
import re
import io
import base64
import mimetypes
from PIL import Image, ImageOps
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
TOP_MARKER_RE = re.compile(r"^(\*|●|chapter\s*\d+|[0-9]{2}\s)", re.IGNORECASE)
SUB_NUMBER_RE = re.compile(r"^[0-9]\.\s")
def iter_block_items(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)
def parse_docx(path):
    doc = Document(path)
    items = list(iter_block_items(doc))
    has_strict_top = any(
        TOP_MARKER_RE.match(it.text.strip()) and len(it.text.strip()) < 40
        for it in items if isinstance(it, Paragraph) and it.text.strip()
    )
    use_subnumber_as_top = not has_strict_top
    sections = {}
    current_key = "header"
    sections[current_key] = []
    current_sub = None
    def add_content(content):
        nonlocal current_sub
        if current_sub is not None:
            current_sub["items"].append(content)
        else:
            sections[current_key].append(content)
    for it in items:
        if isinstance(it, Paragraph):
            t = it.text.strip()
            if not t:
                continue
            style_name = it.style.name if it.style else ""
            is_heading2 = style_name.startswith("Heading 2")
            if TOP_MARKER_RE.match(t) and len(t) < 40:
                current_key, current_sub = t, None
                sections[current_key] = []
            elif SUB_NUMBER_RE.match(t) and len(t) < 40 and use_subnumber_as_top:
                current_key, current_sub = t, None
                sections[current_key] = []
            elif (SUB_NUMBER_RE.match(t) and len(t) < 40) or (is_heading2 and len(t) < 60):
                current_sub = {"subheading": t, "items": []}
                sections[current_key].append(current_sub)
            else:
                add_content(t)
        else:
            rows = [[c.text.strip().replace("\n", " ") for c in row.cells] for row in it.rows]
            add_content({"table": rows})
    return sections
def detect_format_and_draft_copy(sections):
    """유형 판별 + draft_copy(카피 초안) 존재 여부 감지"""
    keys = list(sections.keys())
    has_numbered = any(re.match(r"^[0-9]{2}\s", k) for k in keys)
    has_chapter = any(re.match(r"^chapter", k, re.IGNORECASE) for k in keys)
    has_star = any(k.startswith("*") for k in keys)
    has_dot = any(k.startswith("●") for k in keys)
    if has_numbered or has_chapter:
        fmt = "B"
    elif has_dot:
        fmt = "C"
    elif has_star:
        fmt = "A"
    elif len(keys) <= 1:
        fmt = "D"
    else:
        fmt = "E"
    draft_copy = None
    if fmt == "B":
        for k, v in sections.items():
            if "디자인팀" in k or "카피라이팅" in k or "가이드" in k:
                for block in v:
                    if isinstance(block, dict) and "table" in block:
                        for row in block["table"]:
                            joined = " ".join(row)
                            if "카피" in joined or len(joined) > 20:
                                draft_copy = joined
                                break
    return {"format_type": fmt, "draft_copy": draft_copy}


def parse_pptx(path):
    """PPTX 자료(예: 사업부에서 준 예전 상품 소개 PPT)를 슬라이드별 텍스트/표로 추출.
    docx의 '*'/'●'/번호 마커 같은 계층 구조는 PPT엔 없으므로, 슬라이드 번호를 그대로
    키로 써서 sections와 같은 형태(dict)로 반환한다 — prompt_builder는 dict 구조면
    무엇이든 그대로 JSON화해서 프롬프트에 넣으므로 별도 변환이 필요 없다."""
    from pptx import Presentation
    prs = Presentation(path)
    result = {}
    for i, slide in enumerate(prs.slides, 1):
        items = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                items.append(shape.text_frame.text.strip())
            if getattr(shape, "has_table", False):
                rows = [[c.text.strip().replace("\n", " ") for c in row.cells]
                        for row in shape.table.rows]
                items.append({"table": rows})
        if items:
            result[f"슬라이드 {i}"] = items
    return result


def _to_jpeg_block(img, jpeg_quality):
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=jpeg_quality)
    data = base64.b64encode(buf.getvalue()).decode("utf-8")
    return {
        "type": "image",
        "source": {"type": "base64", "media_type": "image/jpeg", "data": data},
    }


def encode_image_blocks(path, max_side=1568, jpeg_quality=90):
    """이미지 파일 1개를 Claude API 멀티모달 블록 '목록'(리스트, 1개 이상)으로 변환.

    [2026-09-15 수정 배경] 예전엔 encode_image_block(단수)에서 긴 변이 2000px을
    넘으면 그 긴 변 기준으로 통째로 축소했다. 그런데 모바일 최적화 기능 실사용
    테스트(미서부 대장정, 유럽 알프스 3대 미봉 상품)에서 실제로 항목이 통째로
    누락되는 문제가 발견됐고, 원인을 추적해보니 이 축소 방식이 범인이었다 —
    사업부 상품소개 이미지는 보통 가로 1150px 안팎에 세로만 아주 긴(예:
    12,000px) "세로 스크롤 카드뉴스" 형태인데, 세로(긴 변) 기준으로 1568px까지
    줄이면 비율 유지 때문에 가로가 150px 밑으로 줄어버려 글자가 다 뭉개진다.
    가로가 원래 읽기 좋았던 폭인데 세로가 길다고 가로까지 함께 줄여버리는 게
    문제였던 것 — 그 결과 이미지 아래쪽에 있던 내용(예: "몬테로사" 구간
    전체)을 AI가 아예 읽지 못하고 건너뛰었다.

    게다가 Claude API 쪽에서도 이미지를 표준 해상도 등급 기준 긴 변
    1568px(Sonnet 계열 기준)로 자체 축소해서 처리하므로, 애초에 세로로 아주 긴
    이미지를 한 장으로 보내면 이 API 자체 축소에서도 똑같이 가로가 뭉개진다 —
    그래서 축소가 아니라 "분할"이 정답이다. 세로가 max_side를 넘으면 가로/원본
    해상도는 그대로 유지한 채 세로 방향으로 max_side 높이씩 여러 장으로 잘라
    각각을 별도 이미지 블록으로 만든다 — Anthropic 공식 가이드도 긴 이미지는
    축소보다 "여러 장의 이미지로 순서대로 전달"하는 방식을 권장한다.

    가로가 max_side보다 큰 경우(드묾)는 먼저 가로 기준으로 비율 유지 축소한
    뒤 위 분할 로직을 적용한다. 항상 JPEG로 재인코딩해서 media_type 불일치나
    팔레트/투명 채널(RGBA, PNG 등) 문제도 없앤다. 휴대폰으로 찍은 사진의 EXIF
    방향 정보도 여기서 반영해 회전 문제를 방지한다."""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    width, height = img.size
    if width > max_side:
        scale = max_side / width
        width, height = max(1, round(width * scale)), max(1, round(height * scale))
        img = img.resize((width, height), Image.LANCZOS)

    if height <= max_side:
        pieces = [img]
    else:
        n_slices = -(-height // max_side)  # ceil — 조각 개수
        slice_h = -(-height // n_slices)   # 조각당 높이(마지막 조각만 더 짧을 수 있음)
        pieces = [img.crop((0, top, width, min(top + slice_h, height)))
                  for top in range(0, height, slice_h)]

    return [_to_jpeg_block(piece, jpeg_quality) for piece in pieces]


def encode_image_block(path, max_dimension=2000, jpeg_quality=85):
    """하위 호환용 — 이미지 1장을 블록 1개로만 반환(분할 없이 긴 변 기준 축소).
    세로로 긴 카드뉴스형 상품소개 이미지에는 encode_image_blocks(복수형)를
    쓸 것 — 이 함수처럼 긴 변 기준으로 축소하면 세로가 아주 긴 이미지에서
    가로(글자가 실제로 놓인 폭)가 함께 뭉개지는 문제가 있다(위 encode_image_blocks
    설명 참고). 이 함수는 다른 코드에서 참조할 경우를 대비해서만 남겨둔다."""
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    width, height = img.size
    longest_side = max(width, height)
    if longest_side > max_dimension:
        scale = max_dimension / longest_side
        new_size = (max(1, round(width * scale)), max(1, round(height * scale)))
        img = img.resize(new_size, Image.LANCZOS)

    return _to_jpeg_block(img, jpeg_quality)
