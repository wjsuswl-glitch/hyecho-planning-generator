"""사업부 원본자료(docx) 파싱 모듈 — 2단계 계층 구조 지원"""
import re
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
