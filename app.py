"""혜초여행 기획안 자동생성 웹앱 (프로토타입)"""
import streamlit as st
import tempfile, os, json, sys, unicodedata, html

sys.path.insert(0, os.path.dirname(__file__))
import io
from parser import (parse_docx, detect_format_and_draft_copy, parse_pptx, encode_image_blocks,
                    extract_docx_images)
from prompt_builder import build_system_prompt, build_mobile_system_prompt
from generator import generate_content
from reviewer import review_content
from assembler import assemble
import builder as dynamic_builder
import assembly_presets as ap
import hashlib
try:
    from streamlit_sortables import sort_items  # 블록을 마우스로 끌어 옮기는 부품
except Exception:  # 설치 안 됐으면 ↑↓ 버튼만 사용
    sort_items = None

# 블록 카드를 끌어 옮기고, 카드를 누르면 그 자리에서 제목·요청·위/아래/삭제가 펼쳐지는 부품(asm_component 폴더).
# 폴더가 없거나 불러오지 못하면 예전 방식(끌어서 순서 + 위쪽 도구 모음)으로 자동 전환된다.
try:
    import streamlit.components.v1 as _components
    _comp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "asm_component")
    _ASM_COMP = (_components.declare_component("asm_blocks", path=_comp_dir)
                 if os.path.isfile(os.path.join(_comp_dir, "index.html")) else None)
except Exception:
    _ASM_COMP = None

TEMPLATE_MAP_PATH = os.path.join(os.path.dirname(__file__), "template_map.json")


def build_multimodal_image_blocks(image_files):
    """업로드된 이미지 파일들을 Claude API에 보낼 이미지 블록 목록으로 변환한다.

    encode_image_blocks(parser.py)가 세로로 아주 긴 이미지 1장을 여러 조각으로
    나눠서 반환할 수 있으므로(2026-09-15 수정 — "몬테로사" 구간 전체가 통째로
    누락되는 문제의 원인이었음, parser.py의 encode_image_blocks 설명 참고), 한
    파일이 여러 조각으로 나뉘면 그 앞에 "이 조각들은 원래 한 이미지가 이어지는
    부분"이라는 걸 AI가 알 수 있도록 안내 텍스트 블록을 하나 끼워 넣는다 —
    Anthropic이 권장하는 "여러 이미지에 순서 라벨을 붙여 전달" 방식이다."""
    blocks = []
    for idx, (name, path) in enumerate(image_files, 1):
        slices = encode_image_blocks(path)
        if len(slices) > 1:
            blocks.append({
                "type": "text",
                "text": f"[첨부 이미지 {idx}: {name} — 세로로 길어서 위에서 아래로 "
                        f"{len(slices)}개 조각으로 나눠 전달합니다. 아래 이미지 "
                        f"{len(slices)}장은 서로 잘린 게 아니라 이어지는 한 페이지이니, "
                        f"순서대로 이어서 하나의 내용으로 읽어주세요.]",
            })
        blocks.extend(slices)
    return blocks


def build_docx_embedded_image_blocks(docx_files):
    """사업부 워드(.docx) 안에 삽입돼 있는 이미지(고도표/지도/일정표 캡처 등)를 AI에게
    보낼 이미지 블록으로 만든다. 2026-10-27 '돌로미테 하이라이트 Part.2' 비교 테스트에서,
    워드 안 고도표 이미지에만 적혀 있던 거리/소요시간/최고고도/상행·하행 수치가 통째로
    누락된 것이 원인 확인됐다(parse_docx는 텍스트만 읽음). 각 이미지 앞에는 몇 번째
    이미지이고 워드 어느 문단 뒤에 있었는지를 알려주는 라벨 텍스트를 붙인다.
    반환: (블록 목록, 꺼낸 이미지 개수)"""
    blocks, total = [], 0
    for name, path in docx_files:
        try:
            imgs = extract_docx_images(path)
        except Exception:
            continue  # 이미지 추출 실패가 전체 생성을 막으면 안 된다 — 텍스트만으로 계속 진행
        for img in imgs:
            total += 1
            ctx = f" (워드에서 '{img['context']}' 문단 뒤에 위치)" if img["context"] else ""
            blocks.append({
                "type": "text",
                "text": f"[워드 삽입 이미지 {total}: {name} 안의 {img['index']}번째 이미지{ctx}. "
                        f"사업부가 본문에 직접 넣은 자료(고도표/지도/표 등)이므로 이미지 속 글자와 "
                        f"수치도 텍스트 자료와 똑같이 원본자료로 읽어 반영하세요.]",
            })
            blocks.extend(encode_image_blocks(io.BytesIO(img["blob"])))
    return blocks, total

st.set_page_config(page_title="혜초 기획안 자동생성", page_icon="🧳")
st.title("🧳 혜초여행 기획안 자동생성")
st.caption("사업부 원본자료(docx/pptx/이미지, 최대 5개) → AI 카피 생성 → 기획안 PPTX")

with open(TEMPLATE_MAP_PATH, encoding="utf-8") as f:
    # NFC로 정규화 — template_map.json이 macOS 등에서 자모 분리형(NFD)으로 저장돼
    # 있으면, 화면엔 "최정인"으로 똑같이 보여도 prompt_builder.py에 NFC로 적힌
    # STYLE_RULES 등 딕셔너리 키와 바이트 단위로 달라 KeyError가 난다("최정인"
    # 추가 후 실제로 발생한 오류). 여기서 한 번 정규화해두면 이 값을 넘겨받는
    # 모든 곳(build_system_prompt, writer_style in (...) 비교 등)이 안전해진다.
    available_styles = [unicodedata.normalize("NFC", k) for k in json.load(f).keys()]

mode = st.radio(
    "작업 모드",
    ["신규 상품소개 제작", "모바일 최적화 기획안"],
    horizontal=True,
    help="신규 상품소개 제작: 사업부 원본자료(워드 등)로부터 새 기획안을 만듭니다. "
         "모바일 최적화 기획안: 이미 있는 PC용 상품소개 이미지를 읽어, 항목(트레일/코스/"
         "스팟 등)은 하나도 빠짐없이 유지하면서 각 문장만 모바일에 맞게 간결히 정리하고, "
         "폰트/여백/배지 스타일도 모바일 가이드라인에 맞게 다시 배치합니다.",
)
is_mobile = mode == "모바일 최적화 기획안"

# 카피를 쓰는 AI 선택(기본 Gemini). 검수는 항상 반대편 AI가 맡는다 — 같은 AI가 자기 글을
# 검수하면 왜곡/날조를 놓치기 쉬워서.
gen_ai = st.selectbox(
    "카피 생성 AI", ["Gemini", "Claude"], index=0,
    help="기획안 문구를 쓰는 AI입니다. 검수는 반대편 AI(Gemini로 쓰면 Claude, Claude로 쓰면 Gemini)가 맡아 "
         "교차 검수합니다. 결과가 마음에 안 들면 다른 AI로 바꿔서 다시 생성해 비교해보세요.",
)
gen_provider = "gemini" if gen_ai == "Gemini" else "claude"
review_provider = "claude" if gen_provider == "gemini" else "gemini"
review_ai = "Claude" if review_provider == "claude" else "Gemini"

if is_mobile:
    # 모바일 모드는 특정 기획자 문체를 입히지 않고 원본 이미지의 텍스트를 그대로
    # 옮기는 작업이라(build_mobile_system_prompt 참고) 기획자 스타일 선택이 필요
    # 없다 — writer_style은 파일명 등에서만 쓰이는 고정 라벨로 둔다.
    writer_style = "모바일최적화"
    category = unicodedata.normalize("NFC", st.selectbox("카테고리", ["문탐", "트레킹"]))
    enable_web_search = False  # 새 사실을 검색하는 게 아니라 기존 이미지 내용만 옮기므로 불필요
else:
    col1, col2 = st.columns(2)
    with col1:
        writer_style = unicodedata.normalize("NFC", st.selectbox("기획자 스타일", available_styles))
    with col2:
        category = unicodedata.normalize("NFC", st.selectbox("카테고리", ["문탐", "트레킹"]))

    enable_web_search = st.checkbox(
        "사업부 자료에 없는 배경지식/사실을 웹 검색으로 보완",
        value=True,
        help="background_story 등에서 사업부 자료에 없는 사실이 필요할 때 AI가 웹 검색으로 확인합니다. "
             "끄면 사업부 자료와 일반 상식 범위 내에서만 작성합니다(검색 비용/시간 절약).",
    )

# 기획자 요청사항 — 사업부 자료 외에 기획자가 이번 기획안에 바라는 점을 화면에서 직접 적는다.
# (2026-10-27 추가) 새 카피를 만드는 "신규 상품소개 제작" 모드에서만 의미가 있다 — 모바일
# 최적화는 기존 이미지 내용을 그대로 옮기는 작업이라 톤앤매너/강조 지시가 필요 없다.
planner_request = {}
if not is_mobile:
    with st.expander("✍️ 기획자 요청사항 (선택) — 톤앤매너, 강조할 점, 빼야 할 점", expanded=False):
        st.caption("비워두면 사업부 자료와 선택한 기획자 스타일만으로 만듭니다. "
                   "적은 내용은 스타일 기본값보다 우선 반영되지만, 사업부 자료에 없는 사실을 "
                   "만들거나 브랜드 금지 표현(예: 트레킹 '완전정복')을 쓰는 것까지 허용하진 않아요.")
        planner_request = {
            "tone": st.text_area(
                "전체 톤앤매너", height=80,
                placeholder="예) 차분하고 여유로운 느낌. 설명은 부드러운 ~해요체로."),
            "emphasize": st.text_area(
                "강조하면 좋은 부분", height=80,
                placeholder="예) 겹치지 않는 코스(재방문 고객 대상), 호텔 연박으로 누리는 휴식."),
            "avoid": st.text_area(
                "넣으면 안 되는 부분 / 피할 표현", height=80,
                placeholder="예) 난이도가 어렵다는 느낌을 주는 표현, 타사 언급."),
            "extra": st.text_area(
                "기타 요청사항", height=80,
                placeholder="예) 표지 문구는 짧게, 숫자를 앞세워 주세요."),
        }


# ---- 구성 조립 화면 (PC 상품소개 기획안) -----------------------------------------------------
def _asm_bump():
    st.session_state.asm_ver = st.session_state.get("asm_ver", 0) + 1


def _asm_sel_index():
    ss = st.session_state
    for i, b in enumerate(ss.assembly):
        if b["id"] == ss.get("asm_selected"):
            return i
    return -1


def _asm_move_sel(d):
    """선택한 블록을 한 칸 위(-1)/아래(+1)로."""
    blocks = st.session_state.assembly
    i = _asm_sel_index()
    j = i + d
    if i >= 0 and 0 <= j < len(blocks):
        blocks[i], blocks[j] = blocks[j], blocks[i]
        _asm_bump()


def _asm_delete_sel():
    """선택한 블록을 지우고, 같은 자리의 다음 블록(없으면 마지막)을 선택 상태로 둔다."""
    ss = st.session_state
    i = _asm_sel_index()
    if i >= 0:
        ss.assembly.pop(i)
        if ss.assembly:
            ss.asm_selected = ss.assembly[min(i, len(ss.assembly) - 1)]["id"]
        _asm_bump()


def _asm_add(block_type):
    nb = ap.new_block(block_type)
    st.session_state.assembly.append(nb)
    st.session_state.asm_focus = nb["id"]  # 새 블록은 펼쳐진 채로 보여준다(제목·요청을 바로 쓸 수 있게)
    _asm_bump()


def _asm_load_preset(name):
    st.session_state.assembly = ap.preset_blocks(name)
    _asm_bump()


def _asm_load_saved(name):
    st.session_state.assembly = ap.blocks_from_saved(ap.load_saved().get(name))
    _asm_bump()


def render_assembly(category):
    """기획자가 상품에 맞게 블록을 골라 순서대로 조립하는 화면.
    반환: 조립 결과 [{"type","title","memo"}...] (AI에게 맡기면 None)."""
    ss = st.session_state
    if "assembly" not in ss:
        ss.assembly = ap.preset_blocks(ap.DEFAULT_PRESET_BY_CATEGORY.get(category, "트레킹"))

    mode_pick = st.radio(
        "구성 방식", ["직접 조립", "AI에게 맡기기"], horizontal=True, key="asm_mode",
        help="직접 조립: 아래에서 고른 블록·순서·제목 그대로 기획안을 만듭니다. "
             "AI에게 맡기기: AI가 사업부 자료를 읽고 이 상품에 맞는 구성을 스스로 정합니다.",
    )
    if mode_pick == "AI에게 맡기기":
        st.caption("AI가 자료를 읽고 구성을 정합니다. 생성 후 '이 상품의 구성과 이유'에서 확인할 수 있어요.")
        return None

    with st.expander("🧩 구성 조립 — 이 상품에 필요한 블록만 골라 순서를 정하세요", expanded=True):
        c1, c2 = st.columns([3, 1])
        preset = c1.selectbox("상품 유형 프리셋", list(ap.PRESETS.keys()), key="asm_preset",
                              help="유형을 고르고 '불러오기'를 누르면 기본 조립이 채워지고, 이후 자유롭게 고칠 수 있어요.")
        c2.write("")
        c2.write("")
        c2.button("불러오기", key="asm_load_preset", on_click=_asm_load_preset, args=(preset,))

        blocks = ss.assembly

        def _box(label, sub="", height=None):
            """고정 페이지(표지/배너) 카드 — 회색 점선 + 자물쇠, 이동·삭제 불가."""
            sub_html = (f'<div style="font-size:12.5px;opacity:.7;margin-top:2px">{html.escape(sub)}</div>') if sub else ""
            st.markdown(
                '<div style="display:flex;align-items:center;gap:14px;border:1px dashed rgba(128,128,128,.55);'
                'border-left:5px solid rgba(128,128,128,.55);background:rgba(128,128,128,.09);border-radius:10px;'
                'padding:12px 18px;margin-bottom:8px;">'
                '<span style="font-size:18px;opacity:.8">🔒</span>'
                f'<div><div style="font-size:15px;font-weight:700;opacity:.75">{html.escape(label)}</div>{sub_html}</div></div>',
                unsafe_allow_html=True)

        def _block_text(b):
            """구성도 카드 문구: 첫 줄(굵게) = 블록 이름, 둘째 줄 = 설명 또는 지정한 제목."""
            t = (b.get("title") or "").strip()
            emoji = ap.BLOCK_EMOJI.get(b["type"], "▫️")
            first = f"{emoji}  {ap.BLOCK_LABEL.get(b['type'], b['type'])}"
            sub = f"제목 · {t}" if t else ap.BLOCK_HELP.get(b["type"], "")
            if (b.get("memo") or "").strip():
                sub += ("   |   " if sub else "") + "📝 요청 있음"
            return first + ("\n" + sub if sub else "")

        toolbar = st.container()  # 아래 구성도를 그린 뒤(드래그 반영 후) 채운다 — 번호가 항상 최신 순서와 맞도록
        if _ASM_COMP is not None:
            st.caption("기획안 한 장(위→아래)의 구성도예요. 카드를 끌어서 순서를 바꾸고, 카드를 누르면 그 자리에서 "
                       "제목·요청 쓰기와 ↑ ↓ 삭제를 할 수 있어요. 🔒 회색 카드는 자동으로 들어가는 고정 페이지예요.")
        else:
            st.caption("기획안 한 장(위→아래)의 구성도예요. 카드를 끌어서 순서를 바꾸거나, 위 도구 모음에서 블록을 골라 "
                       "↑ ↓ 삭제를 누르세요. 🔒 회색 카드는 자동으로 들어가는 고정 페이지예요.")
        _box("표지", "타이틀 · 상품요약 · 후보 카피 — 자동으로 들어가요")

        if not blocks:
            st.info("아직 쌓인 블록이 없어요. 위에서 프리셋을 불러오거나 아래 버튼으로 블록을 쌓아보세요.")
        elif _ASM_COMP is not None:
            payload = [{"id": b["id"], "type": b["type"], "title": b.get("title", ""), "memo": b.get("memo", ""),
                        "label": ap.BLOCK_LABEL.get(b["type"], b["type"]), "emoji": ap.BLOCK_EMOJI.get(b["type"], "▫️"),
                        "help": ap.BLOCK_HELP.get(b["type"], "")} for b in blocks]
            ver = ss.get("asm_ver", 0)
            if ss.get("asm_focus_ver") != ver:      # 구성이 새로 채워진 첫 화면에서만 '펼칠 블록'을 알려준다
                ss.asm_focus_ver = ver
                focus = ss.pop("asm_focus", None)
            else:
                focus = None
            res = _ASM_COMP(blocks=payload, ver=ver, selected=focus, key=f"asm_comp_{ver}", default=None)
            if isinstance(res, dict) and isinstance(res.get("blocks"), list):
                by_id = {b["id"]: b for b in blocks}
                new_blocks = []
                for r in res["blocks"]:
                    b = by_id.get(r.get("id")) if isinstance(r, dict) else None
                    if b is not None and b not in new_blocks:
                        b["title"] = str(r.get("title") or "")
                        b["memo"] = str(r.get("memo") or "")
                        new_blocks.append(b)
                ss.assembly = blocks = new_blocks
        elif sort_items is not None:
            labels, label_to_block, seen = [], {}, {}
            for b in blocks:
                base = _block_text(b)
                n = seen.get(base, 0)
                seen[base] = n + 1
                lab = base + "\u200b" * n  # 같은 이름이 여럿이어도 문자열이 겹치지 않게(눈에 안 보임)
                labels.append(lab)
                label_to_block[lab] = b
            sig = hashlib.md5("|".join(sorted(labels)).encode("utf-8")).hexdigest()[:8]
            result = sort_items(
                labels, direction="vertical", key=f"asm_sort_{ss.get('asm_ver', 0)}_{sig}",
                custom_style="""
                .sortable-component { background: transparent !important; padding: 0 !important; }
                .sortable-container { background: transparent !important; padding: 0 !important; border: 0 !important; }
                .sortable-container-body { background: transparent !important; padding: 0 !important; gap: 0 !important;
                    counter-reset: blk; }
                .sortable-item { counter-increment: blk; position: relative !important; display: block !important;
                    background: #ffffff !important; color: #6b7280 !important; text-align: left !important;
                    white-space: pre-line !important; font-size: 12.5px !important; font-weight: 400 !important;
                    line-height: 1.5 !important; height: auto !important; min-height: 66px !important;
                    padding: 13px 50px 13px 60px !important; margin: 0 0 8px 0 !important; box-sizing: border-box !important;
                    width: 100% !important; border: 1px solid #dde3ea !important; border-left: 5px solid #222A35 !important;
                    border-radius: 10px !important; box-shadow: 0 1px 3px rgba(20,30,50,.08) !important; cursor: grab !important;
                    transition: box-shadow .15s, border-color .15s, background .15s; }
                .sortable-item::first-line { font-size: 16px; font-weight: 700; color: #1f2933; }
                .sortable-item::before { content: counter(blk); position: absolute; left: 16px; top: 50%;
                    transform: translateY(-50%); width: 28px; height: 28px; line-height: 28px; border-radius: 50%;
                    background: #222A35; color: #ffffff; text-align: center; font-size: 13px; font-weight: 700; }
                .sortable-item::after { content: "\\22EE\\22EE"; position: absolute; right: 16px; top: 50%;
                    transform: translateY(-50%); color: #aab4c0; font-size: 18px; letter-spacing: -6px; font-weight: 700; }
                .sortable-item:hover { background: #f7fafd !important; border-color: #41719C !important;
                    border-left-color: #41719C !important; box-shadow: 0 3px 10px rgba(20,30,50,.15) !important; }
                """,
            )
            if (isinstance(result, list) and len(result) == len(blocks)
                    and all(r in label_to_block for r in result)
                    and len(set(result)) == len(result)):
                ss.assembly = blocks = [label_to_block[r] for r in result]
        else:
            for i, b in enumerate(blocks):
                _box(f"{i + 1}. " + _block_text(b).replace("\n", " — "))
            st.caption("(끌어서 옮기는 부품이 설치되지 않아 위 도구 모음의 ↑ ↓ 버튼으로 순서를 바꿔주세요.)")

        _box("배너 기획 요청", "마지막 페이지 — 자동으로 들어가요")

        with toolbar:
            if blocks and _ASM_COMP is None:
                ids = [b["id"] for b in blocks]
                # 선택값은 위젯 상태가 아니라 별도 변수(asm_selected)에 둔다 — 드래그로 화면이 다시 그려지는
                # 도중에 위젯 상태가 사라져도 선택이 유지되도록, 매번 이 값으로 위젯을 맞춘다.
                if ss.get("asm_selected") not in ids:
                    ss.asm_selected = ids[0]
                ss.asm_sel = ss.asm_selected
                names = {b["id"]: f"{n + 1}. {ap.BLOCK_LABEL.get(b['type'], b['type'])}"
                         + (f"  ·  {(b.get('title') or '').strip()}" if (b.get("title") or "").strip() else "")
                         for n, b in enumerate(blocks)}
                cur = ids.index(ss.asm_selected)
                t0, t1, t2, t3 = st.columns([6, 1.3, 1.3, 1.5])
                # 선택지 순서는 블록 id 순으로 고정 — 순서를 바꿔도 위젯이 새것으로 인식돼 선택이 풀리는 일을 막는다
                t0.selectbox("편집할 블록", sorted(ids), key="asm_sel",
                             on_change=lambda: ss.__setitem__("asm_selected", ss.asm_sel), format_func=lambda x: names.get(x, x),
                             label_visibility="collapsed")
                t1.button("", key="asm_tb_up", icon=":material/arrow_upward:", help="선택한 블록을 위로",
                          on_click=_asm_move_sel, args=(-1,), disabled=(cur == 0), use_container_width=True)
                t2.button("", key="asm_tb_dn", icon=":material/arrow_downward:", help="선택한 블록을 아래로",
                          on_click=_asm_move_sel, args=(1,), disabled=(cur == len(blocks) - 1),
                          use_container_width=True)
                t3.button("삭제", key="asm_tb_del", icon=":material/delete:", help="선택한 블록 삭제",
                          on_click=_asm_delete_sel, use_container_width=True)
                sel = blocks[cur]
                with st.expander("✏️ 선택한 블록의 제목·요청 쓰기", expanded=False):
                    st.caption(ap.BLOCK_HELP.get(sel["type"], ""))
                    f1, f2 = st.columns(2)
                    sel["title"] = f1.text_input("섹션 제목 (비우면 AI가 작성)", value=sel.get("title", ""),
                                                 key=f"asm_t_{sel['id']}", placeholder="예) 걷는 만큼 달라지는 풍경")
                    sel["memo"] = f2.text_input("AI에게 요청할 점 (선택)", value=sel.get("memo", ""),
                                                key=f"asm_m_{sel['id']}", placeholder="예) 일차별로 풀어서 / 식사만 다뤄줘")

        st.write("")
        st.markdown("**➕ 블록 쌓기** — 누르면 맨 아래에 추가돼요")
        for row_start in range(0, len(ap.BLOCK_CHOICES), 3):
            cols = st.columns(3)
            for col, (t, label, hlp) in zip(cols, ap.BLOCK_CHOICES[row_start:row_start + 3]):
                col.button(label, key=f"asm_add_{t}", on_click=_asm_add, args=(t,), help=hlp,
                           icon=ap.BLOCK_ICON.get(t), use_container_width=True)
        st.caption("같은 블록을 여러 번 쌓을 수 있어요(예: 코스를 '트레킹 코스'와 '관광 코스'로 두 번). "
                   "자료에 근거가 없는 블록은 지어내지 않고 자동으로 건너뜁니다.")

        st.markdown("**자주 쓰는 구성**")
        saved = ap.load_saved()
        s1, s2 = st.columns([3, 1])
        save_name = s1.text_input("구성 이름", key="asm_save_name", placeholder="예) 정현지식 트레킹",
                                  label_visibility="collapsed")
        if s2.button("저장", key="asm_save_btn"):
            ok, msg = ap.save_assembly(save_name, blocks)
            (st.success if ok else st.warning)(msg)
        if saved:
            l1, l2, l3 = st.columns([3, 1, 1])
            pick = l1.selectbox("저장된 구성", list(saved.keys()), key="asm_saved_pick", label_visibility="collapsed")
            l2.button("불러오기", key="asm_load_saved", on_click=_asm_load_saved, args=(pick,))
            if l3.button("삭제", key="asm_del_saved"):
                ap.delete_saved(pick)
                st.rerun()
        st.caption("저장한 구성은 팀이 함께 쓰는 목록이에요. 앱을 다시 시작하면 지워질 수 있으니 오래 쓸 구성은 "
                   "아래 '내보내기'로 파일도 보관하세요.")
        e1, e2 = st.columns(2)
        e1.download_button("📤 파일로 내보내기", json.dumps(ap.blocks_for_export(blocks), ensure_ascii=False, indent=1),
                           file_name="구성_조립.json", mime="application/json", key="asm_export")
        imp = e2.file_uploader("📥 파일에서 가져오기", type=["json"], key="asm_import", label_visibility="collapsed")
        if imp is not None and ss.get("asm_import_done") != imp.file_id:
            try:
                ss.assembly = ap.blocks_from_saved(json.loads(imp.getvalue().decode("utf-8")))
                _asm_bump()
                ss.asm_import_done = imp.file_id
                st.rerun()
            except (ValueError, UnicodeDecodeError):
                st.warning("가져올 수 없는 파일이에요. '내보내기'로 저장한 파일을 선택해주세요.")

    if not ss.assembly:
        st.warning("조립된 블록이 없어서 이번에는 AI가 구성을 정합니다. 블록을 추가하면 직접 조립한 대로 만들어요.")
        return None
    return ap.blocks_for_export(ss.assembly)


assembly = None
if not is_mobile:
    assembly = render_assembly(category)

MAX_FILES = 5
if is_mobile:
    uploaded_files = st.file_uploader(
        "모바일 최적화 대상 이미지 업로드 (기존 PC용 상품소개 이미지, 최대 5개 — 같은 "
        "상품의 이어지는 페이지면 여러 장 함께 올려주세요)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )
else:
    uploaded_files = st.file_uploader(
        "사업부 원본자료 업로드 (.docx, .pptx, 이미지 — 최대 5개, 그중 .docx 최소 1개 필요)",
        type=["docx", "pptx", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

if uploaded_files and len(uploaded_files) > MAX_FILES:
    st.error(f"파일은 최대 {MAX_FILES}개까지만 업로드할 수 있어요. "
             f"지금 {len(uploaded_files)}개가 선택됐어요 — {len(uploaded_files) - MAX_FILES}개를 빼주세요.")
    uploaded_files = None

has_claude_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
has_gemini_key = bool(os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"))
_key_name = {"gemini": "GEMINI_API_KEY", "claude": "ANTHROPIC_API_KEY"}
has_api_key = has_gemini_key if gen_provider == "gemini" else has_claude_key
has_review_key = has_gemini_key if review_provider == "gemini" else has_claude_key
if not has_api_key:
    st.warning(f"{_key_name[gen_provider]}가 설정되지 않았어요. 지금은 프롬프트 미리보기까지만 가능해요. "
               f"(Streamlit Cloud라면 앱 Settings → Secrets에 {_key_name[gen_provider]}를 추가하세요.)")
if not is_mobile and not has_review_key:
    st.info(f"{_key_name[review_provider]}가 설정되지 않았어요. 생성은 가능하지만 {review_ai}가 하는 "
            f"왜곡/표절/사실확인 검수 단계는 건너뜁니다.")

if uploaded_files and st.button("생성하기", type="primary"):
    # 확장자별로 분류: docx/pptx는 텍스트 파싱해서 sections에 병합, 이미지는 별도로 모아서
    # 나중에 API 멀티모달 메시지에 그대로 첨부한다.
    tmp_paths = []
    docx_files, pptx_files, image_files = [], [], []
    for uf in uploaded_files:
        ext = os.path.splitext(uf.name)[1].lower()
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(uf.read())
            tmp_path = tmp.name
        tmp_paths.append(tmp_path)
        if ext == ".docx":
            docx_files.append((uf.name, tmp_path))
        elif ext == ".pptx":
            pptx_files.append((uf.name, tmp_path))
        elif ext in (".jpg", ".jpeg", ".png"):
            image_files.append((uf.name, tmp_path))

    if is_mobile:
        # 모바일 모드는 업로드 자체를 이미지 전용으로 제한해뒀으므로(type=["jpg","jpeg","png"])
        # docx/pptx 파싱이 필요 없다 — 업로드된 파일 전부를 원본 이미지로 취급한다.
        primary_path = tmp_paths[0]
        sections, format_info = {}, {}
        with st.spinner("이미지 확인 중..."):
            image_blocks = build_multimodal_image_blocks(image_files)
        img_count = sum(1 for b in image_blocks if b.get("type") == "image")
        st.success(f"이미지 확인 완료(원본 {len(image_files)}개 → 전송 {img_count}조각) — "
                   f"모바일 최적화 기획안으로 재배치합니다.")
        with st.spinner("프롬프트 조립 중..."):
            prompt = build_mobile_system_prompt(category)
    else:
        if not docx_files:
            st.error("사업부 원본자료(.docx)가 최소 1개는 있어야 해요. "
                     "PPT/이미지는 보조 자료로만 쓰이고, 기준이 되는 워드 문서가 필요합니다.")
            st.stop()

        with st.spinner("사업부 자료 파싱 중..."):
            # 첫 번째 docx를 기준 문서로 사용 (유형 판별 · 카피 초안 감지는 이 문서 기준)
            primary_name, primary_path = docx_files[0]
            sections = parse_docx(primary_path)
            format_info = detect_format_and_draft_copy(sections)

            # 나머지 docx/pptx는 보조 자료로 병합 (키 충돌 방지를 위해 파일명으로 접두)
            for name, path in docx_files[1:]:
                extra = parse_docx(path)
                for k, v in extra.items():
                    sections[f"[추가자료: {name}] {k}"] = v
            for name, path in pptx_files:
                extra = parse_pptx(path)
                for k, v in extra.items():
                    sections[f"[추가자료: {name}] {k}"] = v

            image_blocks = build_multimodal_image_blocks(image_files)
            # 워드 안에 삽입된 이미지(고도표 등)도 원본자료 — 업로드 이미지 뒤에 이어 붙인다.
            docx_img_blocks, docx_img_count = build_docx_embedded_image_blocks(docx_files)
            image_blocks = list(image_blocks) + docx_img_blocks

        if docx_img_count:
            st.info(f"워드 파일 안에 삽입된 이미지 {docx_img_count}개를 찾아 AI에게 함께 전달합니다 "
                    f"(고도표·지도·표 속 수치 반영용).")

        extra_count = len(docx_files) - 1 + len(pptx_files) + len(image_files)
        st.success(
            f"파싱 완료 — 유형 {format_info['format_type']}, "
            f"카피 초안 {'있음 (다듬기 모드)' if format_info['draft_copy'] else '없음 (창작 모드)'}"
            + (f" · 보조 자료 {extra_count}개 반영(pptx/추가 docx {len(docx_files) - 1 + len(pptx_files)}개, "
               f"이미지 {len(image_files)}개)" if extra_count else "")
        )

        with st.spinner("프롬프트 조립 중..."):
            prompt = build_system_prompt(writer_style, category, sections, format_info,
                                          has_images=bool(image_blocks),
                                          planner_request=planner_request,
                                          assembly=assembly)

    with st.expander("조립된 프롬프트 보기"):
        st.text(prompt[:3000])

    if not has_api_key:
        st.info("API 키가 없어 여기서 멈춥니다. 위 프롬프트가 실제로 AI에게 전달될 내용이에요.")
    else:
        try:
            spinner_msg = f"{gen_ai}로 카피 생성 중..." + (" (필요 시 웹 검색 포함)" if enable_web_search else "")
            with st.spinner(spinner_msg):
                content = generate_content(
                    prompt, image_blocks=image_blocks, enable_web_search=enable_web_search,
                    provider=gen_provider,
                )
        except Exception as e:
            st.error(f"AI 콘텐츠 생성 중 오류가 발생했습니다:\n\n{e}")
            st.stop()

        st.success("생성 완료!")

        source_material_text = "\n\n".join(
            f"[{k}]\n{v}" for k, v in sections.items()
        )

        def _show_composition(version_content, tag="", assembly_notes=None):
            """PC 기획안: AI가 이 상품에 맞춰 고른 구성(블록 순서)과 그 이유를 보여준다.
            마음에 안 들면 위 '기획자 요청사항' 칸에 원하는 구성을 적고 다시 생성하면 된다."""
            secs = version_content.get("sections") if isinstance(version_content, dict) else None
            secs = secs if isinstance(secs, list) else []
            if is_mobile or (not secs and not assembly_notes):
                return
            labels = dynamic_builder.BLOCK_LABELS
            with st.expander(f"🧩 이 상품의 구성과 이유{tag} ({len(secs)}개 섹션)", expanded=True):
                note = (version_content.get("composition_note") or "").strip()
                if note:
                    st.markdown(f"**구성 방향** — {note}")
                for i, sec in enumerate(secs, 1):
                    if not isinstance(sec, dict):
                        continue
                    kind = dynamic_builder.normalize_block_type(sec.get("type"))
                    kind_label = labels.get(kind, "기타(일반 표시)")
                    title = (sec.get("title") or sec.get("message") or "").strip().replace("\n", " ")
                    why = (sec.get("why") or "").strip()
                    line = f"{i}. **{title or kind_label}** · {kind_label}"
                    if why:
                        line += f"  \n   ↳ {why}"
                    st.markdown(line)
                for n in (assembly_notes or []):
                    st.warning(n)
                if assembly:
                    st.caption("직접 조립한 구성대로 만들었어요. 바꾸려면 위 '구성 조립'에서 고치고 다시 생성하세요.")
                else:
                    st.caption("구성이 마음에 안 들면 위 '구성 조립'에서 직접 조립하거나, '기획자 요청사항'에 "
                               "원하는 구성을 적고 다시 생성하세요.")

        def _process_version(version_content, version_label=None):
            """검수 → PPTX 조립 → 다운로드 버튼까지, 콘텐츠 한 벌(=한 버전)에 대해
            수행한다. 단일 버전 상품은 이 함수를 한 번만 부르고, "봄/가을 2개 버전"처럼
            사업부 자료가 다중 버전을 명시적으로 요청한 상품은 버전마다 한 번씩 부른다."""
            tag = f" — {version_label}" if version_label else ""

            assembly_notes = []
            if assembly and not is_mobile:
                # 조립한 순서·블록대로 AI 결과를 맞춘다(AI가 임의로 더하거나 바꾼 부분 정리)
                version_content, assembly_notes = dynamic_builder.apply_assembly(version_content, assembly)

            _show_composition(version_content, tag, assembly_notes)

            if is_mobile:
                # 모바일 모드는 사업부 문서(source_material_text)가 없어 Gemini
                # 텍스트 대조 검수(왜곡/표절/사실확인)를 할 기준 원본이 없다 — 원본은
                # 이미지이기 때문. 그래서 이 단계는 건너뛰고, 대신 아래 JSON을 원본
                # 이미지와 직접 비교해 내용 누락이 없는지 확인해달라고 안내한다.
                review = {"issues": [], "summary": ""}
                st.info("ℹ️ 모바일 모드에서는 자동 검수를 건너뜁니다 — 원본이 사업부 문서가 아니라 "
                        "이미지라 텍스트 대조 방식의 사실확인이 어렵습니다. 생성된 JSON을 원본 "
                        "이미지와 비교해 누락된 내용이 없는지 확인해주세요.")
            else:
                try:
                    with st.spinner(f"{review_ai}로 검수 중...{tag} (왜곡/표절/사실확인)"):
                        review = review_content(version_content, source_material_text, provider=review_provider)
                except Exception as e:
                    st.warning(f"검수 중 오류가 발생해 이 단계는 건너뜁니다{tag}:\n\n{e}")
                    review = {"issues": [], "summary": ""}

                if review.get("_dry_run"):
                    st.info(f"ℹ️ {review['_note']}")
                else:
                    issues = review.get("issues", [])
                    if issues:
                        st.warning(f"⚠️ {review_ai} 검수에서{tag} {len(issues)}건이 발견됐습니다 — 확인이 필요합니다.")
                        with st.expander(f"🔎 검수 상세 내역{tag} ({len(issues)}건)", expanded=True):
                            if review.get("summary"):
                                st.text(review["summary"])
                            for issue in issues:
                                st.markdown(
                                    f"**[{issue.get('category')} · {issue.get('severity')}] "
                                    f"{issue.get('field')}**\n\n"
                                    f"> {issue.get('quote')}\n\n"
                                    f"{issue.get('explanation')}"
                                )
                    else:
                        st.info(f"✅ {review_ai} 검수 통과{tag} — 왜곡/날조·저작권/표절·사실확인·비문/맞춤법·구성/배치 "
                                f"이슈가 발견되지 않았습니다.")

            try:
                with st.spinner(f"PPTX 조립 중...{tag}"):
                    path_suffix = f"_{version_label}" if version_label else ""
                    out_base, _ = os.path.splitext(primary_path)
                    out_path = f"{out_base}{path_suffix}_결과.pptx"
                    if is_mobile:
                        # 모바일 최적화 기획안 — 스키마는 build()과 동일하지만
                        # 폰트/여백/배지/배경색이 모바일 가이드라인에 맞게 다르게 그려진다.
                        dynamic_builder.build_mobile(version_content, out_path, review=review)
                        log = [("dynamic_build_mobile", "OK — 모바일 최적화 스타일로 생성됨")]
                    elif writer_style in ("정현지", "박소설", "신윤정", "최정인", "최민지"):
                        # v2: 옛 기획안을 열어 덮어쓰지 않고, 매번 새로 슬라이드를 생성.
                        # review도 같이 넘겨서 검수 결과가 PPTX 마지막 페이지에도 남게
                        # 한다 — 예전엔 이 화면(Streamlit)에만 표시되고 다운로드한
                        # 파일에는 안 남아서, 파일만 넘기면 검수 내역이 사라지는
                        # 문제가 있었다.
                        dynamic_builder.build(version_content, out_path, review=review)
                        log = [("dynamic_build", "OK — 새 슬라이드로 생성됨 (템플릿 재사용 없음)")]
                    else:
                        log = assemble(version_content, writer_style, TEMPLATE_MAP_PATH, out_path)
            except Exception as e:
                st.error(f"PPTX 조립 중 오류가 발생했습니다{tag}:\n\n{e}")
                st.stop()

            ok_count = sum(1 for _, r in log if r == "OK")
            leak_entries = [(k, v) for k, v in log if "LEAK CHECK" in k]

            st.success(f"PPTX 조립 완료!{tag} ({ok_count}/{len(log)} 필드 반영)")

            if leak_entries:
                st.warning(
                    f"⚠️ 아직 매핑되지 않아 이전 템플릿 원본 내용이 남아있는 도형이{tag} "
                    f"{len(leak_entries)}개 있습니다. 다운로드한 파일에서 해당 부분은 "
                    f"직접 확인/수정이 필요합니다."
                )
                with st.expander(f"⚠️ 누수 상세 내역{tag} ({len(leak_entries)}건)", expanded=True):
                    for k, v in leak_entries:
                        st.text(f"{k}\n  → {v}")
            else:
                st.info(f"✅ 누수 검사 통과{tag} — 이전 템플릿 원본 내용이 남아있는 도형이 없습니다.")

            with st.expander(f"조립 로그 전체 보기{tag} (필드별 반영 결과)"):
                for k, v in log:
                    st.text(f"{k} -> {v}")

            # 파일명을 기획자명_카테고리 대신 실제 상품명으로 — 여러 개 만들 때 구분되도록
            product_name = (
                version_content.get("cover", {}).get("product_name")
                or version_content.get("cover", {}).get("headline")
                or f"{writer_style}_{category}"
            )
            safe_name = "".join(
                c for c in str(product_name) if c not in '\\/:*?"<>|\n'
            ).strip()[:60] or f"{writer_style}_{category}"
            name_suffix = f"_{version_label}" if version_label else ""

            with open(out_path, "rb") as f:
                st.download_button(
                    f"📥 PPTX 다운로드{tag}", f,
                    file_name=f"{safe_name}{name_suffix}_기획안.pptx",
                    key=f"download_{version_label or 'single'}",
                )

        # multi_version: 사업부 자료가 "봄 버전과 가을 버전으로 2개" 처럼 결과물 자체를
        # 여러 개로 나눠달라고 명시적으로 요청한 경우에만 온다(prompt_builder.py
        # [버전 분기 판단] 참고) — 카라코람 테스트에서 이런 요청이 있었는데도 결과물이
        # 봄 버전 하나만 나온 문제를 고치기 위해 추가됨. 대부분의 상품은 이 분기를
        # 타지 않고 바로 else로 간다.
        if content.get("multi_version") and isinstance(content.get("versions"), list) and content["versions"]:
            versions = content["versions"]
            st.info(f"ℹ️ 사업부 자료에 여러 버전 제작 요청이 있어 {len(versions)}개 버전으로 나눠 생성했습니다.")
            with st.expander("생성된 JSON 보기 (버전 전체)"):
                st.json(content)
            for version in versions:
                version_label = version.get("version_label") or ""
                st.markdown(f"#### {version_label} 버전" if version_label else "#### 버전")
                _process_version(version, version_label)
        else:
            with st.expander("생성된 JSON 보기"):
                st.json(content)
            _process_version(content, None)

    for p in tmp_paths:
        try:
            os.unlink(p)
        except OSError:
            pass

st.divider()
st.caption(f"현재 등록된 템플릿: {', '.join(available_styles)}")
