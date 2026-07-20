"""혜초여행 기획안 자동생성 웹앱 (프로토타입)"""
import streamlit as st
import tempfile, os, json, sys

sys.path.insert(0, os.path.dirname(__file__))
from parser import parse_docx, detect_format_and_draft_copy, parse_pptx, encode_image_block
from prompt_builder import build_system_prompt
from generator import generate_content
from reviewer import review_content
from assembler import assemble
import builder as dynamic_builder

TEMPLATE_MAP_PATH = os.path.join(os.path.dirname(__file__), "template_map.json")

st.set_page_config(page_title="혜초 기획안 자동생성", page_icon="🧳")
st.title("🧳 혜초여행 기획안 자동생성")
st.caption("사업부 원본자료(docx/pptx/이미지, 최대 5개) → AI 카피 생성 → 기획안 PPTX")

with open(TEMPLATE_MAP_PATH, encoding="utf-8") as f:
    available_styles = list(json.load(f).keys())

col1, col2 = st.columns(2)
with col1:
    writer_style = st.selectbox("기획자 스타일", available_styles)
with col2:
    category = st.selectbox("카테고리", ["문탐", "트레킹"])

enable_web_search = st.checkbox(
    "사업부 자료에 없는 배경지식/사실을 웹 검색으로 보완",
    value=True,
    help="background_story 등에서 사업부 자료에 없는 사실이 필요할 때 AI가 웹 검색으로 확인합니다. "
         "끄면 사업부 자료와 일반 상식 범위 내에서만 작성합니다(검색 비용/시간 절약).",
)

MAX_FILES = 5
uploaded_files = st.file_uploader(
    "사업부 원본자료 업로드 (.docx, .pptx, 이미지 — 최대 5개, 그중 .docx 최소 1개 필요)",
    type=["docx", "pptx", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files and len(uploaded_files) > MAX_FILES:
    st.error(f"파일은 최대 {MAX_FILES}개까지만 업로드할 수 있어요. "
             f"지금 {len(uploaded_files)}개가 선택됐어요 — {len(uploaded_files) - MAX_FILES}개를 빼주세요.")
    uploaded_files = None

has_api_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
if not has_api_key:
    st.warning("ANTHROPIC_API_KEY가 설정되지 않았어요. 지금은 프롬프트 미리보기까지만 가능해요. "
               "터미널에서 `export ANTHROPIC_API_KEY=\"키\"` 설정 후 앱을 다시 실행하면 실제 생성이 가능합니다.")

has_gemini_key = bool(os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"))
if not has_gemini_key:
    st.info("GEMINI_API_KEY가 설정되지 않았어요. 생성은 가능하지만 왜곡/표절/사실확인 "
            "검수 단계는 건너뜁니다. 검수까지 받으려면 `export GEMINI_API_KEY=\"키\"` 설정 후 다시 실행하세요.")

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

        image_blocks = [encode_image_block(path) for _, path in image_files]

    extra_count = len(docx_files) - 1 + len(pptx_files) + len(image_files)
    st.success(
        f"파싱 완료 — 유형 {format_info['format_type']}, "
        f"카피 초안 {'있음 (다듬기 모드)' if format_info['draft_copy'] else '없음 (창작 모드)'}"
        + (f" · 보조 자료 {extra_count}개 반영(pptx/추가 docx {len(docx_files) - 1 + len(pptx_files)}개, "
           f"이미지 {len(image_files)}개)" if extra_count else "")
    )

    with st.spinner("프롬프트 조립 중..."):
        prompt = build_system_prompt(writer_style, category, sections, format_info,
                                      has_images=bool(image_blocks))

    with st.expander("조립된 프롬프트 보기"):
        st.text(prompt[:3000])

    if not has_api_key:
        st.info("API 키가 없어 여기서 멈춥니다. 위 프롬프트가 실제로 AI에게 전달될 내용이에요.")
    else:
        try:
            spinner_msg = "AI 카피 생성 중..." + (" (필요 시 웹 검색 포함)" if enable_web_search else "")
            with st.spinner(spinner_msg):
                content = generate_content(
                    prompt, image_blocks=image_blocks, enable_web_search=enable_web_search,
                )
        except Exception as e:
            st.error(f"AI 콘텐츠 생성 중 오류가 발생했습니다:\n\n{e}")
            st.stop()

        st.success("생성 완료!")
        with st.expander("생성된 JSON 보기"):
            st.json(content)

        try:
            with st.spinner("Gemini로 검수 중... (왜곡/표절/사실확인)"):
                source_material_text = "\n\n".join(
                    f"[{k}]\n{v}" for k, v in sections.items()
                )
                review = review_content(content, source_material_text)
        except Exception as e:
            st.warning(f"검수 중 오류가 발생해 이 단계는 건너뜁니다:\n\n{e}")
            review = {"issues": [], "summary": ""}

        if review.get("_dry_run"):
            st.info(f"ℹ️ {review['_note']}")
        else:
            issues = review.get("issues", [])
            if issues:
                st.warning(f"⚠️ Gemini 검수에서 {len(issues)}건이 발견됐습니다 — 확인이 필요합니다.")
                with st.expander(f"🔎 검수 상세 내역 ({len(issues)}건)", expanded=True):
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
                st.info("✅ Gemini 검수 통과 — 왜곡/표절/사실확인 이슈가 발견되지 않았습니다.")

        try:
            with st.spinner("PPTX 조립 중..."):
                out_path = primary_path.replace(".docx", "_결과.pptx")
                if writer_style in ("정현지", "박소설", "신윤정"):
                    # v2: 옛 기획안을 열어 덮어쓰지 않고, 매번 새로 슬라이드를 생성
                    dynamic_builder.build(content, out_path)
                    log = [("dynamic_build", "OK — 새 슬라이드로 생성됨 (템플릿 재사용 없음)")]
                else:
                    log = assemble(content, writer_style, TEMPLATE_MAP_PATH, out_path)
        except Exception as e:
            st.error(f"PPTX 조립 중 오류가 발생했습니다:\n\n{e}")
            st.stop()

        ok_count = sum(1 for _, r in log if r == "OK")
        leak_entries = [(k, v) for k, v in log if "LEAK CHECK" in k]

        st.success(f"PPTX 조립 완료! ({ok_count}/{len(log)} 필드 반영)")

        if leak_entries:
            st.warning(
                f"⚠️ 아직 매핑되지 않아 이전 템플릿 원본 내용이 남아있는 도형이 "
                f"{len(leak_entries)}개 있습니다. 다운로드한 파일에서 해당 부분은 "
                f"직접 확인/수정이 필요합니다."
            )
            with st.expander(f"⚠️ 누수 상세 내역 ({len(leak_entries)}건)", expanded=True):
                for k, v in leak_entries:
                    st.text(f"{k}\n  → {v}")
        else:
            st.info("✅ 누수 검사 통과 — 이전 템플릿 원본 내용이 남아있는 도형이 없습니다.")

        with st.expander("조립 로그 전체 보기 (필드별 반영 결과)"):
            for k, v in log:
                st.text(f"{k} -> {v}")

        # 파일명을 기획자명_카테고리 대신 실제 상품명으로 — 여러 개 만들 때 구분되도록
        product_name = (
            content.get("cover", {}).get("product_name")
            or content.get("cover", {}).get("headline")
            or f"{writer_style}_{category}"
        )
        safe_name = "".join(
            c for c in str(product_name) if c not in '\\/:*?"<>|\n'
        ).strip()[:60] or f"{writer_style}_{category}"

        with open(out_path, "rb") as f:
            st.download_button("📥 PPTX 다운로드", f, file_name=f"{safe_name}_기획안.pptx")

    for p in tmp_paths:
        try:
            os.unlink(p)
        except OSError:
            pass

st.divider()
st.caption(f"현재 등록된 템플릿: {', '.join(available_styles)}")
