"""혜초여행 기획안 자동생성 웹앱 (프로토타입)"""
import streamlit as st
import tempfile, os, json, sys

sys.path.insert(0, os.path.dirname(__file__))
from parser import parse_docx, detect_format_and_draft_copy
from prompt_builder import build_system_prompt
from generator import generate_content
from assembler import assemble

TEMPLATE_MAP_PATH = os.path.join(os.path.dirname(__file__), "template_map.json")

st.set_page_config(page_title="혜초 기획안 자동생성", page_icon="🧳")
st.title("🧳 혜초여행 기획안 자동생성")
st.caption("사업부 원본자료(docx) → AI 카피 생성 → 기획안 PPTX")

with open(TEMPLATE_MAP_PATH, encoding="utf-8") as f:
    available_styles = list(json.load(f).keys())

col1, col2 = st.columns(2)
with col1:
    writer_style = st.selectbox("기획자 스타일", available_styles)
with col2:
    category = st.selectbox("카테고리", ["문탐", "트레킹"])

uploaded = st.file_uploader("사업부 원본자료 업로드 (.docx)", type=["docx"])

has_api_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
if not has_api_key:
    st.warning("ANTHROPIC_API_KEY가 설정되지 않았어요. 지금은 프롬프트 미리보기까지만 가능해요. "
               "터미널에서 `export ANTHROPIC_API_KEY=\"키\"` 설정 후 앱을 다시 실행하면 실제 생성이 가능합니다.")

if uploaded and st.button("생성하기", type="primary"):
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    with st.spinner("사업부 자료 파싱 중..."):
        sections = parse_docx(tmp_path)
        format_info = detect_format_and_draft_copy(sections)

    st.success(f"파싱 완료 — 유형 {format_info['format_type']}, "
               f"카피 초안 {'있음 (다듬기 모드)' if format_info['draft_copy'] else '없음 (창작 모드)'}")

    with st.spinner("프롬프트 조립 중..."):
        prompt = build_system_prompt(writer_style, category, sections, format_info)

    with st.expander("조립된 프롬프트 보기"):
        st.text(prompt[:3000])

    if not has_api_key:
        st.info("API 키가 없어 여기서 멈춥니다. 위 프롬프트가 실제로 AI에게 전달될 내용이에요.")
    else:
        with st.spinner("AI 카피 생성 중..."):
            content = generate_content(prompt)
        st.success("생성 완료!")
        with st.expander("생성된 JSON 보기"):
            st.json(content)

        with st.spinner("PPTX 조립 중..."):
            out_path = tmp_path.replace(".docx", "_결과.pptx")
            log = assemble(content, writer_style, TEMPLATE_MAP_PATH, out_path)

        ok_count = sum(1 for _, r in log if r == "OK")
        st.success(f"PPTX 조립 완료! ({ok_count}/{len(log)} 필드 반영)")
        with open(out_path, "rb") as f:
            st.download_button("📥 PPTX 다운로드", f, file_name=f"{writer_style}_{category}_기획안.pptx")

    os.unlink(tmp_path)

st.divider()
st.caption(f"현재 등록된 템플릿: {', '.join(available_styles)} · 박소설 스타일은 사업부 자료 확보 후 추가 예정")
