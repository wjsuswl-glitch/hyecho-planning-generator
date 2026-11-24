"""기획자 조립 화면용 프리셋·블록 목록·저장/불러오기 (app.py에서 사용).

조립 데이터 형식: [{"type": 블록종류, "title": 기획자가 정한 제목(비우면 AI가 작성), "memo": 이 블록 요청 메모}, ...]
블록 종류(type)는 builder.py의 블록 라이브러리와 같다.
"""
import json
import os
import uuid

# (type, 화면에 보이는 이름, 한 줄 설명) — 화면의 '블록 추가' 목록 순서
BLOCK_CHOICES = [
    ("overview", "개요 (핵심 정보 한눈에)", "기간·지역·난이도·이동·숙박 등 4~6칸 + 지도 자리"),
    ("background", "배경 이야기", "여행지·테마의 배경/역사/유래 한 문단"),
    ("highlights", "하이라이트 (장면형)", "이 여행의 대표 장면을 소제목 + 핵심 키워드 + 사진 + 본문으로"),
    ("points", "하이라이트 (번호형)", "이 상품의 하이라이트를 01·02… 번호 탭으로 (사진 + 본문)"),
    ("course", "코스·방문지", "코스/방문지/식당 목록 (그룹 라벨·일차 배지·스펙 칩·사진)"),
    ("schedule", "일정 한눈에", "일차별 한 줄 요약표"),
    ("cards", "사진 카드", "메뉴·맛집·숙소·객실 등 하나하나 보여줄 것 (2열)"),
    ("audience", "추천 대상", "이런 분께 추천해요 (아이콘 카드)"),
    ("season", "시즌·날씨", "최적기 배너 + 설명 + 월별 기온표"),
    ("qna", "Q&A", "식사·난이도·준비물 등 자주 묻는 질문"),
    ("route_compare", "노선 비교", "대안 코스/노선 비교표"),
    ("spec_table", "스펙 표", "열차·크루즈·숙소 등의 스펙"),
    ("guides", "가이드 소개", "인솔자·가이드 이력"),
    ("service_cards", "혜초와 함께 떠나야 하는 이유", "노쇼핑·인솔자 동행 등 혜초 공통 서비스 카드"),
    ("text", "자유 문단", "위 블록에 안 맞는 설명 글"),
    ("list", "체크 리스트", "포함 사항·특전·준비물 나열"),
    ("table", "자유 표", "식사 구성·호텔 목록·요금표 등"),
    ("closing", "마무리 문구", "맨 끝 감성 한 줄 (CTA 밴드)"),
]
BLOCK_LABEL = {t: label for t, label, _ in BLOCK_CHOICES}
BLOCK_HELP = {t: h for t, _, h in BLOCK_CHOICES}

PRESETS = {
    "트레킹": ["overview", "background", "highlights", "course", "audience", "season", "qna", "service_cards", "closing"],
    "문화탐방": ["background", "points", "course", "cards", "qna", "service_cards"],
    "미식·온천·휴양": ["overview", "text", "cards", "schedule", "list", "audience", "closing"],
    "크루즈·열차": ["highlights", "spec_table", "schedule", "cards", "route_compare", "qna"],
    "성지순례·도보": ["background", "course", "schedule", "qna", "season"],
}
DEFAULT_PRESET_BY_CATEGORY = {"트레킹": "트레킹", "문탐": "문화탐방"}

_SAVED_PATH = os.path.join(os.path.dirname(__file__), "data", "saved_assemblies.json")


def new_block(block_type, title="", memo=""):
    return {"id": uuid.uuid4().hex[:8], "type": block_type, "title": title, "memo": memo}


def preset_blocks(name):
    return [new_block(t) for t in PRESETS.get(name, [])]


def blocks_from_saved(items):
    """저장/가져온 데이터(id 없음)를 화면용 블록(id 포함)으로. 알 수 없는 type은 건너뜀."""
    out = []
    for it in items or []:
        if isinstance(it, dict) and it.get("type") in BLOCK_LABEL:
            out.append(new_block(it["type"], str(it.get("title") or ""), str(it.get("memo") or "")))
    return out


def blocks_for_export(blocks):
    return [{"type": b["type"], "title": b.get("title", ""), "memo": b.get("memo", "")} for b in blocks]


# ---- 서버 저장(팀 공용). Streamlit Cloud는 앱을 다시 시작/배포하면 이 파일이 초기화될 수 있어서,
#      오래 두고 쓸 구성은 화면의 '파일로 내보내기'로 내 컴퓨터에도 보관하는 것을 권장한다.

def load_saved():
    try:
        with open(_SAVED_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_assembly(name, blocks):
    name = (name or "").strip()
    if not name:
        return False, "저장할 이름을 입력해주세요."
    data = load_saved()
    data[name] = blocks_for_export(blocks)
    try:
        os.makedirs(os.path.dirname(_SAVED_PATH), exist_ok=True)
        with open(_SAVED_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except OSError as e:
        return False, f"저장하지 못했어요: {e}"
    return True, f"'{name}' 구성을 저장했어요."


def delete_saved(name):
    data = load_saved()
    if name in data:
        del data[name]
        try:
            with open(_SAVED_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
        except OSError:
            return False
    return True
