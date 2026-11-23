"""카피 생성 모듈 — provider="gemini"(GEMINI_API_KEY) 또는 "claude"(ANTHROPIC_API_KEY).

2026-11: 카피 생성을 Gemini로 바꿀 수 있게 provider 선택을 추가했다(화면에서 선택, 기본 Gemini).
검수(reviewer.py)는 생성과 반대편 AI가 맡는다(같은 AI가 자기 글을 검수하지 않도록)."""
import os, json, base64, re

# 서버 실행형 웹 검색 도구. 사업부 자료에 없는 배경지식/사실을 보완할 때 AI가
# 자체적으로 웹을 검색하도록 허용한다(background_story 등). Claude 쪽에서 검색을
# 수행하고 결과를 바로 응답에 반영하므로 별도 도구 실행 루프가 필요 없다.
WEB_SEARCH_TOOL = {"type": "web_search_20260209", "name": "web_search", "max_uses": 5}

# destinations/altitude_profile 배열이 긴 상품(예: 8개 구간짜리 트레킹)에 banner_copy,
# altitude_profile.highlight 등 필드가 추가되면서 8000으로는 종종 부족해짐 — 여유 있게
# 상향. 실제로는 다 안 써도 되고, 쓴 만큼만 과금되므로(출력 토큰 $10/백만) 올려도
# 비용 영향은 미미함.
MAX_TOKENS = 16000

def _generate_with_claude(system_prompt, image_blocks=None, dry_run=False, enable_web_search=True):
    if dry_run or not os.environ.get("ANTHROPIC_API_KEY"):
        return {
            "_dry_run": True,
            "_note": "ANTHROPIC_API_KEY가 설정되지 않아 실제 호출 대신 프롬프트만 반환합니다.",
            "prompt_preview": system_prompt[:500]
        }

    import anthropic
    client = anthropic.Anthropic()

    # 이미지가 첨부되면 텍스트 프롬프트 뒤에 이미지 블록들을 이어붙여 멀티모달 메시지로 전송.
    # 이미지가 없으면 기존과 동일하게 순수 텍스트 문자열 그대로 보낸다 (하위 호환).
    if image_blocks:
        content = [{"type": "text", "text": system_prompt}] + list(image_blocks)
    else:
        content = system_prompt

    messages = [{"role": "user", "content": content}]
    request_kwargs = dict(
        model="claude-sonnet-5",
        max_tokens=MAX_TOKENS,
        thinking={"type": "disabled"},  # 구조화된 JSON 생성엔 추론 불필요.
        # thinking을 켜두면 max_tokens가 "생각+응답" 합산 한도라
        # 응답이 완성되기 전에 잘릴 수 있음 (Sonnet 5부터 기본으로 켜져 있음)
    )
    if enable_web_search:
        request_kwargs["tools"] = [WEB_SEARCH_TOOL]

    resp = client.messages.create(messages=messages, **request_kwargs)

    # 웹 검색이 서버 쪽에서 10회 이상 반복되면 stop_reason이 "pause_turn"으로
    # 끊길 수 있다 — 별도 도구 실행 없이 그대로 재요청하면 이어서 진행된다.
    # (user.message로 "계속" 등을 덧붙이지 않는다 — trailing server_tool_use를
    # 보고 서버가 자동으로 이어서 진행함)
    resume_attempts = 0
    while resp.stop_reason == "pause_turn" and resume_attempts < 3:
        messages = messages + [{"role": "assistant", "content": resp.content}]
        resp = client.messages.create(messages=messages, **request_kwargs)
        resume_attempts += 1

    if resp.stop_reason == "refusal":
        raise RuntimeError(
            "AI가 안전 정책상 이 요청을 거부했습니다(stop_reason=refusal). "
            "웹 검색 대상 자료나 상품 내용에 민감한 표현이 없는지 확인해주세요."
        )

    def _strip_fence(candidate):
        candidate = candidate.strip()
        if candidate.startswith("```"):
            candidate = candidate.split("```")[1]
            if candidate.startswith("json"):
                candidate = candidate[4:]
            candidate = candidate.strip()
        return candidate

    def _normalize(parsed):
        """Claude가 스키마 객체 하나 대신 [{...}] 처럼 배열로 한 번 더 감싸서
        내놓는 경우가 있다 — 프롬프트에 few-shot 예시가 JSON 배열로 통째로
        박혀 있어서, 그 형태를 따라 최종 답변도 배열로 감싸버리는 것으로 보인다.
        원소 1개짜리 배열이면 그 안의 딕셔너리를 꺼내 정상적인 스키마 객체로
        취급한다(내용 자체는 멀쩡한데 겉모양만 다른 경우를 빈 뼈대로 오판하지
        않기 위함)."""
        if isinstance(parsed, list) and len(parsed) == 1 and isinstance(parsed[0], dict):
            return parsed[0]
        return parsed

    def _extract_json(response):
        """응답의 텍스트 블록에서 JSON을 찾는다. 웹 검색이 켜지면 Claude가 최종
        JSON 앞에 검색 과정을 설명하는 서술문 텍스트 블록을 함께 내보낼 수 있어서,
        뒤에서부터(최종 답변일 가능성이 높은 블록부터) 순서대로 파싱을 시도해
        서술문을 건너뛴다. 실패 시 (None, 이어붙인 원문 텍스트)를 반환한다."""
        text_blocks = [b.text for b in response.content if b.type == "text"]
        joined_text = "".join(text_blocks).strip()
        for block in reversed(text_blocks):
            candidate = _strip_fence(block)
            try:
                return _normalize(json.loads(candidate)), joined_text
            except json.JSONDecodeError:
                continue
        # 블록 단위로도 안 되면 전체를 이어붙인 텍스트에서 가장 바깥쪽 중괄호나
        # 대괄호로 감싸인 부분을 추출해본다(배열로 감싸 나오는 경우 대비 둘 다 확인).
        for open_ch, close_ch in (("{", "}"), ("[", "]")):
            start = joined_text.find(open_ch)
            end = joined_text.rfind(close_ch)
            if start != -1 and end > start:
                candidate = joined_text[start:end + 1]
                try:
                    return _normalize(json.loads(candidate)), joined_text
                except json.JSONDecodeError:
                    pass
        return None, joined_text

    def _is_single_substantive(result):
        """cover.tagline/product_name처럼 항상 채워져야 하는 핵심 필드가 실제로
        차 있는지 확인한다. JSON 파싱 자체는 성공했지만 속이 빈 뼈대(예: 검색·서술문
        문제로 급하게 마무리하며 {}에 가까운 결과를 낸 경우)를 성공으로 잘못
        처리하지 않기 위한 안전장치 — 이게 없으면 오류 없이 텅 빈 PPTX가 조용히
        만들어지는 훨씬 나쁜 상황이 된다."""
        if not isinstance(result, dict):
            return False
        cover = result.get("cover")
        if not isinstance(cover, dict):
            return False
        # .get(key, "") 기본값은 키가 아예 없을 때만 적용되고, 값이 명시적으로
        # null(None)이면 그대로 None이 반환돼 .strip()에서 죽는다 — "or \"\"\"로 방어.
        product_name = cover.get("product_name") or ""
        tagline = cover.get("tagline") or ""
        return bool(product_name.strip()) and bool(tagline.strip())

    def _is_substantive(result):
        """일반 상품(스키마 객체 하나)은 _is_single_substantive를 그대로 쓴다.
        사업부 자료가 "봄/가을 2개 버전" 같은 다중 버전을 명시적으로 요청한 경우엔
        {"multi_version": true, "versions": [...]} 형태로 응답이 온다 — 이때는 껍데기
        자체엔 cover가 없으므로, 감싸인 각 버전(스키마 객체) 전부가 substantive해야만
        전체를 substantive로 인정한다(한 버전이라도 비어있으면 재요청 대상)."""
        if not isinstance(result, dict):
            return False
        if result.get("multi_version"):
            versions = result.get("versions")
            if not isinstance(versions, list) or not versions:
                return False
            return all(_is_single_substantive(v) for v in versions)
        return _is_single_substantive(result)

    joined_text = ""
    empty_skeleton = False  # JSON은 찾았지만 cover.product_name 등 핵심 필드가 빈 경우

    if resp.stop_reason == "max_tokens":
        # 잘린 지점까지의 응답을 assistant 턴으로 그대로 다시 보내 "이어서 계속
        # 생성"하게 하고(pause_turn 재개와 같은 방식), 두 응답의 텍스트를 이어붙여
        # 완성한다 — 처음부터 다시 생성하는 것보다 훨씬 저렴하고 빠르다.
        first_text = "".join(b.text for b in resp.content if b.type == "text")
        continue_messages = messages + [{"role": "assistant", "content": resp.content}]
        continue_kwargs = dict(request_kwargs)
        continue_kwargs.pop("tools", None)  # 이어쓰기 단계에서는 추가 검색이 필요 없음
        cont_resp = client.messages.create(messages=continue_messages, **continue_kwargs)
        cont_text = "".join(b.text for b in cont_resp.content if b.type == "text")
        combined = _strip_fence(first_text + cont_text)
        result = None
        try:
            result = _normalize(json.loads(combined))
        except json.JSONDecodeError:
            pass
        if result is not None and _is_substantive(result):
            return result
        if result is not None:
            empty_skeleton = True
        if cont_resp.stop_reason == "max_tokens":
            raise RuntimeError(
                f"AI 응답이 이어서 생성해도 다시 max_tokens({MAX_TOKENS})에서 잘렸습니다 — "
                "상품 정보가 매우 많은 경우로 보입니다(구간/목적지 수가 많은 상품 등). "
                "generator.py의 MAX_TOKENS를 더 늘리거나, 스타일 규칙에서 문장 길이를 "
                "줄이도록 지시하세요.\n"
                f"응답 마지막 300자: ...{cont_text[-300:]}"
            )
        if result is None:
            raise RuntimeError(
                "AI 응답이 max_tokens에서 잘려서 이어쓰기를 시도했지만, 이어붙인 결과를 "
                "JSON으로 파싱하지 못했습니다.\n"
                f"응답 원문 일부: ...{combined[:500]}..."
            )
        joined_text = combined
    else:
        result, joined_text = _extract_json(resp)
        if result is not None and _is_substantive(result):
            return result
        if result is not None:
            empty_skeleton = True

    # 여기 도달하는 두 경우: (1) 응답 어디에도 JSON이 없음(검색을 많이 반복하다
    # "검색을 진행하겠습니다" 같은 서술문만 내놓고 끝난 경우), (2) JSON은 찾았지만
    # cover.product_name/tagline 등 핵심 필드가 비어 있는 빈 뼈대(급하게 마무리하며
    # {}에 가까운 결과를 낸 경우) — 두 경우 모두 이미 확보한 대화 맥락(검색 결과 등)에
    # 이어서 "설명 없이 스키마를 채운 JSON만 출력하라"고 한 번 더 요청해 복구를
    # 시도한다(처음부터 다시 생성하는 것보다 훨씬 저렴하고 빠르다). 웹 검색을 안 썼어도
    # 빈 뼈대 문제는 발생할 수 있으므로 enable_web_search 여부와 무관하게 시도한다.
    retry_instruction = (
        "지금까지 확인한 정보만으로 충분합니다. 추가 검색이나 설명 문장 없이, "
        "위에서 지시한 스키마의 모든 필드를 실제 내용으로 채운 최종 JSON 객체 하나만 "
        "출력하세요. cover.product_name, cover.tagline처럼 핵심 필드를 빈 문자열로 "
        "두면 안 됩니다. '~하겠습니다', '~확인했습니다' 같은 진행 설명은 한 글자도 "
        "포함하지 마세요."
        if empty_skeleton else
        "지금까지 확인한 정보만으로 충분합니다. 추가 검색이나 설명 문장 없이, "
        "위에서 지시한 스키마에 맞는 최종 JSON 객체 하나만 출력하세요. "
        "'~하겠습니다', '~확인했습니다' 같은 진행 설명은 한 글자도 포함하지 마세요."
    )
    retry_messages = messages + [
        {"role": "assistant", "content": resp.content},
        {"role": "user", "content": retry_instruction},
    ]
    retry_resp = client.messages.create(
        messages=retry_messages,
        model="claude-sonnet-5",
        max_tokens=MAX_TOKENS,
        thinking={"type": "disabled"},
    )
    if retry_resp.stop_reason == "max_tokens":
        raise RuntimeError(
            f"JSON 전용 재요청도 max_tokens({MAX_TOKENS})에서 잘렸습니다 — JSON이 완성되지 못했습니다."
        )
    retry_result, retry_text = _extract_json(retry_resp)
    if retry_result is not None and _is_substantive(retry_result):
        return retry_result
    joined_text = retry_text

    if empty_skeleton or (retry_result is not None):
        raise RuntimeError(
            "AI가 스키마 구조만 있고 핵심 내용(cover.product_name/tagline 등)은 비어 있는 "
            "JSON을 반환했습니다(재요청 포함) — 사업부 자료가 너무 짧거나, 다듬기/창작 "
            "모드 판별이 잘못됐을 가능성이 있습니다. 사업부 자료를 확인하거나 다시 "
            "시도해주세요.\n"
            f"응답 원문 일부: ...{joined_text[:500]}..."
        )

    raise RuntimeError(
        "AI 응답에서 JSON을 찾지 못했습니다(설명 없이 JSON만 출력하라는 재요청도 실패).\n"
        f"응답 원문 일부: ...{joined_text[:500]}..."
    )


# ==== Gemini 생성 경로 ======================================================================
# 생성 모델 이름은 환경변수 GEMINI_GENERATION_MODEL로 바꿀 수 있다(코드 수정 없이 모델 교체).
# 기본 모델 이름이 계정에서 안 열려 있으면(404/NOT_FOUND) 검수에 쓰던 flash 모델로 한 번 대체한다.
GEMINI_GENERATION_MODEL = os.environ.get("GEMINI_GENERATION_MODEL", "gemini-3.5-pro")
GEMINI_FALLBACK_MODEL = "gemini-3.5-flash"
# Gemini는 생각(thinking) 토큰도 이 한도에 합산되므로 Claude 쪽(16000)보다 넉넉히 둔다.
GEMINI_MAX_OUTPUT_TOKENS = 32000


def _strip_fence(candidate):
    candidate = (candidate or "").strip()
    if candidate.startswith("```"):
        parts = candidate.split("```")
        candidate = parts[1] if len(parts) > 1 else candidate
        if candidate.startswith("json"):
            candidate = candidate[4:]
        candidate = candidate.strip()
    return candidate


def _loads_loose(text):
    """응답 텍스트에서 JSON을 최대한 찾아낸다(코드펜스, 앞뒤 설명문, [ {..} ]로 감싼 경우 대응)."""
    text = (text or "").strip()
    candidates = [text, _strip_fence(text)]
    for open_ch, close_ch in (("{", "}"), ("[", "]")):
        a, b = text.find(open_ch), text.rfind(close_ch)
        if a != -1 and b > a:
            candidates.append(text[a:b + 1])
    for c in candidates:
        try:
            parsed = json.loads(c)
        except (json.JSONDecodeError, TypeError):
            continue
        if isinstance(parsed, list) and len(parsed) == 1 and isinstance(parsed[0], dict):
            parsed = parsed[0]
        return parsed
    return None


def _single_substantive(result):
    if not isinstance(result, dict):
        return False
    cover = result.get("cover")
    if not isinstance(cover, dict):
        return False
    return bool((cover.get("product_name") or "").strip()) and bool((cover.get("tagline") or "").strip())


def _substantive(result):
    if not isinstance(result, dict):
        return False
    if result.get("multi_version"):
        versions = result.get("versions")
        return isinstance(versions, list) and bool(versions) and all(_single_substantive(v) for v in versions)
    return _single_substantive(result)


def _gemini_parts(types, system_prompt, image_blocks):
    """Anthropic 형식 이미지/텍스트 블록(parser.py가 만든 것)을 Gemini Part 목록으로 변환."""
    parts = [types.Part.from_text(text=system_prompt)]
    for b in image_blocks or []:
        if b.get("type") == "text":
            parts.append(types.Part.from_text(text=b.get("text", "")))
        elif b.get("type") == "image":
            src = b.get("source", {})
            parts.append(types.Part.from_bytes(
                data=base64.b64decode(src.get("data", "")),
                mime_type=src.get("media_type", "image/jpeg")))
    return parts


def _generate_with_gemini(system_prompt, image_blocks=None, dry_run=False, enable_web_search=True):
    if dry_run or not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        return {
            "_dry_run": True,
            "_note": "GEMINI_API_KEY가 설정되지 않아 실제 호출 대신 프롬프트만 반환합니다.",
            "prompt_preview": system_prompt[:500],
        }

    from google import genai
    from google.genai import types

    client = genai.Client()
    state = {"model": GEMINI_GENERATION_MODEL}

    def call(contents, use_search):
        cfg = dict(max_output_tokens=GEMINI_MAX_OUTPUT_TOKENS)
        if use_search:
            # 검색(grounding)과 JSON 모드(response_mime_type)는 함께 못 쓰는 모델이 있어서,
            # 검색을 켠 호출은 텍스트로 받아 직접 JSON을 추출한다.
            cfg["tools"] = [types.Tool(google_search=types.GoogleSearch())]
        else:
            cfg["response_mime_type"] = "application/json"
        config = types.GenerateContentConfig(**cfg)
        try:
            return client.models.generate_content(model=state["model"], contents=contents, config=config)
        except Exception as e:  # 모델 이름이 계정에서 안 열려 있으면 flash로 한 번 대체
            msg = str(e)
            if state["model"] != GEMINI_FALLBACK_MODEL and ("404" in msg or "NOT_FOUND" in msg or "not found" in msg.lower()):
                state["model"] = GEMINI_FALLBACK_MODEL
                return client.models.generate_content(model=state["model"], contents=contents, config=config)
            raise

    def finish_reason(resp):
        try:
            return str(resp.candidates[0].finish_reason or "")
        except Exception:
            return ""

    def text_of(resp):
        try:
            return resp.text or ""
        except Exception:  # 안전 필터 등으로 텍스트가 없는 경우
            return ""

    user_content = types.Content(role="user", parts=_gemini_parts(types, system_prompt, image_blocks))
    resp = call([user_content], enable_web_search)
    text = text_of(resp)

    if "MAX_TOKENS" in finish_reason(resp) and not _loads_loose(text):
        raise RuntimeError(
            f"Gemini 응답이 출력 한도({GEMINI_MAX_OUTPUT_TOKENS} 토큰)에서 잘려 JSON이 완성되지 못했습니다 — "
            "상품 정보가 매우 많은 경우로 보입니다. generator.py의 GEMINI_MAX_OUTPUT_TOKENS를 늘리거나 "
            "웹 검색을 끄고 다시 시도해주세요."
        )

    result = _loads_loose(text)
    if result is not None and _substantive(result):
        return result
    empty_skeleton = result is not None

    # JSON이 없거나 핵심 필드가 빈 뼈대면: 지금까지의 답을 보여주고 JSON만 다시 요청(검색 없이, JSON 모드)
    retry_instruction = (
        "추가 검색이나 설명 문장 없이, 위에서 지시한 스키마의 모든 필드를 실제 내용으로 채운 최종 "
        "JSON 객체 하나만 출력하세요. cover.product_name, cover.tagline처럼 핵심 필드를 비워두면 안 됩니다."
        if empty_skeleton else
        "추가 검색이나 설명 문장 없이, 위에서 지시한 스키마에 맞는 최종 JSON 객체 하나만 출력하세요."
    )
    retry_contents = [
        user_content,
        types.Content(role="model", parts=[types.Part.from_text(text=text or "(응답 없음)")]),
        types.Content(role="user", parts=[types.Part.from_text(text=retry_instruction)]),
    ]
    retry_resp = call(retry_contents, False)
    retry_text = text_of(retry_resp)
    retry_result = _loads_loose(retry_text)
    if retry_result is not None and _substantive(retry_result):
        return retry_result

    if empty_skeleton or retry_result is not None:
        raise RuntimeError(
            "Gemini가 스키마 구조만 있고 핵심 내용(cover.product_name/tagline 등)은 비어 있는 "
            "JSON을 반환했습니다(재요청 포함) — 사업부 자료를 확인하거나 다시 시도해주세요.\n"
            f"응답 원문 일부: ...{(retry_text or text)[:500]}..."
        )
    raise RuntimeError(
        "Gemini 응답에서 JSON을 찾지 못했습니다(JSON만 출력하라는 재요청도 실패).\n"
        f"종료 사유: {finish_reason(retry_resp) or finish_reason(resp) or '알 수 없음'}\n"
        f"응답 원문 일부: ...{(retry_text or text)[:500]}..."
    )


_CITE_RE = re.compile(r"(?:</?|\()\s*cite\b[^>\n]*>|</\s*cite\s*>", re.I)


def _strip_cite_tags(obj):
    """웹 검색을 쓰면 AI가 본문에 <cite index="..."> 인용 태그를 섞어 내보낼 때가 있다 —
    기획안에 그대로 노출되므로 태그만 지우고 안쪽 글은 남긴다."""
    if isinstance(obj, str):
        return _CITE_RE.sub("", obj)
    if isinstance(obj, list):
        return [_strip_cite_tags(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _strip_cite_tags(v) for k, v in obj.items()}
    return obj


def generate_content(system_prompt, image_blocks=None, dry_run=False, enable_web_search=True, provider="claude"):
    """provider: "gemini" 또는 "claude". 반환 형식은 두 경로가 같다(스키마 JSON 객체 하나)."""
    if str(provider).lower() == "gemini":
        result = _generate_with_gemini(system_prompt, image_blocks=image_blocks, dry_run=dry_run,
                                       enable_web_search=enable_web_search)
    else:
        result = _generate_with_claude(system_prompt, image_blocks=image_blocks, dry_run=dry_run,
                                       enable_web_search=enable_web_search)
    return _strip_cite_tags(result)
