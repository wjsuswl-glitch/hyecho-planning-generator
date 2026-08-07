"""Claude API 호출 모듈 — ANTHROPIC_API_KEY 환경변수 필요"""
import os, json

# 서버 실행형 웹 검색 도구. 사업부 자료에 없는 배경지식/사실을 보완할 때 AI가
# 자체적으로 웹을 검색하도록 허용한다(background_story 등). Claude 쪽에서 검색을
# 수행하고 결과를 바로 응답에 반영하므로 별도 도구 실행 루프가 필요 없다.
WEB_SEARCH_TOOL = {"type": "web_search_20260209", "name": "web_search", "max_uses": 5}

# destinations/altitude_profile 배열이 긴 상품(예: 8개 구간짜리 트레킹)에 banner_copy,
# altitude_profile.highlight 등 필드가 추가되면서 8000으로는 종종 부족해짐 — 여유 있게
# 상향. 실제로는 다 안 써도 되고, 쓴 만큼만 과금되므로(출력 토큰 $10/백만) 올려도
# 비용 영향은 미미함.
MAX_TOKENS = 16000

def generate_content(system_prompt, image_blocks=None, dry_run=False, enable_web_search=True):
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
                return json.loads(candidate), joined_text
            except json.JSONDecodeError:
                continue
        # 블록 단위로도 안 되면 전체를 이어붙인 텍스트에서 가장 바깥쪽 {...}만 추출해본다.
        start = joined_text.find("{")
        end = joined_text.rfind("}")
        if start != -1 and end > start:
            candidate = joined_text[start:end + 1]
            try:
                return json.loads(candidate), joined_text
            except json.JSONDecodeError:
                pass
        return None, joined_text

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
        try:
            return json.loads(combined)
        except json.JSONDecodeError:
            pass
        if cont_resp.stop_reason == "max_tokens":
            raise RuntimeError(
                f"AI 응답이 이어서 생성해도 다시 max_tokens({MAX_TOKENS})에서 잘렸습니다 — "
                "상품 정보가 매우 많은 경우로 보입니다(구간/목적지 수가 많은 상품 등). "
                "generator.py의 MAX_TOKENS를 더 늘리거나, 스타일 규칙에서 문장 길이를 "
                "줄이도록 지시하세요.\n"
                f"응답 마지막 300자: ...{cont_text[-300:]}"
            )
        raise RuntimeError(
            "AI 응답이 max_tokens에서 잘려서 이어쓰기를 시도했지만, 이어붙인 결과를 "
            "JSON으로 파싱하지 못했습니다.\n"
            f"응답 원문 일부: ...{combined[:500]}..."
        )

    result, joined_text = _extract_json(resp)
    if result is not None:
        return result

    # 검색을 많이 반복하다 보면(특히 max_uses 한도에 걸렸을 때) Claude가 "검색을
    # 진행하겠습니다 / 결과를 확인했습니다" 같은 서술문만 내놓고 최종 JSON을 아예
    # 못 내는 경우가 있다 — 이미 확보한 검색 결과가 대화 맥락에 남아있으니, 도구
    # 없이 "설명 없이 JSON만 출력하라"고 한 번 더 요청해서 복구를 시도한다
    # (검색 결과를 버리고 처음부터 다시 시도하는 것보다 훨씬 저렴하고 빠르다).
    if enable_web_search:
        retry_messages = messages + [
            {"role": "assistant", "content": resp.content},
            {"role": "user", "content": (
                "지금까지 확인한 정보만으로 충분합니다. 추가 검색이나 설명 문장 없이, "
                "위에서 지시한 스키마에 맞는 최종 JSON 객체 하나만 출력하세요. "
                "'~하겠습니다', '~확인했습니다' 같은 진행 설명은 한 글자도 포함하지 마세요."
            )},
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
        if retry_result is not None:
            return retry_result
        joined_text = retry_text

    raise RuntimeError(
        "AI 응답에서 JSON을 찾지 못했습니다(설명 없이 JSON만 출력하라는 재요청도 실패).\n"
        f"응답 원문 일부: ...{joined_text[:500]}..."
    )
