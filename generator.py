"""Claude API 호출 모듈 — ANTHROPIC_API_KEY 환경변수 필요"""
import os, json

# 서버 실행형 웹 검색 도구. 사업부 자료에 없는 배경지식/사실을 보완할 때 AI가
# 자체적으로 웹을 검색하도록 허용한다(background_story 등). Claude 쪽에서 검색을
# 수행하고 결과를 바로 응답에 반영하므로 별도 도구 실행 루프가 필요 없다.
WEB_SEARCH_TOOL = {"type": "web_search_20260209", "name": "web_search", "max_uses": 5}

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
        max_tokens=8000,
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

    text_blocks = [b.text for b in resp.content if b.type == "text"]
    joined_text = "".join(text_blocks).strip()

    if resp.stop_reason == "refusal":
        raise RuntimeError(
            "AI가 안전 정책상 이 요청을 거부했습니다(stop_reason=refusal). "
            "웹 검색 대상 자료나 상품 내용에 민감한 표현이 없는지 확인해주세요."
        )

    if resp.stop_reason == "max_tokens":
        raise RuntimeError(
            "AI 응답이 max_tokens(8000)에서 잘렸습니다 — JSON이 완성되지 못했습니다. "
            "destinations 개수가 많거나 사업부 자료가 길 때 발생할 수 있습니다. "
            "max_tokens를 더 늘리거나, 스타일 규칙에서 문장 길이를 줄이도록 지시하세요.\n"
            f"응답 마지막 300자: ...{joined_text[-300:]}"
        )

    def _strip_fence(candidate):
        candidate = candidate.strip()
        if candidate.startswith("```"):
            candidate = candidate.split("```")[1]
            if candidate.startswith("json"):
                candidate = candidate[4:]
            candidate = candidate.strip()
        return candidate

    # 웹 검색이 켜지면 Claude가 최종 JSON 앞에 검색 과정을 설명하는 서술문
    # 텍스트 블록을 함께 내보낼 수 있다. 뒤에서부터(최종 답변일 가능성이 높은
    # 블록부터) 순서대로 JSON 파싱을 시도해 서술문을 건너뛴다.
    for block in reversed(text_blocks):
        candidate = _strip_fence(block)
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue

    # 블록 단위로도 안 되면 전체를 이어붙인 텍스트에서 가장 바깥쪽 {...}만 추출해본다.
    start = joined_text.find("{")
    end = joined_text.rfind("}")
    if start != -1 and end > start:
        candidate = joined_text[start:end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass

    raise RuntimeError(
        "AI 응답에서 JSON을 찾지 못했습니다.\n"
        f"응답 원문 일부: ...{joined_text[:500]}..."
    )
