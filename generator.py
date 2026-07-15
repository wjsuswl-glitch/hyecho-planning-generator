"""Claude API 호출 모듈 — ANTHROPIC_API_KEY 환경변수 필요"""
import os, json

def generate_content(system_prompt, image_blocks=None, dry_run=False):
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

    resp = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=8000,
        thinking={"type": "disabled"},  # 구조화된 JSON 생성엔 추론 불필요.
        # thinking을 켜두면 max_tokens가 "생각+응답" 합산 한도라
        # 응답이 완성되기 전에 잘릴 수 있음 (Sonnet 5부터 기본으로 켜져 있음)
        messages=[{"role": "user", "content": content}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    text = text.strip()

    if resp.stop_reason == "max_tokens":
        raise RuntimeError(
            "AI 응답이 max_tokens(8000)에서 잘렸습니다 — JSON이 완성되지 못했습니다. "
            "destinations 개수가 많거나 사업부 자료가 길 때 발생할 수 있습니다. "
            "max_tokens를 더 늘리거나, 스타일 규칙에서 문장 길이를 줄이도록 지시하세요.\n"
            f"응답 마지막 300자: ...{text[-300:]}"
        )

    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        # 어디서 깨졌는지 앞뒤 맥락을 보여줘서 디버깅 가능하게 함
        start = max(0, e.pos - 150)
        end = min(len(text), e.pos + 150)
        raise RuntimeError(
            f"AI 응답을 JSON으로 파싱하는 데 실패했습니다: {e}\n"
            f"문제 지점 근처 텍스트:\n...{text[start:end]}..."
        ) from e
