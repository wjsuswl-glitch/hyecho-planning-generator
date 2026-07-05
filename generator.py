"""Claude API 호출 모듈 — ANTHROPIC_API_KEY 환경변수 필요"""
import os, json

def generate_content(system_prompt, dry_run=False):
    if dry_run or not os.environ.get("ANTHROPIC_API_KEY"):
        return {
            "_dry_run": True,
            "_note": "ANTHROPIC_API_KEY가 설정되지 않아 실제 호출 대신 프롬프트만 반환합니다.",
            "prompt_preview": system_prompt[:500]
        }

    import anthropic
    client = anthropic.Anthropic()
    resp = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=4000,
        messages=[{"role": "user", "content": system_prompt}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)
