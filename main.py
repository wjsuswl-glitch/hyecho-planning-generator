"""전체 파이프라인: 사업부 자료(docx) → 파싱 → 프롬프트 → (AI 생성) → PPTX 조립"""
import sys, json, argparse
from parser import parse_docx, detect_format_and_draft_copy
from prompt_builder import build_system_prompt
from generator import generate_content
from assembler import assemble

TEMPLATE_MAP_PATH = "/Users/yondi/Desktop/자동화_프로토타입/pipeline/template_map.json"

def run(docx_path, writer_style, category, out_path, dry_run=False, manual_json=None):
    print(f"[1/4] 파싱: {docx_path}")
    sections = parse_docx(docx_path)
    format_info = detect_format_and_draft_copy(sections)
    print(f"      → 유형 {format_info['format_type']}, draft_copy={'있음' if format_info['draft_copy'] else '없음'}")

    print(f"[2/4] 프롬프트 조립 ({writer_style} / {category})")
    prompt = build_system_prompt(writer_style, category, sections, format_info)
    print(f"      → 프롬프트 길이 {len(prompt)}자")

    print(f"[3/4] AI 생성 {'(dry-run)' if dry_run else ''}")
    if manual_json:
        with open(manual_json, encoding="utf-8") as f:
            content = json.load(f)
        print("      → manual_json 파일 사용 (API 키 없이 테스트)")
    else:
        content = generate_content(prompt, dry_run=dry_run)
        if content.get("_dry_run"):
            print(f"      → {content['_note']}")
            return {"status": "dry_run_stopped", "prompt": prompt}

    print(f"[4/4] PPTX 조립 → {out_path}")
    log = assemble(content, writer_style, TEMPLATE_MAP_PATH, out_path)
    for field, result in log:
        print(f"      {field}: {result}")

    return {"status": "done", "output": out_path, "log": log}

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("docx_path")
    ap.add_argument("writer_style")
    ap.add_argument("category")
    ap.add_argument("out_path")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--manual-json", default=None)
    args = ap.parse_args()
    result = run(args.docx_path, args.writer_style, args.category, args.out_path,
                 dry_run=args.dry_run, manual_json=args.manual_json)
    print("\n결과:", result["status"])
