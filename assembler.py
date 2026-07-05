"""검증된 텍스트 주입 로직 — JSON 필드를 template_map 기준으로 PPTX 도형에 씀"""
import json
from copy import deepcopy
from pptx import Presentation

NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def _set_text(slide, shape_id, new_text):
    for shape in slide.shapes:
        if shape.shape_id != shape_id or not shape.has_text_frame:
            continue
        tf = shape.text_frame
        if not tf.paragraphs or not tf.paragraphs[0].runs:
            return "SKIPPED(no base run)"
        lines = str(new_text).split("\n")
        base_para_xml = deepcopy(tf.paragraphs[0]._p)
        txBody = tf._txBody
        for p in list(tf.paragraphs):
            txBody.remove(p._p)
        for line in lines:
            new_p = deepcopy(base_para_xml)
            r_elems = new_p.findall(f'.//{NS}r')
            if r_elems:
                keep = r_elems[0]
                for extra in r_elems[1:]:
                    extra.getparent().remove(extra)
                t_elem = keep.find(f'{NS}t')
                t_elem.text = line
            txBody.append(new_p)
        return "OK"
    return "NOT FOUND"

def _get_nested(data, path):
    """'why_hyecho.points.0' 같은 경로 문자열로 JSON에서 값 추출"""
    cur = data
    for key in path.split("."):
        if isinstance(cur, list):
            idx = int(key)
            if idx >= len(cur):
                return None
            cur = cur[idx]
        elif isinstance(cur, dict):
            cur = cur.get(key)
        else:
            return None
        if cur is None:
            return None
    return cur

def assemble(content_json, writer_style, template_map_path, out_path):
    with open(template_map_path, encoding="utf-8") as f:
        tmap = json.load(f)[writer_style]

    prs = Presentation(tmap["template"])
    log = []
    for field_path, (slide_idx, shape_id) in tmap["field_map"].items():
        value = _get_nested(content_json, field_path)
        if value is None:
            log.append((field_path, "NO DATA"))
            continue
        # points가 dict({"title":..,"description":..})면 한 줄로 합침
        if isinstance(value, dict):
            value = "\n".join(str(v) for v in value.values())
        result = _set_text(prs.slides[slide_idx], shape_id, value)
        log.append((field_path, result))

    prs.save(out_path)
    return log
