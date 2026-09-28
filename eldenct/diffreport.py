# -*- coding: utf-8 -*-
"""乱序 diff 对照。

与行级 diff 不同，本模块以 CT 条目的稳定标识（ID，已实测全局唯一）
对齐两棵条目树，因此文件内条目顺序变化不影响对照。

模式
----
1. same_ct：同一 CT 在给定译文集下的翻译覆盖对照
2. two_ct：任意两个 CT 按 ID 对照 —— 用于 cn 参考版与 en 新版的术语对齐
   （注意：不同版本 ID 可能不重合，会报告 only_left / only_right）
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

from .ctmodel import CTModel
from . import forms as _forms


def _form_strings(model: CTModel):
    """惰性解码全部表单块，产出 (form_idx, tag, obj_name, 序内唯一键后缀, 规范化文本)。
    键规范化与提取端一致（\\r\\n → \\n，xlsx 往返语义）；
    同一 obj 含多条字符串属性时以行内序号区分，保证键唯一。"""
    for fi, fb in enumerate(model.form_blobs):
        if fb.root is None:
            try:
                fb.root = _forms.parse_form(_forms.decode_blob(fb.blob))
            except Exception:
                continue
        for j, (_o, _p) in enumerate(_forms.collect_strings(fb.root)):
            try:
                s = _p.value.decode("utf-8")
            except UnicodeDecodeError:
                continue
            yield fi, fb.tag, _o.obj_name, j, s.replace("\r\n", "\n").replace("\r", "\n")


def _snapshot(path: str | Path, model: CTModel | None = None) -> dict[tuple[str, str], str]:
    """提取 {(ID, field): text}，field ∈ Description | DropDown:<item_id> |
    Lua:<script>.<ord> | Form:<tag>.<obj>.<ord>"""
    if model is None:
        model = CTModel(path)
    snap = {}
    for e in model.entries:
        if e.description is not None:
            snap[(e.entry_id, "Description")] = e.description
        for it in e.dropdown:
            if it.name:
                snap[(e.entry_id, f"DropDown:{it.item_id}")] = it.name
    for li in model.lua_items:
        snap[(f"L{li.script_index}", f"Lua:{li.script_index}.{li.ordinal}")] = li.text
    for fi, tag, oname, j, s in _form_strings(model):
        snap[(f"F{fi}", f"Form:{tag}.{oname}.{j}")] = s
    return snap


def diff_ct(path_a: str | Path, path_b: str | Path):
    """两 CT 对照（乱序无关，按 ID+field 对齐）。

    返回 (rows, stats)：
    rows: [(key, field_label, text_a, text_b, status)]
      status: changed | same | only_a | only_b
    """
    snap_a, snap_b = _snapshot(path_a), _snapshot(path_b)
    keys = set(snap_a) | set(snap_b)
    rows = []
    for eid, field in sorted(keys, key=lambda k: (k[0], k[1])):
        va, vb = snap_a.get((eid, field)), snap_b.get((eid, field))
        if va is None:
            status = "only_b"
        elif vb is None:
            status = "only_a"
        else:
            status = "same" if va == vb else "changed"
        rows.append((eid, field, va or "", vb or "", status))
    return rows, dict(Counter(r[4] for r in rows))


def diff_same_ct(ct_path: str | Path, translations: dict[str, str]):
    """同一 CT 在给定译文集下的翻译覆盖对照（乱序无关）。

    按出现次数计数（同一 (ID, field) 出现多行也逐处统计），
    使 translated 与行级替换数量严格一致。

    rows: [(entry_id, field, source, translated_or_empty, status)]
      status: translated | unchanged(空译文) | source_unknown(译文集的源不在 CT)
    """
    model = CTModel(ct_path)
    snap = _snapshot(ct_path, model)   # 复用同一模型，避免重复解析
    occ_count: Counter = Counter()   # (eid, field) -> 出现次数
    occ_text: dict = {}              # (eid, field) -> 文本
    for e in model.entries:
        if e.description is not None:
            key = (e.entry_id, "Description")
            occ_count[key] += 1
            occ_text[key] = e.description
        for it in e.dropdown:
            if it.name:
                key = (e.entry_id, f"DropDown:{it.item_id}")
                occ_count[key] += 1
                occ_text[key] = it.name
    for li in model.lua_items:
        key = (f"L{li.script_index}", f"Lua:{li.script_index}.{li.ordinal}")
        occ_count[key] += 1
        occ_text[key] = li.text
    for fi, tag, oname, j, s in _form_strings(model):
        key = (f"F{fi}", f"Form:{tag}.{oname}.{j}")
        occ_count[key] += 1
        occ_text[key] = s
    known = set(occ_text.values())
    # 与 apply 同口径：模板键（'Huw +{0}'）先展开为具体键再统计覆盖，
    # 否则模板行全被误报为 unknown_sources
    from .replace import _TPL_PH, _expand_templates
    lookup, tpl_expanded, _tpl_issues, tpl_used = _expand_templates(
        dict(translations), known)
    unknown = [k for k in translations if k not in known and _TPL_PH not in k]
    applied = 0
    rows = []
    for key, text in sorted(occ_text.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        eid, field = key
        dst = lookup.get(text)
        # dst != text 与 apply 口径一致：identity 译文 apply 会跳过，这里同样不计
        if dst and dst != text:
            # Form 字段与 apply 同口径：译文 ≥255 字节时 apply 拒写（TPF0 pstr 上限），
            # 这里同样不计入 translated，保证 cover 数 == apply 数严格成立
            if field.startswith("Form:") and len(dst.encode("utf-8")) >= 255:
                rows.append((eid, field, text, "", "unchanged"))
                continue
            rows.append((eid, field, text, dst, "translated"))
            applied += occ_count[key]
        else:
            rows.append((eid, field, text, "", "unchanged"))
    total = sum(occ_count.values())
    tpl_unmatched = [k for k in translations
                     if _TPL_PH in k and k not in tpl_used]
    return rows, {"translated": applied, "unchanged": total - applied,
                  "unknown_sources": len(unknown),
                  "template_expanded": tpl_expanded,
                  "templates_unmatched": len(tpl_unmatched)}, unknown


def report_txt(rows, title="CT 对照报告") -> str:
    """落盘报告：全文不截断（截断是屏幕预览策略，写文件丢数据是事故）。"""
    lines = [f"# {title}", ""]
    for eid, field, a, b, status in rows:
        lines.append(f"[{status}] {eid} :: {field}")
        lines.append(f"    A: {a}")
        lines.append(f"    B: {b}")
    return "\n".join(lines)
