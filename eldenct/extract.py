# -*- coding: utf-8 -*-
"""词条提取与分类。

分类规则
--------
- param-like（参数/字段名，建议不翻译）：
    * 单词无空格，且含下划线 / 数字 / 驼峰（defPlayerDmgCorrectRate_Dark）
  这类在导出 xlsx 时标记"建议保留"，避免误翻导致 CT 失效。
- 不提取（根本不进词条清单）：无字母的纯数字/空格/符号串（'[ 0'、'1.5-2.0'、
  '---' 分隔线）——无语义可译，提取出来只会浪费翻译与审校时间。
  注意：未提取 = 不翻译 = CT 中原样保留，与"提取后留空"效果一致。
- Lua 注释中的代码成分（仅对纯 Lua 层词条生效）：复合赋值/比较运算符、
  函数调用（标识符紧贴左括号，≥3 字符，排除 'Name (说明)' 散文括号）、
  Lua 语句关键字（then/elseif/function/return）、十六进制地址。
  保守设计：宁可漏判（散文留 normal 交给人工），不误伤可翻译注释。
- normal：其余（短语、物品名、功能名）——真正需要翻译的部分。
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from lxml import etree

from .ctmodel import CTModel

# 代码信号：运算符/箭头/十六进制（散文中极少出现；'->' 偶作连接符，
# 误判代价低——param 只是"建议保留"标注，人工仍可翻译）
_CODE_TOKENS = re.compile(
    r"\+=|-=|\*=|/=|==|~=|!=|>=|<=|->|=>|::|\+\+|\b0x[0-9a-fA-F]+\b")
# 函数调用：≥3 字符标识符紧贴左括号（排除散文括号 'Cosmetics (Mirror)'
# 与字面 '\n' 转义造成的假相邻）
_CODE_CALL = re.compile(r"[A-Za-z_]\w{2,}\(")
# Lua 语句关键字（不含 if/for/end/local 等散文高频词）
_CODE_KW = re.compile(r"\b(?:then|elseif|function|return)\b")
# 整行纯关键字（'end'、'else'）
_CODE_KW_LINE = re.compile(r"^(?:end|else|do)\s*$")


def _is_code_like(s: str) -> bool:
    return bool(_CODE_TOKENS.search(s) or _CODE_CALL.search(s)
                or _CODE_KW.search(s) or _CODE_KW_LINE.match(s))


def classify(text: str, lua_comment: bool = False) -> str:
    """返回 'normal' 或 'param'。

    lua_comment: 词条是否仅出现在 Lua/AA 注释层（Description/DropDown/
    Form 字符串面向最终用户，永不套用代码启发式）。
    """
    s = text.strip()
    if not s:
        return "param"
    # 无字母（含空格的数字/符号混合一并排除）：无语义可译
    if not re.search(r"[A-Za-z]", s):
        return "param"
    if lua_comment and _is_code_like(s):
        return "param"
    if re.search(r"\s", s):
        return "normal"
    if "_" in s or re.search(r"\d", s) or re.search(r"[a-z][A-Z]", s):
        return "param"
    return "normal"


def extract_terms(ct_path: str | Path, out_json: str | Path | None = None) -> dict:
    """提取去重词条并分类。out_json 可选输出词条清单。"""
    model = CTModel(ct_path)
    terms, stats = model.collect_terms()

    rows = []
    skipped_no_letter = 0
    for src, info in terms.items():
        # 无字母（纯数字/空格/符号）根本不提取：无语义可译，
        # 未提取 = CT 原样保留，效果等同"提取后留空"且省审校
        if not re.search(r"[A-Za-z]", src):
            skipped_no_letter += 1
            continue
        types = sorted(info["types"])
        # 代码启发式仅对"纯 Lua/AA 注释层"词条生效：
        # 同一文本若也出现在 Description/DropDown/Form（面向最终用户），一律按 UI 文本对待
        pure_lua_comment = bool(types) and all(
            t in ("Lua:LuaComment", "Lua:AAComment") for t in types)
        rows.append({
            "source": src,
            "type": types,
            "count": info["count"],
            "class": classify(src, lua_comment=pure_lua_comment),
            "entries": info["entries"],
            "contexts": sorted(info.get("contexts", set())),
        })
    rows.sort(key=lambda r: (-r["count"], r["source"]))
    stats["class_dist"] = {"normal": 0, "param": 0}
    for r in rows:
        stats["class_dist"][r["class"]] += 1
    stats["skipped_no_letter"] = skipped_no_letter
    # XML 注释节点统计：iter(etree.Comment) 才匹配 <!-- --> 注释节点；
    # iter("Comment") 匹配的是名为 Comment 的元素（CE 的 itemgib 等），语义不同
    stats["xml_comments"] = len(list(model.tree.getroot().iter(etree.Comment)))
    stats["comment_elements"] = len(list(model.tree.getroot().iter("Comment")))

    result = {"ct": str(ct_path), "stats": stats, "terms": rows}
    if out_json:
        Path(out_json).write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result
