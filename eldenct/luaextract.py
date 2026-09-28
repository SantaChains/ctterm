# -*- coding: utf-8 -*-
"""Lua/AA 脚本层的安全上下文提取与替换。

适用范围
--------
CT 的 <AssemblerScript>（含 {$lua} 块——CE 的 AA 脚本可内嵌 Lua，
本表 795 个脚本全部为 AssemblerScript，无独立 <LuaScript> 元素）。

只提取五类"安全上下文"，绝不裸替换字符串：
1. LuaComment   Lua 行注释 ``-- text``（{$lua} 模式下；含 ``1, -- Bow`` 行尾注释
                与单行 ``--[[ ... ]]`` 块注释）
2. AAComment    AA 行注释 ``// text``（汇编模式下）
3. Caption      ``Caption = "..."`` 表单控件标题（Name= 是控件标识符，禁译）
4. Msg          messageDialog/error/showMessage 的字符串实参（单行调用）
5. Guide        长字符串 ``[[ ]]`` 内的引导文本行 ``"key(s)": content``
                （冒号后为空/空白则忽略；``"OK", "Cancel": ...`` 多键特例
                只导出冒号后内容——按源文本去重后天然归并）

不提取：Name = "..."（控件名，Lua 按 win.ctrl["Name"] 引用，改即断链）、
AA 的 [ENABLE]/[DISABLE] 段标记、代码结构与标识符。

字节保真
--------
脚本是 XML 文本节点，源行内 ``<`` 写作 ``&lt;``（实测 L107）。
提取/替换在"未转义逻辑行"域进行，另建 logical→raw 下标映射，
写回时只拼接 raw 前缀 + 重新转义的新文本 + raw 后缀，
非目标字节零改动。写回统一经 XML 转义（& < > "），保持文件合法。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .replace import _escape_xml, _unescape_xml

# ---------------------------------------------------------------------------
# XML raw↔logical 映射

_ENTITY_SCAN = re.compile(r"&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);")


def xml_unescape_map(raw: str) -> tuple[str, list[int]]:
    """返回 (未转义文本, 下标映射)。map[i] = logical[i] 在 raw 中的起始下标；
    map[len(logical)] = raw 结束位置。非实体字符一一对应。"""
    out: list[str] = []
    idx: list[int] = []
    pos = 0
    for m in _ENTITY_SCAN.finditer(raw):
        out.append(raw[pos:m.start()])
        idx.extend(range(pos, m.start()))
        tok = m.group(1)
        if tok == "amp":
            ch = "&"
        elif tok == "lt":
            ch = "<"
        elif tok == "gt":
            ch = ">"
        elif tok == "quot":
            ch = '"'
        elif tok == "apos":
            ch = "'"
        elif tok[1] in "xX":
            ch = chr(int(tok[2:], 16))
        else:
            ch = chr(int(tok[1:]))
        out.append(ch)
        idx.append(m.start())
        pos = m.end()
    out.append(raw[pos:])
    idx.extend(range(pos, len(raw) + 1))
    return "".join(out), idx


def lua_escape(s: str, quote: str) -> str:
    """Lua 短字符串字面量转义（\\ 与引号；\\r\\n 已在上游剔除）。"""
    s = s.replace("\\", "\\\\")
    if quote:
        s = s.replace(quote, "\\" + quote)
    return s


# ---------------------------------------------------------------------------
# 提取条目

@dataclass
class LuaItem:
    kind: str            # LuaComment | AAComment | Caption | Msg | Guide
    line_no: int         # 源文件 1-based
    raw_line: str        # 所在原始行（防漂移对照用）
    start: int           # 可翻译文本在 logical 行中的起始下标
    end: int             # 结束下标（不含）
    text: str            # 可翻译文本（Lua 字面量已反转义；注释为原文）
    quote: str           # 字符串引号（注释为 ""）
    script_index: int    # 所属脚本序号（diff 锚）
    ordinal: int         # 脚本内序号（diff 锚）


_SCRIPT_OPEN = re.compile(r"<AssemblerScript[^>]*>")
_MSG_FUNCS = ("messageDialog", "error", "showMessage")
_CAPTION_RE = re.compile(r"\bCaption\s*=\s*")
_SHORT_STR = re.compile(r'(["\'])((?:\\.|[^\\])*?)\1')
_GUIDE_RE = re.compile(r'^\s*"[^"]+"(?:\s*,\s*"[^"]+")*\s*:\s?(\S.*)$')

# 不视为可翻译注释的行：纯分隔线 / 空白
_SEP_RE = re.compile(r"^[-=*_#~\s]+$")


def _skip_comment_text(t: str) -> bool:
    t = t.strip()
    return (not t) or bool(_SEP_RE.fullmatch(t))


def extract_lua_items(source_lines: list[str]) -> list[LuaItem]:
    """扫描全部脚本区域，返回提取条目（含写回定位信息）。"""
    items: list[LuaItem] = []
    in_script = False
    lua_mode = False          # {$lua} 起 / {$asm} 止
    long_str: str = ""        # "" 无 | "comment" 跨行块注释 | "string" 长字符串
    script_index = -1
    ordinal = 0
    for ln, raw in enumerate(source_lines, start=1):
        if not in_script:
            if _SCRIPT_OPEN.search(raw):
                in_script = True
                lua_mode = False
                long_str = ""
                script_index += 1
                ordinal = 0
                # 开标签行内可能已带 {$lua}（实测形态：<AssemblerScript>{$lua}）
                if "{$lua}" in raw:
                    lua_mode = True
            continue
        if "</AssemblerScript>" in raw:
            in_script = False
            continue
        logical, imap = xml_unescape_map(raw.rstrip("\r\n"))

        # 模式切换（{$lua} / {$asm} 可与代码同行，按先后近似处理）
        if "{$lua}" in logical:
            lua_mode = True
        if "{$asm}" in logical:
            lua_mode = False

        def add(kind: str, start: int, end: int, quote: str = ""):
            nonlocal ordinal
            ordinal += 1
            items.append(LuaItem(kind, ln, raw, start, end,
                                 logical[start:end], quote,
                                 script_index, ordinal))

        # ---- 长字符串 / 块注释 continuation（仅 Lua 模式）----
        if lua_mode and long_str:
            close = logical.find("]]")
            seg = logical[: close if close >= 0 else len(logical)]
            if long_str == "string":
                m = _GUIDE_RE.match(seg)
                if m and m.group(1).strip():
                    add("Guide", m.start(1), m.start(1) + len(m.group(1)))
            else:  # "comment"：块注释续行整行入词条（跳过分隔线）
                if not _skip_comment_text(seg):
                    add("LuaComment", 0, len(seg))
            if close >= 0:
                long_str = ""
            continue
        if lua_mode:
            open_idx = logical.find("[[")
            if open_idx >= 0 and logical.find("]]", open_idx + 2) < 0:
                # 跨行块注释 or 长字符串：看 [[ 前是否紧跟 --
                long_str = "comment" if logical[open_idx - 2:open_idx] == "--" else "string"
                head = logical[open_idx + 2:]
                if head and not _skip_comment_text(head) and long_str == "comment":
                    add("LuaComment", open_idx + 2, len(logical))
                continue

        # ---- Caption = "..."（任何模式）----
        cm = _CAPTION_RE.search(logical)
        if cm:
            sm = _SHORT_STR.match(logical, cm.end())
            if sm and sm.group(2).strip():
                add("Caption", sm.start(2), sm.end(2), sm.group(1))
                continue

        # ---- 消息函数字符串实参（单行调用）----
        for fn in _MSG_FUNCS:
            fi = logical.find(fn + "(")
            if fi < 0 or (fi > 0 and (logical[fi - 1].isalnum()
                                      or logical[fi - 1] == "_")):
                continue
            for sm in _SHORT_STR.finditer(logical, fi):
                if sm.group(2).strip():
                    add("Msg", sm.start(2), sm.end(2), sm.group(1))
            break

        # ---- 注释 ----
        if lua_mode:
            pos, n = 0, len(logical)
            while pos < n:
                c = logical[pos]
                if c in "\"'":                     # 跳过字符串字面量
                    pos += 1
                    while pos < n:
                        if logical[pos] == "\\":
                            pos += 2
                            continue
                        if logical[pos] == c:
                            pos += 1
                            break
                        pos += 1
                    continue
                if logical.startswith("--[[", pos):
                    close = logical.find("]]", pos + 4)
                    if close >= 0:
                        txt = logical[pos + 4: close]
                        if not _skip_comment_text(txt):
                            add("LuaComment", pos + 4, close)
                        pos = close + 2
                        continue
                    # 跨行块注释：本行剩余部分，后续行进入 comment continuation
                    txt = logical[pos + 4:]
                    if not _skip_comment_text(txt):
                        add("LuaComment", pos + 4, n)
                    long_str = "comment"
                    break
                if logical.startswith("--", pos):
                    txt = logical[pos + 2:]
                    if not _skip_comment_text(txt):
                        add("LuaComment", pos + 2, n)
                    break
                pos += 1
        else:
            hi = logical.find("//")
            if hi >= 0 and not _skip_comment_text(logical[hi + 2:]):
                add("AAComment", hi + 2, len(logical))
    return items


def replace_lua_line(raw_line: str, item: LuaItem, new_text: str) -> str:
    """按 item 的 logical span 写回原始行：raw 前缀 + 转义新文本 + raw 后缀。

    注释类直接写 XML 转义文本；字符串类先做 Lua 转义再 XML 转义。
    """
    logical, imap = xml_unescape_map(raw_line.rstrip("\r\n"))
    if logical[item.start:item.end] != item.text:
        raise ValueError("LuaItem 与源行不匹配")
    body = lua_escape(new_text, item.quote) if item.quote else new_text
    s_raw, e_raw = imap[item.start], imap[item.end]
    return raw_line[:s_raw] + _escape_xml(body) + raw_line[e_raw:]
