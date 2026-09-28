# -*- coding: utf-8 -*-
"""向 CT 追加注意事项词条（汉化教程工作流"在表中新建注意事项"的自动化）。

实现约束与 apply 同一哲学：不做全树序列化（会重排整个文件、破坏
CE 未识别属性的字节保真），只在文本层把一个最小 CheatEntry 块插入
根级 ``</CheatEntries>`` 之前，写盘前过严格 XML 校验，ID 全表唯一。

词条形态用 GroupHeader（纯文本显示、无脚本、无激活语义）——与本表
502 处现有分组头一致，CE 打开即显示为一行文字。
"""
from __future__ import annotations

import random
import re
from pathlib import Path

from lxml import etree

_ENTRY_TPL = """{i}<CheatEntry>
{i}  <ID>{eid}</ID>
{i}  <Description>"{text}"</Description>
{i}  <Color>000080</Color>
{i}  <GroupHeader>1</GroupHeader>
{i}</CheatEntry>
"""

_ID_RE = re.compile(r"<ID>(\d+)</ID>")


def _escape_xml(s: str) -> str:
    """Description 文本转义（引号在 CT 里以 &quot; 存储不破坏属性外层）。"""
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def existing_ids(content: str) -> set[str]:
    """全表已有词条 ID（含嵌套分组），用于唯一性校验。"""
    return set(_ID_RE.findall(content))


def _new_id(taken: set[str]) -> str:
    """生成 10 位随机 ID：避开 6-9 位游戏数值 ID 区，碰撞重试。"""
    while True:
        eid = str(random.randint(1_000_000_000, 9_999_999_999))
        if eid not in taken:
            return eid


def add_note_entry(ct_path: str | Path, text: str,
                   out_path: str | Path | None = None) -> dict:
    """向根级词条列表末尾追加一条注意事项，返回 {"id", "out"}。

    源文件只读；不传 out_path 时写到新文件（-note 后缀），绝不覆盖源。
    缩进处理：content[:tail] 以 </CheatEntries> 行的收尾缩进结尾，
    必须先截到行首再拼接，否则首行缩进叠加（实测 2+4=6 空格错位）。
    """
    src = Path(ct_path)
    content = src.read_text(encoding="utf-8")
    tail = content.rfind("</CheatEntries>")
    if tail == -1:
        raise ValueError("未找到根级 </CheatEntries>，不是合法 CT 结构")
    line_start = content.rfind("\n", 0, tail) + 1
    close_indent = content[line_start:tail]
    if close_indent.strip():
        raise ValueError("根级 </CheatEntries> 行含非空白前缀，结构异常")
    child_indent = close_indent + "  "
    taken = existing_ids(content)
    eid = _new_id(taken)
    block = _ENTRY_TPL.format(i=child_indent, eid=eid, text=_escape_xml(text))
    head = content[:line_start]          # 截到上一行行尾（剥离收尾缩进）
    new_content = head + block + close_indent + content[tail:]
    # 写盘前严格校验（与 apply 同一防线）：坏结构永不落盘
    etree.fromstring(new_content.encode("utf-8"),
                     etree.XMLParser(recover=False, resolve_entities=False))
    out = Path(out_path) if out_path else src.with_name(src.stem + "-note.CT")
    from .replace import _atomic_write
    _atomic_write(out, new_content)
    return {"id": eid, "out": str(out), "total_ids": len(taken) + 1}
