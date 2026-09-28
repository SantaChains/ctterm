# -*- coding: utf-8 -*-
"""CT 文件解析模型。

结构保护策略
------------
解析用 lxml（libxml2，C 层）。替换阶段不整体序列化，而是用 sourceline
定位目标元素，在原文文本层做行级替换，非目标字节原样保留。

翻译对象仅两类：Description（单行，已实测无换行）与 DropDownList 名称。
其余字段（AssemblerScript / Address / Offsets / VariableType / Hotkeys /
GroupHeader / Options / Color / ID 等）一律视为保护字段，不做任何改动。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from .luaextract import LuaItem, extract_lua_items
from . import forms as _forms


@dataclass
class FormBlob:
    """<Forms> 下单个 Ascii85 表单块（blob 占据源文件一整行）。"""
    tag: str
    cls: str
    line_no: int                 # 1-based
    blob: str
    root: object = None          # forms.Obj（延迟解析）


@dataclass
class DropItem:
    """DropDownList 中的一行：``ID:名称``。"""
    item_id: str          # 冒号前的 ID（原样保留）
    name: str             # 冒号后的名称（翻译对象）
    line_no: int          # 源文件中的行号（1-based）


@dataclass
class Entry:
    """一个 CheatEntry 的翻译相关视图。"""
    entry_id: str
    description: str | None
    desc_line: int | None            # <Description> 所在行（1-based）
    parents: tuple[str, ...] = ()    # 父级 Description 链（自顶向下，场景消歧用）
    dropdown: list[DropItem] = field(default_factory=list)


def _strip_desc_quotes(text: str) -> str:
    """CE 的 <Description> 文本以一对引号包裹字符串字面量（如 "Data"）。
    提取词条时剥离外层引号，使词条 source 与术语表键（不带引号）对齐。"""
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        return text[1:-1]
    return text


class CTModel:
    """解析 CT 文件，提供全部 CheatEntry 的翻译视图与源码级定位信息。"""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        # collect_ids=False：不构建 id 属性索引，解析更快（本项目用 sourceline 定位）
        parser = etree.XMLParser(remove_blank_text=False, recover=True,
                                 resolve_entities=False, huge_tree=True,
                                 collect_ids=False)
        self.tree = etree.parse(str(self.path), parser)
        self.entries: list[Entry] = []
        self._cached_lines: list[str] | None = None
        # Lua/AA 脚本层安全上下文条目（注释/Caption/Msg/Guide）
        self.lua_items: list[LuaItem] = extract_lua_items(self.source_lines())
        # Ascii85 表单块（表单 UI 字符串，延迟解析）
        self.form_blobs: list[FormBlob] = []
        self._collect_forms()
        self._collect(self.tree.getroot().find("CheatEntries"))

    # ------------------------------------------------------------------
    def _collect_forms(self):
        forms_e = self.tree.getroot().find("Forms")
        if forms_e is None:
            return
        for el in forms_e:
            blob = el.text or ""
            if not blob.strip() or el.get("Encoding") != "Ascii85":
                continue
            self.form_blobs.append(FormBlob(
                tag=el.tag, cls=el.get("Class") or "",
                line_no=el.sourceline,
                blob=blob.strip()))

    # ------------------------------------------------------------------
    def _collect(self, entries_elem, parents: tuple[str, ...] = ()):
        if entries_elem is None:
            return
        for ce in entries_elem.findall("CheatEntry"):
            desc_e = ce.find("Description")
            description = None
            desc_line = None
            if desc_e is not None and desc_e.text:
                description = _strip_desc_quotes(desc_e.text)
                desc_line = desc_e.sourceline

            entry = Entry(entry_id=ce.findtext("ID") or "",
                          description=description,
                          desc_line=desc_line,
                          parents=parents)

            dd_e = ce.find("DropDownList")
            if dd_e is not None and dd_e.text:
                start = dd_e.sourceline
                for offset, raw in enumerate(dd_e.text.split("\n")):
                    s = raw.strip()
                    if not s:
                        continue
                    if ":" in s:
                        iid, _, name = s.partition(":")
                        entry.dropdown.append(DropItem(iid.strip(), name.strip(), start + offset))
                    else:
                        # 无冒号行（异常），整行作为名称
                        entry.dropdown.append(DropItem("", s, start + offset))
            self.entries.append(entry)
            sub = ce.find("CheatEntries")
            if sub is not None:
                # 父级链：自身 Description 入链（供子条目场景消歧）
                chain = parents + ((description,) if description else ())
                self._collect(sub, chain)

    # ------------------------------------------------------------------
    def source_lines(self) -> list[str]:
        """源文件按行拆分（utf-8，保留原始行尾字符）；进程内缓存。

        newline="" 禁用 universal newline 归一化：CRLF 文件每行末尾保留 \\r，
        与写入端（newline=""）配合可做到字节级行尾一致。
        """
        if self._cached_lines is None:
            with self.path.open(encoding="utf-8", newline="") as f:
                text = f.read()
            self._cached_lines = text.split("\n")
        return self._cached_lines

    # ------------------------------------------------------------------
    def collect_terms(self):
        """提取全部翻译词条（去重，含 Lua/AA 脚本层）。

        返回 (terms, stats)。
        terms: {source_text: {types:set, count:int, entries:[eid], dropdown:[(eid, item_id, line)],
                              contexts:set, lua:[LuaItem]}}
        """
        from collections import defaultdict
        terms: dict = defaultdict(lambda: {"types": set(), "count": 0,
                                           "entries": [], "dropdown": [],
                                           "contexts": set(), "lua": []})
        for e in self.entries:
            if e.description:
                t = terms[e.description]
                t["types"].add("Description")
                t["count"] += 1
                t["entries"].append(e.entry_id)
                t["contexts"].update(e.parents[-2:])   # 上下文取最近两级父描述
            for item in e.dropdown:
                if not item.name:
                    continue
                t = terms[item.name]
                t["types"].add("DropDown")
                t["count"] += 1
                t["dropdown"].append((e.entry_id, item.item_id, item.line_no))
                # 上下文 = 所属条目 Description + 最近父级：
                # 同名下拉项（"Item"/"Name"）散布上千个功能下，缺上下文
                # AI 无法消歧（实测 DropDown 7232 行上下文全空）
                own = (e.description,) if e.description else ()
                t["contexts"].update(own + e.parents[-1:])
        for li in self.lua_items:
            t = terms[li.text]
            t["types"].add(f"Lua:{li.kind}")
            t["count"] += 1
            t["lua"].append(li)
        # 表单 UI 字符串（延迟解析：只在需要时解码）
        for fb in self.form_blobs:
            if fb.root is None:
                try:
                    fb.root = _forms.parse_form(_forms.decode_blob(fb.blob))
                except Exception:
                    continue      # 解码失败 → 表单整体跳过，不影响其余链路
            for o, p in _forms.collect_strings(fb.root):
                try:
                    s = p.value.decode("utf-8")
                except UnicodeDecodeError:
                    continue
                if not s.strip():
                    continue
                # \r\n 规范化为 \n（xlsx 往返会把 \r 转义成 _x000D_ 破坏哈希）
                s = s.replace("\r\n", "\n").replace("\r", "\n")
                t = terms[s]
                t["types"].add("Form:String")
                t["count"] += 1
                # 上下文 = 表单标签.对象名（Caption/Tips 归属消歧）
                if o.obj_name:
                    t["contexts"].add(f"{fb.tag}.{o.obj_name}")
        stats = {
            "entries": len(self.entries),
            "unique_terms": len(terms),
            "desc_terms": sum(1 for t in terms.values() if "Description" in t["types"]),
            "dd_terms": sum(1 for t in terms.values() if "DropDown" in t["types"]),
            "lua_terms": sum(1 for t in terms.values() if any(x.startswith("Lua:") for x in t["types"])),
            "total_occurrences": sum(t["count"] for t in terms.values()),
        }
        return terms, stats
