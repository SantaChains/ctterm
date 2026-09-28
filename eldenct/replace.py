# -*- coding: utf-8 -*-
"""精准替换（结构保护核心）。

原理
----
不用 lxml 整体序列化（那会重写全文件、可能改变缩进/属性/自闭合样式）。
改为：源文件按行读入 → 依据 lxml 提供的 sourceline 精确定位每个
Description（单行，已实测）与 DropDownList（行区间）→ 只替换目标行内
的文本内容，其余字节原样。替换后再重新解析校验，报告未生效项。

安全设计
--------
1. 输出路径与源路径 resolve 后相同 → 拒绝执行（绝不覆盖源文件）。
2. 写入用 newline=""（关闭换行转换）+ 临时文件 os.replace 原子落盘：
   源是 CRLF 则产物 CRLF、源是 LF 则产物 LF，字节级一致。
3. 译文先 sanitize：剔除 XML 非法控制字符；再做实体转义（& < > "）。
4. 替换后以严格模式（recover=False）重新解析，结构被破坏立即报错。
5. fallback_glossary 默认关闭；开启时也只作用于"xlsx 未出现过的词条"，
   用户留空的行永远不会被术语表兜底覆盖。

Description 行格式：  <Description>"英文原文"</Description>
  - 保持外层引号（CE 的字符串字面量语法）
  - 内部文本做 XML 实体转义后写入
DropDownList 行格式：  ID:英文名称（每行一条）
  - 只替换冒号后的名称，ID 原样保留
"""
from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path

from lxml import etree

from .glossary import is_usable_zh
from . import forms as _forms   # forms 不依赖包内其他模块，无循环风险

_DESC_RE = re.compile(r'^(?P<pre>.*<Description>)(?P<inner>.*)(?P<post></Description>.*)$', re.S)
# DropDown 行：冒号分隔；name 非贪婪，显式排除行尾的 </DropDownList> 结束标签
# （DropDownList 最后一行形如 "ID:Name</DropDownList>"，贪婪 .* 会吞掉标签）
_DD_RE = re.compile(r'^(?P<pre>.*?)(?P<sep>:[ \t]*)(?P<name>.*?)(?P<tail></DropDownList>)?\r?$', re.S)

# XML 1.0 非法控制字符（\x00-\x08 \x0B \x0C \x0E-\x1F），译文里出现则剔除
_XML_ILLEGAL_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
# 译文中的换行/回车会破坏单行结构（Description 假设单行，写入 \n 会把
# 一行物理拆成多行、行号整体错位），必须剔除
_NEWLINE_RE = re.compile(r"[\r\n]")


def _escape_xml(text: str) -> str:
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


# XML 语境专用反转义：仅五个预定义实体 + 数字字符引用。
# 不用 html.unescape——它会把 HTML-only 实体（&nbsp;、无分号形式）也解开，
# 与 XML 语义不符（CT 是 XML，源中不可能出现合法的 HTML-only 实体）。
_ENTITY_RE = re.compile(r"&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);")


def _entity_sub(m: re.Match) -> str:
    tok = m.group(1)
    if tok == "amp":
        return "&"
    if tok == "lt":
        return "<"
    if tok == "gt":
        return ">"
    if tok == "quot":
        return '"'
    if tok == "apos":
        return "'"
    if tok.startswith("#x") or tok.startswith("#X"):
        return chr(int(tok[2:], 16))
    return chr(int(tok[1:]))


def _unescape_xml(text: str) -> str:
    # 注意：正则一次遍历（替代链式 replace），&amp;lt; 不会被二次解码
    return _ENTITY_RE.sub(_entity_sub, text)


# 译文二次转义检测：译者从 XML/网页复制译文时可能带入 &amp; 等预转义实体，
# 写入时会再被 _escape_xml 转义成 &amp;amp;（CE 显示乱字面量）。
# 检测到则先解一层，让写回路径统一完成唯一一次转义。
_DOUBLE_ESCAPE_RE = re.compile(r"&(amp|lt|gt|quot|apos);")


def _clean(text: str) -> str:
    """译文净化：剔除 XML 非法控制字符与 \r\n；解除二次转义。

    \r\n 必须剔除：Description/DropDown 均为单行结构，译文含换行会把
    一行物理拆成多行，导致行号整体错位、文件体积膨胀。
    制表符 \t 保留（合法空白，不影响结构）。
    """
    if _DOUBLE_ESCAPE_RE.search(text):
        text = _unescape_xml(text)
    return _NEWLINE_RE.sub("", _XML_ILLEGAL_RE.sub("", text))


def _replace_description_line(line: str, old_text: str, new_text: str):
    """替换单行 Description 的内部文本；不匹配则返回 None。

    old_text / new_text 均为剥离外层引号后的文本。
    源行内匹配时兼容两种写法：引号包裹（标准 CE 格式）与无引号。
    """
    m = _DESC_RE.match(line)
    if not m:
        return None
    inner = _unescape_xml(m.group("inner"))
    if inner != old_text and inner != '"' + old_text + '"':
        return None
    return m.group("pre") + '"' + _escape_xml(_clean(new_text)) + '"' + m.group("post")


def _replace_dd_line(line: str, item_name: str, new_text: str):
    """替换 DropDownList 单行的名称部分（ID: 前缀保留）。

    源行名称可能含 XML 实体（&amp; 等），比较前先 unescape；
    新名称写入前做实体转义（名称里的 & < > " 必须合法化）。
    行尾可能携带 </DropDownList> 结束标签（最后一行），原样保留。
    """
    m = _DD_RE.match(line)
    if not m:
        return None
    tail = m.group("tail") or ""
    if _unescape_xml(m.group("name").strip()) != item_name.strip():
        return None
    return m.group("pre") + m.group("sep") + _escape_xml(_clean(new_text)) + tail


def _atomic_write(out_path: Path, content: str):
    """newline='' 原样写（保留源行尾风格），先写临时文件再原子替换。"""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(out_path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            f.write(content)
        os.replace(tmp, out_path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


# 数字模板展开：xlsx 中 'Huw +{0}' 一行对应 CT 中 '+1'/'+2'/… 多个具体变体。
# 对 CT 已知文本逐条尝试模板匹配：数字跑占位化 → 命中模板键 → 按出现顺序回填。
_DIG_RUN = re.compile(r"\d+")
_TPL_PH = "{0}"


def _expand_templates(lookup: dict[str, str], known: set[str]):
    """把模板键展开为具体键。返回 (展开后 lookup, 展开成功数, 问题列表, 已用模板键集)。

    精确键优先：文本已有译文则不展开。占位符数与数字跑数不符的译文
    判为问题并跳过（宁可漏替换，绝不产出丢失数字的错译）。
    """
    tpl = {k: v for k, v in lookup.items() if _TPL_PH in k}
    if not tpl:
        return lookup, 0, [], set()
    from .xlsxio import fill_template, make_template  # 延迟导入（重依赖）
    expanded = dict(lookup)
    n = 0
    used: set[str] = set()
    issues: list[str] = []
    for text in known:
        if text in expanded or _TPL_PH in text or not _DIG_RUN.search(text):
            continue
        tkey = make_template(text)
        v = tpl.get(tkey)
        if v is None:
            continue
        dst = fill_template(v, _DIG_RUN.findall(text))
        if dst is None:
            issues.append(f"模板占位符数不符: {v[:40]!r} (需 {_DIG_RUN.findall(text)}) 对 {text[:40]!r}")
            continue
        if dst != text:
            expanded[text] = dst
            used.add(tkey)
            n += 1
    return expanded, n, issues, used


def strict_parser():
    """严格模式解析器（供内存字节流校验复用）。"""
    return etree.XMLParser(remove_blank_text=False, recover=False,
                           resolve_entities=False, huge_tree=True,
                           collect_ids=False)


def apply_translations(ct_path: str | Path, translations: dict[str, str],
                       out_path: str | Path,
                       glossary_items: list[tuple[str, str]] | None = None,
                       fallback_glossary: bool = False,
                       exported_sources: set[str] | None = None,
                       allowed_types: set[str] | None = None) -> dict:
    """执行替换。返回统计信息。

    translations: {源文本: 译文}（最高优先级，通常来自 xlsx 导回）
    glossary_items: 术语表条目（仅 fallback 用）
    fallback_glossary: 仅当 True 时，对"xlsx 从未出现的词条"按术语表补漏。
        出现在 exported_sources 中的词条（无论译文是否为空）绝不被兜底覆盖——
        用户留空 = 明确不译。
    exported_sources: 导出时出现的全部 source 集合（用于界定 fallback 范围）。
    allowed_types: 类型开关（None = 全部）。 occurrence 的类型 ∈
        {"Description", "DropDown", "Lua:LuaComment", "Lua:AAComment",
         "Lua:Caption", "Lua:Msg", "Lua:Guide"}；未勾选的类型绝不参与替换。
    """

    def _ok(t: str) -> bool:
        return allowed_types is None or t in allowed_types

    src = Path(ct_path).resolve()
    out = Path(out_path).resolve()
    if out == src:
        raise ValueError("输出路径与源文件相同，拒绝执行（绝不覆盖源文件）。请指定不同的 -o 路径。")

    from .ctmodel import CTModel  # 延迟导入，避免 ctf → luaextract → replace 循环
    model = CTModel(src)
    lines = model.source_lines()

    # 收集所有已知源文本（用于校验错配）
    known = {e.description for e in model.entries if e.description}
    known.update(it.name for e in model.entries for it in e.dropdown if it.name)
    known.update(li.text for li in model.lua_items)
    for fb in model.form_blobs:
        if fb.root is None:
            try:
                fb.root = _forms.parse_form(_forms.decode_blob(fb.blob))
            except Exception:
                continue
        for _o, _p in _forms.collect_strings(fb.root):
            try:
                s = _p.value.decode("utf-8")
            except UnicodeDecodeError:
                continue
            known.add(s)
            # 提取端键做了 \r\n→\n 规范化（xlsx 往返），known 口径须一致，
            # 否则规范化键被误报进 unknown_sources（替换本身不受影响）
            s_norm = s.replace("\r\n", "\n").replace("\r", "\n")
            if s_norm != s:
                known.add(s_norm)

    # lookup 基础 = 用户 xlsx（最高优先级）；fallback 仅补"从未导出"的词条
    lookup = dict(translations)
    if fallback_glossary and glossary_items:
        exported = exported_sources or set()
        lookup.update((s, d) for s, d in glossary_items
                      if s not in exported and s not in lookup
                      and is_usable_zh(d))

    # 数字模板展开（'Huw +{0}' → '+1'/'+2'/… 具体键）：
    # 精确键优先，展开只补空白；对全部四层（Desc/DD/Lua/Form）统一生效
    lookup, tpl_expanded, tpl_issues, tpl_used = _expand_templates(lookup, known)

    desc_applied = dd_applied = lua_applied = 0
    missed: list[str] = []
    for e in model.entries:
        if e.description and e.desc_line and _ok("Description"):
            dst = lookup.get(e.description)
            if dst and dst != e.description:
                idx = e.desc_line - 1
                if 0 <= idx < len(lines):
                    new = _replace_description_line(lines[idx], e.description, dst)
                    if new is not None:
                        lines[idx] = new
                        desc_applied += 1
                    else:
                        missed.append(f"[{e.entry_id}] Description 行未匹配: {e.description[:50]!r}")
        for item in e.dropdown:
            if not _ok("DropDown"):
                break
            dst = lookup.get(item.name)
            if dst and dst != item.name:
                idx = item.line_no - 1
                if 0 <= idx < len(lines):
                    new = _replace_dd_line(lines[idx], item.name, dst)
                    if new is not None:
                        lines[idx] = new
                        dd_applied += 1
                    else:
                        missed.append(f"[{e.entry_id}] DropDown 行未匹配: {item.name[:50]!r} @L{item.line_no}")

    # 同一行可能存在多个 Lua 词条（如 L107 三元表达式内两个 messageDialog
    # 字符串）。按 (行号, 行内起始) 倒序替换：先替换靠后的 span 不影响靠前
    # span 的字节偏移，正序则会因前行内容变化导致后续 item 校验失败。
    for li in sorted(model.lua_items,
                     key=lambda x: (x.line_no, x.start), reverse=True):
        ltype = f"Lua:{li.kind}"
        if not _ok(ltype):
            continue
        dst = lookup.get(li.text)
        if dst and dst != li.text:
            idx = li.line_no - 1
            if 0 <= idx < len(lines):
                try:
                    from .luaextract import replace_lua_line  # 延迟导入，打破 import 环
                    lines[idx] = replace_lua_line(lines[idx], li, dst)
                    lua_applied += 1
                except ValueError:
                    missed.append(f"[Lua:{li.kind}] 行未匹配: {li.text[:50]!r} @L{li.line_no}")
            else:
                missed.append(f"[Lua:{li.kind}] 行号越界 @L{li.line_no}")

    # Ascii85 表单块：解码 → 替换字符串属性 → 重编码 → 整行替换
    # （blob 占据源文件一整行，是唯一允许"整行替换"的形态；blob 字符集
    #   为 XML 安全子集，无需转义；行尾 \r 因只做子串替换而天然保留）
    form_applied = 0
    form_skipped_long = 0
    if allowed_types is None or "Form:String" in allowed_types:
        for fb in model.form_blobs:
            if fb.root is None:
                continue
            n, skipped = _forms.apply_form_translations(fb.root, lookup)
            form_skipped_long += skipped
            if n:
                new_blob = _forms.encode_form(_forms.serialize_form(fb.root))
                idx = fb.line_no - 1
                if 0 <= idx < len(lines) and fb.blob in lines[idx]:
                    lines[idx] = lines[idx].replace(fb.blob, new_blob, 1)
                    form_applied += n
                else:
                    missed.append(f"[Form:{fb.tag}] blob 行未匹配 @L{fb.line_no}")
    if form_skipped_long:
        # 显式暴露：译文超 255 字节被拒写的条数（否则"看似已译实则未译"无从排查）
        missed.append(f"[Form] {form_skipped_long} 条译文超 255 字节被拒写"
                      f"（TPF0 pstr 上限），请缩短译文")

    # 写盘前严格校验（recover=False）：任何替换引入的非法字符/结构破坏
    # 在此拦截，拒绝落盘——比"写后再验"更安全（坏文件永远不会出现在磁盘上）
    content = "\n".join(lines)
    try:
        etree.fromstring(content.encode("utf-8"), strict_parser())
    except etree.XMLSyntaxError as e:
        raise ValueError(f"替换结果未通过严格 XML 校验，已放弃写盘：{e}") from e

    _atomic_write(out, content)
    return {
        "desc_applied": desc_applied,
        "dd_applied": dd_applied,
        "lua_applied": lua_applied,
        "form_applied": form_applied,
        "template_expanded": tpl_expanded,
        "template_issues": tpl_issues[:10],
        "missed": missed,
        # unknown 排除模板键：模板键是合法的合并形态，其未命中的情况
        # 由 templates_unmatched 单独报告（CT 改版 / 占位符数不符）
        "unknown_sources": [k for k in translations
                            if k not in known and _TPL_PH not in k][:20],
        "templates_unmatched": [k for k in translations
                                if _TPL_PH in k and k not in tpl_used][:10],
        # fallback 安全告警：exported.json 缺失时无法界定"用户曾导出"的词条，
        # 兜底范围静默扩大到全部词表——用户留空行可能被覆盖，必须显式提示
        "fallback_warning": (
            "exported.json 缺失，fallback 范围未受限（含用户留空行）"
            if fallback_glossary and exported_sources is None else ""),
    }
