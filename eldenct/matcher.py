# -*- coding: utf-8 -*-
"""基于 pyahocorasick 的多模式最长匹配替换。

为什么需要"最长匹配优先"
------------------------
词表可能同时含 "Bow"（敬礼）与 "Greatbow"（大弓）。若先替换短词条，
"Greatbow" 会被拆成 "Great"+"敬礼"，造成错译。
本模块对所有命中的模式按 leftmost-longest（最左最长）消解：
同一开始位置取最长模式；已被覆盖的区间跳过，保证短词条不会吃掉长词条。

术语表优先级
------------
Glossary 在合并时已按优先级消解同键冲突（用户术语 > combine > eldenring > tarnished）。
本模块只负责"位置重叠"的消解。
"""
from __future__ import annotations

import ahocorasick


class LongestMatcher:
    """多词典自动机。replace() 做最左最长替换。"""

    def __init__(self, glossary_items: list[tuple[str, str]]):
        automaton = ahocorasick.Automaton()
        for src, dst in glossary_items:
            if not src:
                continue
            try:
                automaton.add_word(src, (src, dst))
            except (ValueError, TypeError):
                continue  # 非法模式（如空串）跳过
        if len(automaton):
            automaton.make_automaton()
        self._automaton = automaton

    # ------------------------------------------------------------------
    def replace_with_hits(self, text: str):
        """返回 (替换后文本, 命中列表[(src, dst, start, end)])，供统计与审查。"""
        if not len(self._automaton) or not text:
            return text, []
        matches = []
        for end_idx, (src, dst) in self._automaton.iter(text):
            start = end_idx - len(src) + 1
            before = text[start - 1] if start > 0 else ""
            after = text[end_idx + 1] if end_idx + 1 < len(text) else ""
            # 整词边界：命中片段前后必须是分隔符，拒绝把词从中间切开
            # （词表有 `C`→`Ｃ` 等单字符条目，否则 `Corpse` 会被切出 C）。
            # 词形字符限定 ASCII 字母数字：CJK 邻接视为词边界
            # （'大Bow' 中 Bow 是独立词，Unicode isalnum 会把它误判为词中）
            if ((not before or not (before.isascii() and before.isalnum()))
                    and (not after or not (after.isascii() and after.isalnum()))):
                matches.append((start, end_idx + 1, src, dst))
        if not matches:
            return text, []

        # 最左最长：按 start 升序；同 start 取最长（end 最大）
        matches.sort(key=lambda m: (m[0], -m[1]))
        parts = []
        hits = []
        cursor = 0
        for start, end, src, dst in matches:
            if start < cursor:
                continue
            if start > cursor:
                parts.append(text[cursor:start])
            parts.append(dst)
            hits.append((src, dst, start, end))
            cursor = end
        if cursor < len(text):
            parts.append(text[cursor:])
        return "".join(parts), hits
