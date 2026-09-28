# -*- coding: utf-8 -*-
"""术语表加载与合并。

优先级（高 → 低，冲突时高者覆盖）：
  1. 用户自定义术语表（--user-terms，tsv/json，格式 EN<TAB>ZH 或 {en: zh}）
  2. glossary_qualified.tsv（合格表：键 100% 来自 CT 真实词条，译文经
     官方 xlsx > id列表 > 人工校对 > combine 四级优先校准，权威）
  3. combine_flat.json（主表，全量聚合、已做异值分离；覆盖 Lua/表单层键）
  4. eldenring.json（本体词表，补充 combine 缺失键）
  5. tarnished.json（褪色者版新增，补充）

异值表（combine_conflicts.tsv）默认不进入自动替换：
同一英文对应多个中文译文（如 Bow→敬礼/长弓），自动替换有歧义风险，
交由人工在 xlsx 中处理。可通过 include_conflicts=True 显式启用
（取第一个非空译文，仅当用户确认场景）。

键含真实换行 / 制表符的词条不进入自动匹配（自动替换只针对
可安全整串匹配的条目；含换行的长文本由 xlsx 人工翻译覆盖）。
"""
from __future__ import annotations

import csv
import io
import json
import os
import re
from pathlib import Path

# 词表根目录：默认 <仓库根>/eldenring/（克隆后可在该目录放入词表文件），
# 可用环境变量 ELDENCT_WORD_ROOT 覆盖；目录缺失时各词表文件按缺失处理（空词表）。
_PROJ_ROOT = Path(__file__).resolve().parent.parent.parent   # eldenct/ → 仓库根

DEFAULT_WORD_ROOT = Path(
    os.environ.get("ELDENCT_WORD_ROOT") or (_PROJ_ROOT / "eldenring"))

# 全角 ASCII 区（U+FF01-U+FF5E）：全角字母/数字/标点是词表污染信号
_FULLWIDTH_ASCII = re.compile(r"[\uFF01-\uFF5E]")


def is_usable_zh(dst: str) -> bool:
    """自动译文可用性判定。

    自动生成的译文（xlsx 预填 / 备注建议 / fallback 替换）必须：
    1. 含中文字符（翻译的语义 = 输出中文）；
    2. 不含全角 ASCII（combine 词表存在 `Dont`→`Ｄont` 这类
       全角污染条目，会让自动替换产出不可读文本）。
    不满足则视为词表退化条目，自动路径一律不采用。
    """
    return (any("\u4e00" <= c <= "\u9fff" for c in dst)
            and not _FULLWIDTH_ASCII.search(dst))


def _load_json_dict(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    with io.open(path, encoding="utf-8") as f:
        data = json.load(f)
    # 仅保留字符串值；嵌套 dict（message ID）不参与自动替换
    return {k: v for k, v in data.items() if isinstance(v, str) and v != ""}


def _load_tsv_dict(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    # QUOTE_NONE：TSV 无 CSV 引号语义，行首的 " 是词条本身的一部分
    # （如 "Champion's Song" Painting；默认 quote 解析会剥掉引号造成键错配）
    with io.open(path, encoding="utf-8-sig", newline="") as f:
        return {row[0]: row[1] for row in csv.reader(f, delimiter="\t", quoting=csv.QUOTE_NONE)
                if len(row) == 2 and row[1].strip()}


class Glossary:
    """合并后的术语表。key=英文原文（不转义换行），value=中文。"""

    def __init__(self, maps: list[dict[str, str]]):
        # maps 按优先级从高到低排列，后写覆盖先写
        self._merged: dict[str, str] = {}
        for m in maps:
            self._merged.update(m)

    @classmethod
    def from_defaults(cls, root: Path = DEFAULT_WORD_ROOT,
                      user_terms: str | Path | None = None,
                      include_conflicts: bool = False) -> "Glossary":
        qualified = _load_tsv_dict(root / "glossary_qualified.tsv")
        combine = _load_json_dict(root / "combine_flat.json")
        eldenring = _load_json_dict(root / "eldenring.json")
        tarnished = _load_json_dict(root / "tarnished.json")
        maps = []
        if user_terms:
            p = Path(user_terms)
            maps.append(_load_json_dict(p) if p.suffix.lower() == ".json" else _load_tsv_dict(p))
        maps.append(qualified)
        maps.append({k: v for k, v in combine.items() if k not in qualified})
        maps.append({k: v for k, v in eldenring.items()
                     if k not in qualified and k not in combine})
        maps.append({k: v for k, v in tarnished.items()
                     if k not in qualified and k not in combine and k not in eldenring})
        if include_conflicts:
            maps.append(_load_tsv_dict(root / "combine_conflicts.tsv"))
        return cls(maps)

    # ------------------------------------------------------------------
    def items(self):
        return self._merged.items()

    def filtered_items(self):
        """排除键含换行/制表符的条目（避免破坏自动替换的整串匹配语义）。"""
        for k, v in self._merged.items():
            if "\n" not in k and "\t" not in k:
                yield k, v
