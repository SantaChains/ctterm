# -*- coding: utf-8 -*-
"""AI 翻译模块（默认 DeepSeek，OpenAI 兼容 /chat/completions）。

设计动机
--------
xlsx 是给人看的（8 列 + 样式 + 预填混杂），直接丢给 AI 有三个问题：
1. 噪音大：预填行 / 副表 / 备注稀释有效信息，浪费 token 且干扰输出；
2. 无术语锁定：'感应/Arcane' 这类官方译名靠模型自觉，必错；
3. 无校验：AI 漏掉 {0} 占位符或数字，直接替换会产出坏 CT。

本模块职责（与人工译文同一条链路，不另起炉灶）
------------------------------------------------
1. read_pending : 从 terms.xlsx 主表提取待译行（normal 且译文为空）
2. translate    : 分批调 API + checkpoint 断点续跑（崩溃不丢进度）
3. validate     : 严格校验（{0} 数量 / 数字保真 / 术语未译 / 空译），
                  不合格行一律留白并写 flagged 报告——宁缺勿错
4. fill_xlsx    : 合格译文填回 xlsx 副本（备注标"AI 译文（待校对）"），
                  之后照常 导入 → apply

防错译：每行只挂"该行实际命中的术语"（LongestMatcher 整词命中），
官方属性名（Arcane=感应 等）内置于最高优先级，压过词表异值。
"""
from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"

_DIG_RUN = re.compile(r"\d+")
_SUGG_MARK = "词表建议参考："
# 源串已含中文/假名（zh 基底混排词条）：无翻译增益，AI 只会原样返回。
# 实测 1127 条「译文与原文相同」拦截中 1112 条属此类——纯烧 API，读取即跳过。
_HAS_CJK_SRC = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff66-\uff9f]")

# 官方属性/状态译名（最高优先级，压过词表里的异值如 Arcane→奥术）。
# 来源：艾尔登法环官方简中。用户术语表（--user-terms）仍可再覆盖。
OFFICIAL_STATS = {
    "Vigor": "生命力",
    "Mind": "集中力",
    "Endurance": "耐力",
    "Strength": "力气",
    "Dexterity": "灵巧",
    "Intelligence": "智力",
    "Faith": "信仰",
    "Arcane": "感应",
    "HP": "HP",
    "FP": "FP",
    "Stamina": "精力值",
}

_SYSTEM_PROMPT = """[ROLE] FromSoftware 官方术语库锁定的 CT（Cheat Engine 修改表）翻译引擎。
术语优先级：艾尔登法环 > 黑暗之魂 > 血源诅咒 > 只狼；四级均无官方译名时才允许自译。

[TASK] 将输入 JSON 数组中每行的 src 由英文翻译为简体中文，逐行回填。

[RULES] 按优先级降序，冲突时高者优先：
R1 术语锁定：行内 terms 给出的 EN→ZH 为官方译名，必须逐字采用；同一英文术语全文复现同一译名，不得省略、不得改写、不得换译。奥术为错译，Arcane 官方译名=感应。
R2 格式锁：原样保留所有符号、驼峰、下划线、括号、占位符 {0}（数量不得增减）、数值、百分比、公式、空行与缩进；不增删换行。ref 是词表自动参考（可采纳、可修正），ctx 是界面位置或同族词条，供消歧。
R3 数字映射：同基底不同数字后缀（WeaponDamage_01 / _02）→ 同译名对应同数字，一一映射；译文中的数字集合必须与 src 完全一致，不得丢失或新增。
R4 合成词推断：驼峰/下划线/无空格拼接的函数风格片段（ApplyAshOfWar、Set_ItemCount）按语素拆解，结合 ctx 与 ref 推断含义后译出；确为多义且无法消歧时，译文后追加半角单引号注释 '可能含义'——这是唯一允许的注释，其余位置零注释、零说明。
R5 叙事腔：肃穆、仪式化；禁口语、禁缩略、禁语气词。
R6 输出契约：只输出 JSON 对象 {"items":[{"i":<行号>,"zh":"<译文>"}]}，输入的每一行都必须有对应条目；zh 内不得残留 terms 中的英文原词，不得附加任何前后缀文字或解释。

[GOLDEN SAMPLES] 逐字模仿以下模式：
WeaponDamage_01 → 武器伤害_01
WeaponDamage_02 → 武器伤害_02
Set_ItemCount_{ID}_{Qty} → 设置物品数量_{ID}_{Qty}
ApplyAshOfWar''Skill/Enhance'' → 应用战灰''技能/强化''
Malenia, Blade of Miquella → 玛莲妮亚,米凯拉的锋刃
HP: 1000, Damage: 50% → HP: 1000, 伤害: 50%（数值与符号原样保留）
Flask of Cerulean Tears +{0} → 蓝露滴圣杯瓶 +{0}"""


# ----------------------------------------------------------------------
def suggestion_from_note(note: str) -> str:
    """从备注列提取词表建议（'…；词表建议参考：蓝露滴圣杯瓶 +{0}'）。"""
    if not note:
        return ""
    idx = note.rfind(_SUGG_MARK)
    return note[idx + len(_SUGG_MARK):].strip() if idx >= 0 else ""


def read_pending(xlsx_path: str | Path) -> list[dict]:
    """读主表待译行：分类 normal 且译文为空。返回行列表（src 唯一）。"""
    wb = load_workbook(str(xlsx_path), read_only=True, data_only=True)
    ws = wb["词条"] if "词条" in wb.sheetnames else wb.active
    pending: list[dict] = []
    seen: set[str] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or len(row) < 7:
            continue
        src = row[2]
        if not src or row[5] != "normal":
            continue
        dst, note = row[3], str(row[6] or "")
        ctx = str(row[7]) if len(row) > 7 and row[7] else ""   # 上下文列可缺省
        src = str(src)
        if str(dst or "").strip() or src in seen:
            continue
        # 带中文的源串不译（增量口径）：zh 基底已汉化的混排词条原样保留
        if _HAS_CJK_SRC.search(src):
            continue
        seen.add(src)
        pending.append({
            "src": src,
            "ph": src.count("{0}"),
            "digits": sorted(_DIG_RUN.findall(src)),
            "ref": suggestion_from_note(note),
            "ctx": ctx,
        })
    wb.close()
    return pending


def build_row_terms(src: str, matcher) -> dict[str, str]:
    """该行实际命中的术语（整词、≥2 字符、译文含中文）+ 官方属性名覆盖。"""
    terms: dict[str, str] = {}
    _, hits = matcher.replace_with_hits(src)
    for en, zh, _s, _e in hits:
        if len(en) >= 2 and any("\u4e00" <= c <= "\u9fff" for c in zh):
            terms[en] = zh
    for en, zh in OFFICIAL_STATS.items():
        if re.search(rf"(?<![A-Za-z]){re.escape(en)}(?![A-Za-z])", src):
            terms[en] = zh   # 官方译名压过词表异值
    return terms


_FN_STYLE = re.compile(r"[a-z][A-Z]|[A-Za-z]_\w")


def _is_fn_style(src: str) -> bool:
    """函数风格片段（驼峰 / 下划线拼接）——R4 语素拆解规则的提示信号。"""
    return bool(_FN_STYLE.search(src))


def _enrich_context(pending: list[dict]) -> int:
    """ctx 空洞时主动补足：数字模板族 / 三词前缀族的邻居词条作同族上下文。

    AI 翻译最大的信息缺口是孤立短词条（无界面位置、无词表命中）；
    同族邻居共享语义场（如 'Golden Rune [1]' 与 'Golden Rune [13]'），
    提供邻居 = 让模型自行类比。返回补足行数。
    """
    fam: dict[str, list[int]] = {}
    for idx, r in enumerate(pending):
        if r["ctx"]:
            continue
        tpl = re.sub(r"\d+", "{0}", r["src"])
        fam.setdefault("T:" + tpl, []).append(idx)
        words = r["src"].split()
        if len(words) >= 3:
            fam.setdefault("P:" + " ".join(words[:3]), []).append(idx)
    filled = 0
    for idxs in fam.values():
        if len(idxs) < 2:
            continue
        for i in idxs:
            sibs = [pending[j]["src"] for j in idxs if j != i][:3]
            if sibs:
                pending[i]["ctx"] = "同族词条：" + " | ".join(sibs)
                filled += 1
    return filled


# ----------------------------------------------------------------------
def _validate(row: dict, zh: str, row_terms: dict[str, str]) -> str | None:
    """返回 None = 通过；否则返回拒绝原因。宁缺勿错。"""
    if not zh or not zh.strip():
        return "空译文"
    zh = zh.strip()
    if zh == row["src"]:
        return "译文与原文相同"
    if zh.count("{0}") != row["ph"]:
        return f"占位符数不符（需 {row['ph']}）"
    if sorted(_DIG_RUN.findall(zh)) != row["digits"]:
        return f"数字不保真（源 {row['digits']} vs 译 {sorted(_DIG_RUN.findall(zh))}）"
    for en, locked in row_terms.items():
        # 恒等锁定（HP→HP、FP→FP 等要求保留原样的术语）不判残留，
        # 否则 "HP 消耗" 这类正确译文必被假阳性拦截
        if en in zh and en != locked:
            return f"术语未译：{en!r} 残留"
    return None


def _http_chat(base_url: str, api_key: str, model: str,
               messages: list[dict], temperature: float) -> str:
    url = base_url.rstrip("/") + "/chat/completions"
    req = urllib.request.Request(url, method="POST",
                                 data=json.dumps({
                                     "model": model,
                                     "messages": messages,
                                     "temperature": temperature,
                                     "response_format": {"type": "json_object"},
                                 }).encode("utf-8"),
                                 headers={"Content-Type": "application/json",
                                          "Authorization": f"Bearer {api_key}"})
    with urllib.request.urlopen(req, timeout=300) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return data["choices"][0]["message"]["content"]


def _parse_items(content: str, batch_n: int) -> dict[int, str]:
    data = json.loads(content)
    items = data.get("items") if isinstance(data, dict) else data
    out: dict[int, str] = {}
    for it in items or []:
        try:
            i, zh = int(it["i"]), str(it["zh"])
        except (KeyError, TypeError, ValueError):
            continue
        if 0 <= i < batch_n:
            out[i] = zh
    return out


def _mock_batch(batch: list[dict]) -> dict[int, str]:
    """离线演练：前缀标记（{0} 与数字天然保真；术语残留会被校验拦截——顺便演示校验）。"""
    return {pos: "[AI] " + r["src"] for pos, r in enumerate(batch)}


# ----------------------------------------------------------------------
def translate(xlsx_path: str | Path, out_json: str | Path, xlsx_out: str | Path,
              api_key: str = "", base_url: str = DEFAULT_BASE_URL,
              model: str = DEFAULT_MODEL, batch_size: int = 30,
              temperature: float = 0.2, dry_run: bool = False,
              user_terms: str | Path | None = None,
              include_conflicts: bool = False, max_workers: int = 5,
              progress_cb=None, should_cancel=None) -> dict:
    """主入口：提取待译行 → 上下文补全 → 并发分批翻译 → 校验 → 填回 + 导出术语表。

    max_workers: 并发批次数（API 吞吐的主要杠杆；dry-run 恒为串行）。
    should_cancel: 协作式取消探针（无参 callable，每完成一批检查一次）。
        取消语义：已完成批次照常入 checkpoint（重跑自动续传，零浪费）；
        尚未开始的批次直接取消不调 API；已在途批次等自然返回后丢弃结果。
    返回 {pending, done_before, translated, flagged, batches, ctx_filled,
          glossary_out, out_json, xlsx_out, cancelled}。
    """
    from .glossary import Glossary
    from .matcher import LongestMatcher

    pending = read_pending(xlsx_path)
    g = Glossary.from_defaults(user_terms=user_terms,
                               include_conflicts=include_conflicts)
    matcher = LongestMatcher(list(g.filtered_items()))
    for idx, r in enumerate(pending):
        r["i"] = idx
        r["terms"] = build_row_terms(r["src"], matcher)
    ctx_filled = _enrich_context(pending)

    # checkpoint 断点续跑：崩溃后重跑自动跳过已完成行。
    # dry-run 绝不读/写 checkpoint——mock 译文一旦混入 done，真实运行会把
    # "[AI] 原文" 当成合格译文直接填表（污染源，实测复现）。
    progress_path = Path(out_json).with_suffix(".ai_progress.json")
    done: dict[str, str] = {}
    if not dry_run and progress_path.exists():
        done = json.loads(progress_path.read_text(encoding="utf-8"))
    done_before = len(done)

    todo = [r for r in pending if r["src"] not in done]
    flagged: list[dict] = []
    batches = 0

    def run_batch(chunk: list[dict]) -> tuple[dict[str, str], list[dict]]:
        """单批：请求 + 校验。返回 (合格译文, 拦截记录)；不改共享状态。"""
        # payload 用批内局部索引（0..len(chunk)-1）：全局索引会让
        # _parse_items 的批界校验丢弃第 2 批之后的全部返回
        payload = [{"i": pos, "src": r["src"],
                    **({"style": "function-like"} if _is_fn_style(r["src"]) else {}),
                    **({"ref": r["ref"]} if r["ref"] else {}),
                    **({"terms": r["terms"]} if r["terms"] else {}),
                    **({"ctx": r["ctx"]} if r["ctx"] else {})}
                    for pos, r in enumerate(chunk)]
        zh_map: dict[int, str] = {}
        last_err = ""
        for attempt in range(3):
            try:
                if dry_run:
                    zh_map = _mock_batch(chunk)
                    break
                content = _http_chat(base_url, api_key, model,
                                     [{"role": "system", "content": _SYSTEM_PROMPT},
                                      {"role": "user",
                                       "content": json.dumps(payload, ensure_ascii=False)}],
                                     temperature)
                zh_map = _parse_items(content, len(chunk))
                if len(zh_map) >= len(chunk) * 0.8:   # 少量缺失可接受，逐行留白
                    break
                last_err = f"返回 {len(zh_map)}/{len(chunk)} 条"
            except (urllib.error.URLError, KeyError, ValueError, json.JSONDecodeError) as e:
                last_err = str(e)[:120]
            if attempt < 2 and not dry_run:
                time.sleep(2 * (attempt + 1))   # 退避重试，防打点限流
        ok: dict[str, str] = {}
        fl: list[dict] = []
        for pos, r in enumerate(chunk):
            zh = zh_map.get(pos)
            if zh is None:
                fl.append({"src": r["src"], "reason": f"API 未返回（{last_err}）"})
                continue
            reason = _validate(r, zh, r["terms"])
            if reason:
                fl.append({"src": r["src"], "zh": zh, "reason": reason})
            else:
                ok[r["src"]] = zh.strip()
        return ok, fl

    chunks = [todo[i:i + batch_size] for i in range(0, len(todo), batch_size)]

    def settle(ok: dict[str, str], fl: list[dict]):
        """主线程合并批结果：共享状态（done/flagged/checkpoint）只在此更新。"""
        nonlocal batches
        done.update(ok)
        flagged.extend(fl)
        batches += 1
        if not dry_run:   # dry-run 不落 checkpoint（防 mock 污染真实续跑）
            # 原子写：直接覆盖时崩溃在写入中途会留下截断 JSON，
            # 下次续跑 json.loads 抛错、断点全部作废
            tmp = progress_path.with_suffix(".json.tmp")
            tmp.write_text(
                json.dumps(done, ensure_ascii=False, indent=1), encoding="utf-8")
            tmp.replace(progress_path)
        if progress_cb:
            progress_cb(len(done), len(pending))

    cancelled = False
    if dry_run or max_workers <= 1:
        for chunk in chunks:
            settle(*run_batch(chunk))
            if should_cancel and should_cancel():
                cancelled = True
                break
    else:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        ex = ThreadPoolExecutor(max_workers=max_workers)
        try:
            futs = [ex.submit(run_batch, c) for c in chunks]
            for fut in as_completed(futs):
                settle(*fut.result())
                if should_cancel and should_cancel():
                    cancelled = True
                    break
        finally:
            if cancelled:
                # 排队未启动的批次直接撤单（不浪费 API）；在途批次自然跑完，
                # 结果不并入 done——行保持待译，下次续跑重做
                ex.shutdown(wait=False, cancel_futures=True)
            else:
                ex.shutdown(wait=True)

    # 合格译文填回 xlsx 副本（绝不覆盖原表），备注标注来源
    _fill_xlsx(xlsx_path, xlsx_out, done)
    # out_json 也走原子写：与 checkpoint 同一哲学（截断 JSON 会毁掉下游合并链路）
    out_json_p = Path(out_json)
    _tmp = out_json_p.with_suffix(".json.tmp")
    _tmp.write_text(json.dumps(done, ensure_ascii=False, indent=1), encoding="utf-8")
    _tmp.replace(out_json_p)
    # 拦截报告反映"最近一轮"的结果：本轮有拦截则覆盖；本轮零拦截且正常跑完
    # 则删掉上一轮遗留文件——否则 ⑤ 定夺窗口会拿过期条目让人重复审已译内容
    # （取消时在途批次的拦截记录被丢弃、对应行仍在待译，保留旧文件反而更接近真相）
    flagged_path = Path(out_json).with_suffix(".flagged.json")
    if flagged:
        flagged_path.write_text(
            json.dumps(flagged, ensure_ascii=False, indent=1), encoding="utf-8")
    elif not cancelled and flagged_path.exists():
        flagged_path.unlink()
    # 原文→译文导出用户术语表（人工校对后可直接 --user-terms 复用）；
    # dry-run 的 mock 译文不落术语表，防污染
    glossary_out = ""
    if done and not dry_run:
        glossary_out = str(Path(out_json).with_suffix(".glossary.tsv"))
        _export_glossary_tsv(glossary_out, done)
    return {"pending": len(pending), "done_before": done_before,
            "translated": len(done) - done_before, "flagged": len(flagged),
            "batches": batches, "ctx_filled": ctx_filled,
            "glossary_out": glossary_out,
            "out_json": str(out_json), "xlsx_out": str(xlsx_out),
            "cancelled": cancelled}


def _export_glossary_tsv(path: str | Path, translations: dict[str, str]):
    """原文→译文写成两列 TSV 用户术语表（Glossary._load_tsv_dict 直接可读）。

    源/译中的制表符与换行降级为空格（TSV 单行语义），其余字符原样保留。
    """
    def _clean(s: str) -> str:
        return s.replace("\t", " ").replace("\r", " ").replace("\n", " ")
    with open(path, "w", encoding="utf-8", newline="") as f:
        for src in sorted(translations):
            f.write(f"{_clean(src)}\t{_clean(translations[src])}\n")


def _fill_xlsx(xlsx_in: str | Path, xlsx_out: str | Path, translations: dict[str, str]):
    """把 AI 译文填入副本的译文列（只填空行），备注追加来源标注。"""
    from shutil import copy2
    copy2(xlsx_in, xlsx_out)
    wb = load_workbook(str(xlsx_out))
    ws = wb["词条"] if "词条" in wb.sheetnames else wb.active
    n = 0
    for row in ws.iter_rows(min_row=2):
        src = row[2].value
        if src and src in translations and not (row[3].value or "").strip():
            row[3].value = translations[src]
            note = row[6].value or ""
            row[6].value = (note + "；" if note else "") + "AI 译文（待校对）"
            n += 1
    wb.save(str(xlsx_out))
    return n
