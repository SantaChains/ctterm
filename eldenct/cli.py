# -*- coding: utf-8 -*-
"""命令行入口。

用法
----
eldenct extract      <CT> [-o terms.json]
eldenct export       terms.json [-o terms.xlsx] [--glossary] [--skip-param]
eldenct import       terms.xlsx [-o translations.json]
eldenct apply        <CT> translations.json -o <新CT> [--glossary] [--fallback]
eldenct cover        <CT> translations.json [-o cover.txt]
eldenct diff         <CT_A> <CT_B> [-o report.txt]
eldenct pipeline     <CT> [-d outdir] [--glossary] [--skip-param]
eldenct delta        <新CT> [-o delta.xlsx] [--zh 旧zhCT] [--en 旧enCT] [--tm 旧json]
eldenct merge        base.json delta.json -o merged.json
eldenct ai-translate terms.xlsx [--dry-run] [--api-key ...] [--model ...]
eldenct add-note     <CT> "注意事项文本" [-o note.CT]
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from .extract import extract_terms
from .glossary import Glossary, is_usable_zh
from .replace import apply_translations
from .diffreport import diff_ct, diff_same_ct, report_txt
from .xlsxio import export_terms_xlsx, import_translations_xlsx
from .incremental import run_delta as _delta, run_merge as _merge
from .aitrans import DEFAULT_BASE_URL, DEFAULT_MODEL

PROJ = Path(__file__).resolve().parent.parent

ALL_TYPES = {"Description", "DropDown", "Lua:LuaComment", "Lua:AAComment",
             "Lua:Caption", "Lua:Msg", "Lua:Guide", "Form:String"}


def _parse_types(spec: str | None) -> set[str] | None:
    """, 分隔的类型开关；None/空 = 全部。未知类型直接报错（防拼写错误静默失效）。"""
    if not spec:
        return None
    parts = {s.strip() for s in spec.split(",") if s.strip()}
    unknown = parts - ALL_TYPES
    if unknown:
        raise ValueError(f"未知类型: {sorted(unknown)}；可选: {sorted(ALL_TYPES)}")
    return parts


def _load_glossary(user_terms: str | None, conflicts: bool) -> Glossary | None:
    return Glossary.from_defaults(user_terms=user_terms, include_conflicts=conflicts)


def _save_json(obj, path: str | Path):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")


def _print_json(obj):
    print(json.dumps(obj, ensure_ascii=False))


def run_extract(ct: str, out: str) -> dict:
    # terms 数组体量大，只返回统计摘要
    return extract_terms(ct, out)["stats"]


def run_export(terms_json: str, out: str, glossary: bool, skip_param: bool,
               user_terms: str | None, conflicts: bool,
               prefill_partial: bool = False,
               types: str | None = None) -> dict:
    terms = json.loads(Path(terms_json).read_text(encoding="utf-8"))["terms"]
    allowed = _parse_types(types)
    auto = {}
    if glossary:
        g = _load_glossary(user_terms, conflicts)
        # 预填只做整串精确命中：已验证的术语表条目直接填译文（专有名词锁定），
        # 未命中的整条留空交给翻译软件，避免中英混合文本干扰机器翻译。
        exact = dict(g.items())
        for t in terms:
            dst = exact.get(t["source"])
            # 自动预填只接受"可用中文译文"：滤掉全角污染/原文照抄等退化条目
            if dst and dst != t["source"] and is_usable_zh(dst):
                auto[t["source"]] = dst
    # 数字模板族合并必须在 partial 建议之前：合并后的模板行
    # （'Golden Rune [{0}]'）才能拿到词表短语建议（'金符文 [{0}]'）
    from .xlsxio import merge_numeric_families
    terms, fam_info = merge_numeric_families(terms, auto)
    contexts = {t["source"]: t.get("contexts", []) for t in terms if t.get("contexts")}
    partial = {}
    if glossary and prefill_partial:
        from .matcher import LongestMatcher
        matcher = LongestMatcher(list(g.filtered_items()))
        for t in terms:
            if t["source"] in auto or t["class"] != "normal":
                continue
            suggested, hits = matcher.replace_with_hits(t["source"])
            # 只保留可用中文建议：全角英文/切词残留（词表退化条目）
            # 对翻译软件是噪音，滤掉。
            if hits and is_usable_zh(suggested):
                partial[t["source"]] = suggested
    info = export_terms_xlsx(terms, out, auto_translations=auto,
                             partial_notes=partial,
                             include_param=not skip_param,
                             contexts=contexts,
                             allowed_types=allowed,
                             merge_families=False)
    return {**info, **fam_info}


def run_import(xlsx: str, out: str) -> dict:
    translations, issues, exported = import_translations_xlsx(xlsx)
    _save_json(translations, out)
    Path(out).with_suffix(".exported.json").write_text(
        json.dumps(sorted(exported), ensure_ascii=False), encoding="utf-8")
    return {"translations": len(translations), "exported": len(exported), "issues": issues}


def run_apply(ct: str, translations_json: str, out: str, glossary: bool,
              fallback: bool, user_terms: str | None, conflicts: bool,
              types: str | None = None) -> dict:
    translations = json.loads(Path(translations_json).read_text(encoding="utf-8"))
    g_items = None
    exported_sources = None
    if glossary:
        g_items = list(_load_glossary(user_terms, conflicts).filtered_items())
    if fallback:
        exported_path = Path(translations_json).with_suffix(".exported.json")
        if exported_path.exists():
            exported_sources = set(json.loads(exported_path.read_text(encoding="utf-8")))
    return apply_translations(ct, translations, out, glossary_items=g_items,
                              fallback_glossary=fallback,
                              exported_sources=exported_sources,
                              allowed_types=_parse_types(types))


def run_cover(ct: str, translations_json: str, out: str | None) -> dict:
    translations = json.loads(Path(translations_json).read_text(encoding="utf-8"))
    rows, stats, unknown = diff_same_ct(ct, translations)
    if out:
        Path(out).write_text(report_txt(rows, title=f"翻译覆盖对照 {ct}"), encoding="utf-8")
    return {**stats, "unknown_preview": unknown[:10]}


def run_diff(a: str, b: str, out: str | None) -> dict:
    rows, stats = diff_ct(a, b)
    if out:
        Path(out).write_text(report_txt(rows, title=f"{a}  vs  {b}"), encoding="utf-8")
    return stats


def run_pipeline(ct: str, outdir: str, glossary: bool, skip_param: bool,
                 user_terms: str | None, conflicts: bool,
                 prefill_partial: bool = False,
                 types: str | None = None) -> dict:
    outdir_p = Path(outdir)
    outdir_p.mkdir(parents=True, exist_ok=True)
    terms_json = outdir_p / "terms.json"
    xlsx = outdir_p / "terms.xlsx"
    run_extract(ct, str(terms_json))
    info = run_export(str(terms_json), str(xlsx), glossary, skip_param,
                      user_terms, conflicts, prefill_partial, types)
    return {"terms_json": str(terms_json), "xlsx": str(xlsx), **info}


def run_delta_cmd(ct: str, out: str, zh: str | None, en: str | None,
                  tm: str | None, glossary: bool, user_terms: str | None,
                  conflicts: bool, prefill_partial: bool = False) -> dict:
    return _delta(ct, out, zh_ct=zh, en_ct=en, tm_json=tm,
                  glossary=glossary, user_terms=user_terms,
                  conflicts=conflicts, prefill_partial=prefill_partial)


def run_merge_cmd(base: str, delta: str, out: str) -> dict:
    return _merge(base, delta, out)


def run_ai_translate(xlsx: str, out: str | None, api_key: str | None,
                     base_url: str, model: str, batch_size: int,
                     user_terms: str | None, dry_run: bool,
                     glossary: bool = False, conflicts: bool = False) -> dict:
    key = api_key or os.environ.get("DEEPSEEK_API_KEY", "")
    if not key and not dry_run:
        raise ValueError("缺少 API Key：设 DEEPSEEK_API_KEY 环境变量或加 --api-key")
    xlsx_p = Path(xlsx)
    out_json = Path(out) if out else xlsx_p.with_name(xlsx_p.stem + "-ai.json")
    xlsx_out = out_json.with_suffix(".xlsx")
    from .aitrans import translate
    return translate(xlsx_p, out_json, xlsx_out, api_key=key,
                     base_url=base_url or DEFAULT_BASE_URL,
                     model=model or DEFAULT_MODEL, batch_size=batch_size,
                     dry_run=dry_run, user_terms=user_terms,
                     include_conflicts=conflicts)


def run_add_note(ct: str, text: str, out: str | None) -> dict:
    from .entryops import add_note_entry
    return add_note_entry(ct, text, out)


# ----------------------------------------------------------------------
# argparse 薄包装
def _add_glossary_opts(p: argparse.ArgumentParser):
    p.add_argument("--glossary", action="store_true", help="加载术语表")
    p.add_argument("--user-terms", default=None, help="用户自定义术语表(tsv/json)")
    p.add_argument("--conflicts", action="store_true", help="启用异值表（有歧义，慎用）")


def build_parser():
    p = argparse.ArgumentParser(prog="eldenct",
                                description="ELDEN RING CT 翻译工具")
    sub = p.add_subparsers(dest="cmd", required=True)

    pe = sub.add_parser("extract", help="解析 CT，提取去重词条")
    pe.add_argument("ct")
    pe.add_argument("-o", "--out", default=str(PROJ / "data" / "terms.json"))
    pe.set_defaults(fn=run_extract)

    px = sub.add_parser("export", help="词条导出为 xlsx（待翻译）")
    px.add_argument("terms_json")
    px.add_argument("-o", "--out", default=str(PROJ / "data" / "terms.xlsx"))
    px.add_argument("--skip-param", action="store_true", help="跳过疑似参数类词条")
    px.add_argument("--prefill-partial", action="store_true",
                    help="短语级词表建议写进备注列（供翻译参考，不填译文列）")
    px.add_argument("--types", default=None,
                    help="类型开关（逗号分隔，如 Description,DropDown,Lua:Caption；默认全部）")
    _add_glossary_opts(px)
    px.set_defaults(fn=run_export)

    pi = sub.add_parser("import", help="导回翻译后的 xlsx")
    pi.add_argument("xlsx")
    pi.add_argument("-o", "--out", default=str(PROJ / "data" / "translations.json"))
    pi.set_defaults(fn=run_import)

    pa = sub.add_parser("apply", help="执行替换，生成新 CT（源文件只读，绝不覆盖）")
    pa.add_argument("ct")
    pa.add_argument("translations_json")
    pa.add_argument("-o", "--out", required=True)
    pa.add_argument("--fallback", action="store_true",
                    help="对 xlsx 未导出过的词条按术语表补漏（默认关闭；用户留空行不受影响）")
    pa.add_argument("--types", default=None,
                    help="类型开关（逗号分隔；未勾选的类型绝不参与替换）")
    _add_glossary_opts(pa)
    pa.set_defaults(fn=run_apply)

    pc = sub.add_parser("cover", help="翻译覆盖对照（乱序无关）")
    pc.add_argument("ct")
    pc.add_argument("translations_json")
    pc.add_argument("-o", "--out", default=None)
    pc.set_defaults(fn=run_cover)

    pd = sub.add_parser("diff", help="两个 CT 按 ID 乱序对照")
    pd.add_argument("a")
    pd.add_argument("b")
    pd.add_argument("-o", "--out", default=None)
    pd.set_defaults(fn=run_diff)

    pp = sub.add_parser("pipeline", help="一条龙：extract + export")
    pp.add_argument("ct")
    pp.add_argument("-d", "--outdir", default=str(PROJ / "data"))
    pp.add_argument("--skip-param", action="store_true")
    pp.add_argument("--prefill-partial", action="store_true",
                    help="短语级词表建议写进备注列（供翻译参考，不填译文列）")
    pp.add_argument("--types", default=None)
    _add_glossary_opts(pp)
    pp.set_defaults(fn=run_pipeline)

    pd_ = sub.add_parser("delta", help="增量导出：新表词条 + TM 继承预填 + 上下文列")
    pd_.add_argument("ct", help="新版本 EN CT")
    pd_.add_argument("-o", "--out", default=str(PROJ / "data" / "delta.xlsx"))
    pd_.add_argument("--zh", default=None, help="旧版已译 CT（TM 来源一）")
    pd_.add_argument("--en", default=None, help="旧版 EN CT（与 --zh 配对做 ID 对齐）")
    pd_.add_argument("--tm", default=None, help="旧 translations.json（TM 来源二）")
    pd_.add_argument("--prefill-partial", action="store_true")
    _add_glossary_opts(pd_)
    pd_.set_defaults(fn=run_delta_cmd)

    pm = sub.add_parser("merge", help="合并译文集（delta 优先，冲突报告）")
    pm.add_argument("base", help="旧 translations.json")
    pm.add_argument("delta", help="delta 导回的 translations.json")
    pm.add_argument("-o", "--out", required=True)
    pm.set_defaults(fn=run_merge_cmd)

    pai = sub.add_parser("ai-translate",
                         help="AI 翻译待译行（默认 DeepSeek；校验后填回 xlsx 副本）")
    pai.add_argument("xlsx", help="terms.xlsx（只读；产物为 <name>-ai.xlsx 副本）")
    pai.add_argument("-o", "--out", default=None, help="译文 JSON 输出路径")
    pai.add_argument("--api-key", default=None, help="缺省读 DEEPSEEK_API_KEY 环境变量")
    pai.add_argument("--base-url", default=DEFAULT_BASE_URL,
                     help="OpenAI 兼容接口地址")
    pai.add_argument("--model", default=DEFAULT_MODEL)
    pai.add_argument("--batch-size", type=int, default=30)
    pai.add_argument("--dry-run", action="store_true",
                     help="离线演练：不调 API，用假译文走完整校验/回填链路")
    _add_glossary_opts(pai)
    pai.set_defaults(fn=run_ai_translate)

    pn = sub.add_parser("add-note", help="追加注意事项词条（GroupHeader 纯文本行，ID 自动唯一）")
    pn.add_argument("ct")
    pn.add_argument("text", help="注意事项文本")
    pn.add_argument("-o", "--out", default=None,
                    help="输出路径（默认 <源名>-note.CT；源文件只读）")
    pn.set_defaults(fn=run_add_note)

    return p


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args(argv)
    fn, params = args.fn, vars(args)
    params.pop("cmd", None)
    params.pop("fn", None)
    try:
        result = fn(**params)
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1
    if isinstance(result, dict):
        _print_json(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
