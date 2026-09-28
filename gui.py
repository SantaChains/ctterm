# -*- coding: utf-8 -*-
"""法环 CT 翻译工具 — 可视化界面（Tkinter，标准库实现，无外部依赖）。

布局（顶部路径栏 + 主流程步骤条 / 左设置栏 + 中预览区 + 右功能页签 / 底部状态栏）：
  路径栏：选择 CT（路径 / 窗口几何 / 类型开关状态均记忆于 gui_config.json）
  主流程：① 提取词条 → ② 导出 XLSX → ③ AI 翻译 → ④ 导入译文 ▾ → ⑤ 批量替换
          步骤按钮按前置完成情况自动启用/置灰，始终知道下一步点哪
  左设置栏：类型开关（2 列 × 4 行，状态记忆）、最小提取长度、TXT 预览导出、
            字典状态与另存
  预览区：占满剩余高度，分批渲染（万级行不冻结），横向滚动
  功能页签：对照（DIFF/COVER）| 工具（多行注意事项）| 设置（AI 连接参数）
  状态栏：进度条 + 分阶段状态文字

批量替换生成新 CT（源文件只读，绝不覆盖；.bak 备份在首次提取时生成）。
所有耗时操作在后台线程执行，主界面不卡顿；进度条分阶段推进。
Windows 高分屏声明 DPI 感知，字体随系统缩放保持清晰。
"""
from __future__ import annotations

import difflib
import json
import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

PROJ = Path(__file__).resolve().parent
CONFIG = PROJ / "gui_config.json"

TYPE_LABELS = [
    ("Description", "Description"),
    ("DropDown", "DropDownList"),
    ("Lua:LuaComment", "Lua 注释"),
    ("Lua:AAComment", "AA 注释"),
    ("Lua:Caption", "Lua 标题"),
    ("Lua:Msg", "Lua 消息"),
    ("Lua:Guide", "Lua 引导"),
    ("Form:String", "表单字符串"),
]

PREVIEW_BATCH = 800   # 预览分批渲染行数（一次性插 1.5 万行会冻结界面）


def _load_cfg() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_cfg(cfg: dict):
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")


def _enable_high_dpi(root: tk.Tk):
    """Windows 高分屏清晰化：声明 DPI 感知并让 Tk 字号跟随系统缩放。

    不声明时系统按位图拉伸整窗——125%/150% 缩放下全部发虚。
    """
    if os.name != "nt":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            return
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem()
        if dpi:
            root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("法环 CT 翻译工具 — eldenct")
        self.cfg = _load_cfg()
        root.geometry(self.cfg.get("geometry", "1280x800"))
        root.minsize(1024, 640)
        self.terms: list[dict] = []          # 提取结果
        self.model = None
        self.ct_path: str | None = None
        self.dict_map: dict[str, str] = {}   # 合并后的本地字典
        self.last_xlsx: str | None = None

        self._build_ui()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        last = self.cfg.get("last_ct")
        if last and Path(last).exists():
            self.var_ct.set(last)

    # ------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 6, "pady": 4}

        # ---- 顶部：文件路径栏 ----
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        ttk.Label(top, text="CT 文件:").pack(side="left")
        self.var_ct = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_ct).pack(side="left", fill="x",
                                                      expand=True, padx=4)
        ttk.Button(top, text="选择 CT…", command=self.pick_ct).pack(side="left", padx=2)
        ttk.Label(top, text="首次处理自动生成 .bak 备份", foreground="#888").pack(
            side="left", padx=4)

        # ---- 主流程步骤条（按前置完成情况自动启停）----
        steps = ttk.LabelFrame(self.root, text="主流程（按序执行）")
        steps.pack(fill="x", **pad)
        self.btn_extract = ttk.Button(steps, text="① 提取词条", command=self.do_extract)
        self.btn_export = ttk.Button(steps, text="② 导出 XLSX",
                                     command=self.do_export_xlsx, state="disabled")
        self.btn_ai = ttk.Button(steps, text="③ AI 翻译",
                                 command=self.do_ai_translate, state="disabled")
        self.btn_import = ttk.Menubutton(steps, text="④ 导入译文 ▾")
        menu = tk.Menu(self.btn_import, tearoff=0)
        menu.add_command(label="导入 XLSX（正典 8 列，含哈希校验）", command=self.do_import_xlsx)
        menu.add_command(label="导入 TXT（原文/译文交替行）", command=self.do_import_txt)
        menu.add_command(label="导入 TSV 字典（EN→ZH 两列）", command=self.do_import_tsv)
        menu.add_separator()
        menu.add_command(label="清除字典（无字典全翻译）", command=self.do_clear_dict)
        self.btn_import.config(menu=menu)
        self.btn_apply = ttk.Button(steps, text="⑤ 批量替换 → 新文件",
                                    command=self.do_apply, state="disabled")
        for i, b in enumerate((self.btn_extract, self.btn_export, self.btn_ai,
                               self.btn_import, self.btn_apply)):
            b.pack(side="left", fill="x", expand=True, padx=(6 if i == 0 else 3, 3))
        ttk.Label(steps, text="字典在「导出 XLSX」之前导入即自动预填（用户字典优先于内置词表）",
                  foreground="#888").pack(anchor="w", padx=8, pady=(0, 2))

        # ---- 主区：左设置栏 | 中预览 | 右功能页签 ----
        main = ttk.Panedwindow(self.root, orient="horizontal")
        main.pack(fill="both", expand=True, **pad)

        # 左：类型开关 + 参数 + 字典状态
        left = ttk.Frame(main)
        main.add(left, weight=0)
        lf_types = ttk.LabelFrame(left, text="类型开关（导出与替换严格受控）")
        lf_types.pack(fill="x", padx=(4, 2), pady=4)
        self.type_vars = {}
        saved_types = self.cfg.get("types", {})
        for idx, (key, label) in enumerate(TYPE_LABELS):
            v = tk.BooleanVar(value=saved_types.get(key, key in ("Description", "DropDown")))
            self.type_vars[key] = v
            ttk.Checkbutton(lf_types, text=label, variable=v,
                            command=self._save_types).grid(
                row=idx // 2, column=idx % 2, sticky="w", padx=6, pady=1)

        lf_param = ttk.LabelFrame(left, text="参数")
        lf_param.pack(fill="x", padx=(4, 2), pady=4)
        row = ttk.Frame(lf_param)
        row.pack(fill="x", padx=6, pady=2)
        ttk.Label(row, text="最小提取长度:").pack(side="left")
        self.var_minlen = tk.IntVar(value=self.cfg.get("min_len", 2))
        ttk.Spinbox(row, from_=1, to=50, width=4, textvariable=self.var_minlen).pack(
            side="left", padx=4)
        ttk.Button(row, text="导出 TXT 预览", command=self.do_export_txt).pack(
            side="left", padx=6)

        lf_dict = ttk.LabelFrame(left, text="字典")
        lf_dict.pack(fill="x", padx=(4, 2), pady=4)
        self.lbl_dict = ttk.Label(lf_dict, text="本地字典：0 条", font=("", 10, "bold"))
        self.lbl_dict.pack(anchor="w", padx=8, pady=(2, 0))
        row_ds = ttk.Frame(lf_dict)
        row_ds.pack(fill="x", padx=4, pady=2)
        ttk.Button(row_ds, text="另存字典…", command=self._save_dict_as).pack(
            side="left", padx=2)
        ttk.Button(row_ds, text="清除字典", command=self.do_clear_dict).pack(
            side="left", padx=2)
        ttk.Label(lf_dict, text="导入走主流程 ④；导入/修改后自动保存到\n上次使用的字典文件（默认 dict.txt）",
                  foreground="#888", justify="left").pack(anchor="w", padx=6, pady=(0, 2))

        # 中：预览（占满剩余高度）
        prev = ttk.LabelFrame(main, text="预览（原文 / 类型 / 出现次数 / 长度 / 上下文）")
        main.add(prev, weight=1)
        cols = ("src", "type", "line", "len", "ctx")
        self.tree = ttk.Treeview(prev, columns=cols, show="headings")
        for c, (w, t) in zip(cols, [(360, "原文"), (110, "类型"), (70, "出现次数"),
                                    (50, "长度"), (240, "上下文")]):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(prev, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(prev, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, columnspan=2, sticky="ew")
        prev.rowconfigure(0, weight=1)
        prev.columnconfigure(0, weight=1)

        # 右：功能页签（低频操作收纳）
        nb = ttk.Notebook(main)
        main.add(nb, weight=0)

        tab_cmp = ttk.Frame(nb)
        nb.add(tab_cmp, text=" 对照 ")
        row_ct2 = ttk.Frame(tab_cmp)
        row_ct2.pack(fill="x", padx=6, pady=(8, 2))
        self.var_ct2 = tk.StringVar(value=self.cfg.get("last_ct2", ""))
        ttk.Entry(row_ct2, textvariable=self.var_ct2).pack(
            side="left", fill="x", expand=True)
        ttk.Button(row_ct2, text="选文件…", command=self.pick_ct2).pack(side="left", padx=2)
        row_cmp = ttk.Frame(tab_cmp)
        row_cmp.pack(fill="x", padx=6, pady=4)
        ttk.Button(row_cmp, text="DIFF 两 CT", command=self.do_diff).pack(
            side="left", padx=2, expand=True, fill="x")
        ttk.Button(row_cmp, text="COVER 覆盖", command=self.do_cover).pack(
            side="left", padx=2, expand=True, fill="x")
        ttk.Label(tab_cmp, text="DIFF：选另一个 CT（跨版本/中英对照，乱序无关，按 ID+字段对齐）。\n"
                               "COVER：选 translations.json（翻译覆盖核验）。\n"
                               "结果在可视化双栏窗口打开，报告可导出。",
                  foreground="#888", justify="left").pack(anchor="w", padx=8, pady=4)

        tab_tool = ttk.Frame(nb)
        nb.add(tab_tool, text=" 工具 ")
        ttk.Button(tab_tool, text="追加注意事项词条…", command=self.do_add_note).pack(
            fill="x", padx=8, pady=(8, 2))
        ttk.Label(tab_tool, text="在弹出的多行输入框写注意事项，追加为根级纯文本词条\n"
                               "（GroupHeader，ID 自动唯一），输出 <源名>-note.CT，源文件只读。",
                  foreground="#888", justify="left").pack(anchor="w", padx=8, pady=4)

        tab_cfg = ttk.Frame(nb)
        nb.add(tab_cfg, text=" 设置 ")
        lf_ai = ttk.LabelFrame(tab_cfg, text="AI 连接（DeepSeek 默认，任意 OpenAI 兼容接口）")
        lf_ai.pack(fill="x", padx=6, pady=(8, 2))
        row_ai = ttk.Frame(lf_ai)
        row_ai.pack(fill="x", padx=6, pady=2)
        ttk.Label(row_ai, text="API Key:").pack(side="left")
        self.var_apikey = tk.StringVar(value=self.cfg.get("ai_api_key", ""))
        ttk.Entry(row_ai, textvariable=self.var_apikey, show="*").pack(
            side="left", fill="x", expand=True, padx=4)
        row_ai2 = ttk.Frame(lf_ai)
        row_ai2.pack(fill="x", padx=6, pady=2)
        ttk.Label(row_ai2, text="模型:").pack(side="left")
        self.var_model = tk.StringVar(value=self.cfg.get("ai_model", "deepseek-chat"))
        ttk.Entry(row_ai2, textvariable=self.var_model, width=16).pack(side="left", padx=4)
        ttk.Label(row_ai2, text="接口:").pack(side="left")
        self.var_baseurl = tk.StringVar(
            value=self.cfg.get("ai_base_url", "https://api.deepseek.com"))
        ttk.Entry(row_ai2, textvariable=self.var_baseurl).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Label(lf_ai, text="校验不合格的行自动留白并写 flagged 报告，绝不带病替换；\n"
                             "Key 明文保存在本机 gui_config.json，仅本机使用。",
                  foreground="#888", justify="left").pack(anchor="w", padx=6, pady=(0, 2))

        # ---- 底部：进度条 + 状态栏 ----
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=(0, 2))
        self._ai_running = False   # AI 翻译进行中标志（驱动秒级心跳）
        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.var_status, anchor="w").pack(
            fill="x", padx=8, pady=2)

        self._refresh_step_states()

    # ------------------------------------------------------------------
    def _on_close(self):
        self.cfg["geometry"] = self.root.geometry()
        _save_cfg(self.cfg)
        self.root.destroy()

    def _save_types(self):
        self.cfg["types"] = {k: bool(v.get()) for k, v in self.type_vars.items()}
        _save_cfg(self.cfg)

    def _refresh_step_states(self):
        """主流程步骤按前置完成情况启停：②③需已提取，③另需已导出 XLSX，⑤需字典非空。"""
        has_terms = bool(self.terms)
        self.btn_export.config(state="normal" if has_terms else "disabled")
        self.btn_ai.config(
            state="normal" if (has_terms and self.last_xlsx) else "disabled")
        self.btn_apply.config(
            state="normal" if (has_terms and self.dict_map) else "disabled")

    # ------------------------------------------------------------------
    def pick_ct(self):
        p = filedialog.askopenfilename(
            title="选择 CT 文件", filetypes=[("Cheat Table", "*.ct"), ("所有文件", "*.*")])
        if p:
            self.var_ct.set(p)
            self.cfg["last_ct"] = p
            _save_cfg(self.cfg)

    def _selected_types(self) -> set[str]:
        return {k for k, v in self.type_vars.items() if v.get()}

    def _run_bg(self, fn, done):
        """后台线程执行，完成后主线程回调；进度条脉冲。"""
        self.progress.start(12)
        self.var_status.set("处理中…")

        def worker():
            try:
                result = fn()
                err = None
            except Exception as e:      # noqa: BLE001
                result, err = None, e
            self.root.after(0, lambda: self._done(done, result, err))
        threading.Thread(target=worker, daemon=True).start()

    def _done(self, done, result, err):
        self.progress.stop()
        self._ai_running = False   # 复位 AI 翻译心跳（无论成败）
        if err:
            messagebox.showerror("错误", str(err))
            self.var_status.set(f"失败：{err}")
            return
        done(result)

    # ------------------------------------------------------------------
    def _show_rows(self, rows, final_status: str):
        """预览分批渲染：万级行一次性 insert 会冻结界面数秒。"""
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._pending_rows = rows
        self._row_pos = 0
        self._final_status = final_status
        self._render_batch()

    def _render_batch(self):
        rows = self._pending_rows
        for r in rows[self._row_pos: self._row_pos + PREVIEW_BATCH]:
            self.tree.insert("", "end", values=(
                r["src"][:120], r["type"], f"×{r['count']}",
                len(r["src"]), r["ctx"][:80]))
        self._row_pos += PREVIEW_BATCH
        if self._row_pos < len(rows):
            self.var_status.set(f"预览加载中… {self._row_pos}/{len(rows)}")
            self.root.after(1, self._render_batch)
        else:
            self.var_status.set(self._final_status)

    def do_extract(self):
        ct = self.var_ct.get().strip()
        if not ct or not Path(ct).exists():
            messagebox.showwarning("提示", "请先选择有效的 CT 文件")
            return
        minlen = self.var_minlen.get()
        types = self._selected_types()

        def work():
            from eldenct.extract import extract_terms
            from eldenct.replace import apply_translations  # noqa: F401 预热导入
            # .bak 备份（副本在处理前立即生成）
            bak = Path(ct).with_suffix(Path(ct).suffix + ".bak")
            if not bak.exists():
                shutil.copy2(ct, bak)
            # 与 CLI 同一条提取路径：去重 + 分类（param 标注）+ 无字母剔除
            res = extract_terms(ct)
            out = []
            for t in res["terms"]:
                tset = set(t["type"])
                if not (tset & types):
                    continue
                if len(t["source"].strip()) < minlen:
                    continue
            out.append({"src": t["source"], "type": "+".join(t["type"]),
                        "count": t["count"], "cls": t["class"],
                        "ctxs": t.get("contexts", []),
                        "ctx": " > ".join(t.get("contexts", [])[:2])})
            return ct, out, dict(res["stats"])

        def done(res):
            ct, rows, stats = res
            self.ct_path = ct
            self.terms = rows
            cd = stats.get("class_dist", {})
            final = (f"提取完成：{len(rows)} 条唯一词条"
                     f"（normal {cd.get('normal', 0)} / param {cd.get('param', 0)}，"
                     f"无字母已剔除 {stats.get('skipped_no_letter', 0)}；备份 .bak 已确认）")
            self._show_rows(rows, final)

        self._run_bg(work, done)

    # ------------------------------------------------------------------
    def do_export_xlsx(self):
        if not self.terms:
            messagebox.showwarning("提示", "请先提取词条")
            return
        types = self._selected_types()
        if not types:
            messagebox.showwarning("提示", "类型开关全部未勾选：导出将为空表。请至少勾选一类。")
            return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                           initialfile="terms.xlsx",
                                           filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        dict_snapshot = dict(self.dict_map)   # 主线程快照，杜绝后台读竞争
        terms_snapshot = list(self.terms)     # 同一时刻的词条快照

        def work():
            from eldenct.xlsxio import export_terms_xlsx
            from eldenct.glossary import Glossary, is_usable_zh
            # 与 CLI 同一条导出路径：分类/副表/术语表预填/数字模板族合并
            terms = [{"source": r["src"], "type": r["type"].split("+"),
                      "count": r["count"], "class": r["cls"],
                      "contexts": r["ctxs"], "entries": []} for r in terms_snapshot]
            # 预填优先级：用户字典（TSV/XLSX/TXT 导入）> 内置默认词表；
            # 字典为空 = 无字典全翻译（预填零条）
            exact = dict(Glossary.from_defaults().items())
            exact.update(dict_snapshot)
            auto = {}
            for t in terms:
                dst = exact.get(t["source"])
                if dst and dst != t["source"] and is_usable_zh(dst):
                    auto[t["source"]] = dst
            contexts = {t["source"]: t["contexts"] for t in terms if t["contexts"]}
            return export_terms_xlsx(terms, out, auto_translations=auto,
                                     contexts=contexts)

        def done(info):
            self.last_xlsx = out
            self._refresh_step_states()
            messagebox.showinfo(
                "完成", f"已导出：{out}\n\n主表 {info['rows']} 行"
                        f"（术语表预填 {info['auto']} / 待翻译 {info['manual']}）\n"
                        f"副表参数保留 {info['param_kept']} 行\n"
                        f"数字模板合并：{info.get('families', 0)} 族"
                        f"（省 {info.get('variants_merged', 0)} 行）")

        self._run_bg(work, done)

    # ------------------------------------------------------------------
    def do_ai_translate(self):
        if not self.last_xlsx:
            messagebox.showwarning("提示", "请先导出 XLSX")
            return
        # Tk 变量只能在主线程读取——此处取值后以闭包传入后台
        api_key = self.var_apikey.get().strip()
        if not api_key:
            messagebox.showwarning("提示", "请先填写 API Key（DeepSeek 等 OpenAI 兼容接口）")
            return
        model = self.var_model.get().strip() or "deepseek-chat"
        base_url = self.var_baseurl.get().strip() or "https://api.deepseek.com"
        # 记住填写的连接信息（Key 明文存本机 gui_config.json，仅本机使用）
        self.cfg.update(ai_api_key=api_key, ai_model=model, ai_base_url=base_url)
        _save_cfg(self.cfg)
        xlsx_in = self.last_xlsx
        out_json = Path(xlsx_in).with_name(Path(xlsx_in).stem + "-ai.json")

        # 进度可见性：进度条按条数推进 + 秒级心跳（首批返回前界面上不能是死的）
        self._ai_running = True
        t0 = time.time()

        def tick():
            if not self._ai_running:
                return
            self.var_status.set(
                f"AI 翻译中… 已提交模型，等待批次返回（已等待 {int(time.time() - t0)} 秒；"
                f"单批 30 条，接口超时上限 5 分钟，首批返回后每批更新进度条）")
            self.root.after(1000, tick)

        def work():
            from eldenct.aitrans import translate, read_pending
            total = len(read_pending(xlsx_in))

            def init_ui():
                self.progress.stop()   # 关闭 _run_bg 的自动步进，改为真实条数
                self.progress.configure(maximum=max(total, 1), value=0)
            self.root.after(0, init_ui)

            def progress(done_n, total_n):
                def upd():
                    pct = done_n * 100 // max(total_n, 1)
                    self.var_status.set(f"AI 翻译中… {done_n}/{total_n} 条（{pct}%）")
                    self.progress.configure(value=done_n, maximum=max(total_n, 1))
                self.root.after(0, upd)
            return translate(xlsx_in, out_json,
                             Path(xlsx_in).with_name(Path(xlsx_in).stem + "-ai.xlsx"),
                             api_key=api_key, base_url=base_url, model=model,
                             progress_cb=progress)

        def done(r):
            self.var_status.set(
                f"AI 翻译完成：{r['translated']}/{r['pending']} 条，"
                f"校验拦截 {r['flagged']} 条（详见 {Path(r['out_json']).with_suffix('.flagged.json').name}）")
            gl = f"\n术语表：{r['glossary_out']}" if r.get("glossary_out") else ""
            if messagebox.askyesno(
                    "AI 翻译完成",
                    f"合格 {r['translated']} 条已填入：{r['xlsx_out']}\n"
                    f"上下文补全 {r.get('ctx_filled', 0)} 行\n"
                    f"拦截 {r['flagged']} 条（留白待人工）{gl}\n\n"
                    f"是否立即导入该副本作为译文来源？"):
                self._import_ai_result(r["xlsx_out"])

        self.root.after(0, tick)
        self._run_bg(work, done)

    def _import_ai_result(self, xlsx_path):
        def work():
            from eldenct.xlsxio import import_translations_xlsx
            translations, issues, _e = import_translations_xlsx(xlsx_path)
            return translations, len(issues)

        def done(res):
            translations, n_issue = res
            self.dict_map.update(translations)   # 主线程更新共享字典
            self._dict_done(len(translations), n_issue)

        self._run_bg(work, done)

    def do_export_txt(self):
        if not self.terms:
            messagebox.showwarning("提示", "请先提取词条")
            return
        out = filedialog.asksaveasfilename(defaultextension=".txt",
                                           initialfile="terms.txt")
        if not out:
            return

        def work():
            Path(out).write_text(
                "\n".join(r["src"] for r in self.terms), encoding="utf-8")
            return out

        self._run_bg(work, lambda o: messagebox.showinfo("完成", f"已导出：{o}"))

    # ------------------------------------------------------------------
    def do_import_xlsx(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not p:
            return

        def work():
            # 正典读取路径：8 列主表（C=源文本 D=译文，含哈希校验）
            from eldenct.xlsxio import import_translations_xlsx
            translations, issues, _exported = import_translations_xlsx(p)
            return translations, len(issues)

        def done(res):
            translations, n_issue = res
            self.dict_map.update(translations)   # 主线程更新共享字典
            self._dict_done(len(translations), n_issue)

        self._run_bg(work, done)

    def do_import_txt(self):
        p = filedialog.askopenfilename(filetypes=[("文本", "*.txt")])
        if not p:
            return

        def work():
            merged: dict[str, str] = {}
            with open(p, encoding="utf-8-sig") as f:
                lines = [l.rstrip("\r\n") for l in f]
            for i in range(0, len(lines) - 1, 2):
                src, dst = lines[i], lines[i + 1]
                if src.strip() and dst.strip():
                    merged[src] = dst
            return merged

        def done(merged):
            self.dict_map.update(merged)   # 主线程更新共享字典
            self._dict_done(len(merged))

        self._run_bg(work, done)

    def do_import_tsv(self):
        """TSV 术语字典（EN\\tZH 两列，QUOTE_NONE——行首引号是词条的一部分）。"""
        p = filedialog.askopenfilename(
            title="选择 TSV 字典", filetypes=[("TSV 字典", "*.tsv *.txt"), ("所有文件", "*.*")])
        if not p:
            return

        def work():
            import csv
            d: dict[str, str] = {}
            with open(p, encoding="utf-8-sig", newline="") as f:
                for row in csv.reader(f, delimiter="\t", quoting=csv.QUOTE_NONE):
                    # 源列不 strip（词条可能含特殊空白）；译文列剔除首尾空白
                    if len(row) >= 2 and row[0] and row[1].strip():
                        d[row[0]] = row[1].strip()
            return d

        def done(d):
            self.dict_map.update(d)   # 主线程更新共享字典
            self._dict_done(len(d))

        self._run_bg(work, done)

    def do_clear_dict(self):
        """清空字典：无字典模式 = 替换仅用导入的译文，无术语兜底。"""
        if not self.dict_map:
            messagebox.showinfo("提示", "字典已是空")
            return
        if not messagebox.askyesno(
                "清除字典", f"确定清空当前 {len(self.dict_map)} 条字典？\n\n"
                            "清空后（无字典模式）：\n"
                            "· 导出 XLSX 无预填，全部待译行交给翻译/AI\n"
                            "· 批量替换仅使用之后导入的译文"):
            return
        self.dict_map.clear()
        self.lbl_dict.config(text="本地字典：0 条（无字典模式）")
        self._refresh_step_states()
        messagebox.showinfo("已清除", "字典已清空：无字典全翻译模式")

    def _save_dict(self):
        """静默持久化到上次使用的字典文件（默认 dict.txt），不打断导入流程。"""
        p = self.cfg.get("dict_path") or str(PROJ / "dict.txt")
        try:
            Path(p).write_text(
                "\n".join(f"{k};{v}" for k, v in self.dict_map.items()),
                encoding="utf-8")
            self.cfg["dict_path"] = p
            _save_cfg(self.cfg)
        except OSError:
            pass   # 字典持久化失败不阻断主流程（内存字典仍可用）

    def _save_dict_as(self):
        """显式「另存字典…」：选择路径导出 EN;ZH 字典。"""
        if not self.dict_map:
            messagebox.showinfo("提示", "字典为空，无可保存内容")
            return
        p = filedialog.asksaveasfilename(
            defaultextension=".txt", initialfile="dict.txt",
            filetypes=[("字典 TXT", "*.txt")])
        if not p:
            return
        Path(p).write_text(
            "\n".join(f"{k};{v}" for k, v in self.dict_map.items()),
            encoding="utf-8")
        self.cfg["dict_path"] = p
        _save_cfg(self.cfg)
        messagebox.showinfo("完成", f"已保存 {len(self.dict_map)} 条 → {p}")

    def _dict_done(self, n, n_issue=0):
        self.lbl_dict.config(text=f"本地字典：{len(self.dict_map)} 条")
        self._refresh_step_states()
        self._save_dict()   # 主线程执行：写盘不允许出现在后台线程
        extra = f"，忽略 {n_issue} 条异常行" if n_issue else ""
        messagebox.showinfo("完成", f"合并 {n} 条{extra}，字典共 {len(self.dict_map)} 条"
                                    f"（已自动保存到 {self.cfg.get('dict_path', 'dict.txt')}）")

    # ------------------------------------------------------------------
    def pick_ct2(self):
        p = filedialog.askopenfilename(
            title="选择对照文件（CT 或 translations.json）",
            filetypes=[("CT / JSON", "*.ct *.json"), ("所有文件", "*.*")])
        if p:
            self.var_ct2.set(p)
            self.cfg["last_ct2"] = p
            _save_cfg(self.cfg)

    def _cmp_paths(self):
        a = self.var_ct.get().strip()
        b = self.var_ct2.get().strip()
        if not (a and Path(a).exists() and b and Path(b).exists()):
            messagebox.showwarning("提示", "请先在顶部选择 CT，并在「对照」页签选择对照文件")
            return None
        return a, b

    def do_diff(self):
        paths = self._cmp_paths()
        if not paths:
            return
        a, b = paths
        if not b.lower().endswith(".ct"):
            messagebox.showwarning("提示", "DIFF 需要两个 CT 文件（对照译文请用 COVER）")
            return

        def work():
            from eldenct.diffreport import diff_ct
            rows, stats = diff_ct(a, b)
            return rows, stats

        def done(res):
            rows, stats = res
            DiffWindow(self.root,
                       title=f"DIFF：{Path(a).name}  vs  {Path(b).name}",
                       rows=rows, stats=stats, kind="diff",
                       export_initial=f"{Path(a).stem}-diff-{Path(b).stem}.txt")

        self._run_bg(work, done)

    def do_cover(self):
        paths = self._cmp_paths()
        if not paths:
            return
        a, b = paths
        if not b.lower().endswith(".json"):
            messagebox.showwarning("提示", "COVER 需要选择 translations.json（对照两个 CT 请用 DIFF）")
            return

        def work():
            from eldenct.diffreport import diff_same_ct
            translations = json.loads(Path(b).read_text(encoding="utf-8"))
            rows, stats, unknown = diff_same_ct(a, translations)
            return rows, stats, unknown

        def done(res):
            rows, stats, unknown = res
            DiffWindow(self.root,
                       title=f"COVER 覆盖对照：{Path(a).name}",
                       rows=rows, stats=stats, kind="cover",
                       unknown=unknown,
                       export_initial=f"{Path(a).stem}-cover.txt")

        self._run_bg(work, done)

    def _ask_multiline(self, title: str, label: str) -> str | None:
        """多行文本输入对话框（simpledialog.askstring 只有单行，满足不了多行需求）。
        Ctrl+Enter 快捷确认；取消/关闭返回 None。"""
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, text=label).pack(anchor="w", padx=10, pady=(10, 2))
        txt = tk.Text(win, width=58, height=7, font=("Microsoft YaHei UI", 10),
                      wrap="char", undo=True)
        txt.pack(fill="both", expand=True, padx=10)
        result = {"text": None}

        def ok(event=None):
            result["text"] = txt.get("1.0", "end").strip()
            win.destroy()

        def cancel(event=None):
            win.destroy()

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Label(bar, text="Ctrl+Enter 确认", foreground="#888").pack(side="left")
        ttk.Button(bar, text="确定", command=ok).pack(side="right", padx=2)
        ttk.Button(bar, text="取消", command=cancel).pack(side="right")
        txt.focus_set()
        win.bind("<Control-Return>", ok)
        win.wait_window()
        return result["text"]

    def do_add_note(self):
        ct = self.ct_path   # 主线程快照（后台线程不读共享可变状态）
        if not ct:
            messagebox.showwarning("提示", "请先在顶部选择 CT 文件并提取")
            return
        text = self._ask_multiline(
            "追加注意事项", "输入注意事项文本（支持多行，直接写中文）：")
        if not text:
            return
        out = str(Path(ct).with_name(Path(ct).stem + "-note.CT"))

        def work():
            from eldenct.entryops import add_note_entry
            return add_note_entry(ct, text.strip(), out)

        def done(res):
            self.var_status.set(f"注意事项已追加：ID {res['id']} → {Path(res['out']).name}")
            messagebox.showinfo("完成", self.var_status.get())

        self._run_bg(work, done)

    # ------------------------------------------------------------------
    def do_apply(self):
        if not self.ct_path:
            messagebox.showwarning("提示", "请先提取 CT")
            return
        types = self._selected_types()
        if not types:
            messagebox.showwarning("提示", "类型开关全部未勾选：替换将零生效。请至少勾选一类。")
            return
        ct = self.ct_path   # 主线程快照（后台线程不读共享可变状态）
        out = filedialog.asksaveasfilename(
            defaultextension=".CT", initialfile=Path(ct).stem + "-zh.CT")
        if not out:
            return
        dict_snapshot = dict(self.dict_map)   # 主线程快照，杜绝后台读竞争

        def work():
            from eldenct.replace import apply_translations
            return apply_translations(ct, dict_snapshot, out,
                                      allowed_types=types)

        def done(r):
            self.var_status.set(
                f"替换完成：desc {r['desc_applied']} / dd {r['dd_applied']} / "
                f"lua {r['lua_applied']} / form {r['form_applied']}；"
                f"missed {len(r['missed'])}；unknown {len(r['unknown_sources'])}")
            messagebox.showinfo("完成", f"已生成：{out}\n\n{self.var_status.get()}")

        self._run_bg(work, done)


def main():
    root = tk.Tk()
    _enable_high_dpi(root)
    App(root)
    root.mainloop()


# ==========================================================================
# 可视化对照窗口：双栏同步滚动 · 字符级精细 diff · 多色标注 · 过滤/搜索/导出
# ==========================================================================
class DiffWindow:
    """rows 结构：(锚点ID, 字段, 文本A, 文本B, status)。

    DIFF:  changed | same | only_a | only_b
    COVER: translated | unchanged
    渲染：每条记录固定 3 行（头行 / 文本行 / 空行），两栏行数一致 →
    行级对齐 + yview/xview 双向同步；changed 条目做字符级 opcodes 标注。
    """

    LOAD_STEP = 1500        # 分批渲染（tk Text 一次性插 4 万条会卡死）

    # 多色标注（浅色主题）
    STATUS_COLOR = {"changed": "#B25000", "only_a": "#1F4E9C", "only_b": "#B26B00",
                    "same": "#8A8A8A", "translated": "#1E7B34", "unchanged": "#8A8A8A"}
    FILTERS = {
        "diff":  [("diff", "仅差异"), ("all", "全部"), ("changed", "仅 changed"),
                  ("only_a", "仅 only_a（A 独有）"), ("only_b", "仅 only_b（B 独有）"),
                  ("same", "仅 same")],
        "cover": ["translated", "unchanged"],
    }

    def __init__(self, master, title, rows, stats, kind, unknown=None,
                 export_initial="diff-report.txt"):
        self.rows = rows
        self.stats = stats
        self.kind = kind
        self.unknown = unknown or []
        self.export_initial = export_initial
        self.filtered: list = []
        self.shown = 0
        self.search_pos = 0
        self._syncing = False

        win = tk.Toplevel(master)
        win.title(title + " — eldenct 对照")
        win.geometry("1280x760")
        win.transient(master)
        self.win = win

        # ---- 工具条 ----
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=6, pady=4)
        if kind == "diff":
            options = self.FILTERS["diff"]
        else:
            options = [("translated", "仅已翻译"), ("unchanged", "仅未翻译"), ("all", "全部")]
        self._filter_map = {label: value for value, label in options}
        self.var_filter = tk.StringVar(value=options[0][1])
        cb = ttk.Combobox(bar, textvariable=self.var_filter, width=18,
                          state="readonly", values=[label for _, label in options])
        cb.pack(side="left", padx=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self._refilter())
        self.lbl_stat = ttk.Label(bar, text=self._stats_text(), foreground="#555")
        self.lbl_stat.pack(side="left", padx=10)

        ttk.Label(bar, text="查找:").pack(side="left", padx=(10, 0))
        self.var_find = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.var_find, width=24)
        ent.pack(side="left", padx=2)
        ent.bind("<Return>", self.do_find)
        ttk.Button(bar, text="下一个", command=self.do_find).pack(side="left", padx=2)
        ttk.Button(bar, text="加载更多", command=self._load_more).pack(side="left", padx=8)
        ttk.Button(bar, text="导出报告…", command=self.do_export).pack(side="right", padx=2)

        self.lbl_count = ttk.Label(win, text="", anchor="w")
        self.lbl_count.pack(fill="x", padx=8)

        # ---- 双栏文本区 ----
        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, padx=6, pady=(0, 4))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        common = {"wrap": "none", "state": "disabled", "font": ("Consolas", 9),
                  "undo": False, "cursor": "arrow"}
        self.text_a = tk.Text(body, **common)
        self.text_b = tk.Text(body, **common)
        self.vsb = ttk.Scrollbar(body, orient="vertical")
        self.hsb = ttk.Scrollbar(body, orient="horizontal")

        self.text_a.grid(row=0, column=0, sticky="nsew")
        self.text_b.grid(row=0, column=1, sticky="nsew")
        self.vsb.grid(row=0, column=2, sticky="ns")
        self.hsb.grid(row=1, column=0, columnspan=2, sticky="ew")

        # 分栏标题（固定行，不参与滚动）
        for t, head in ((self.text_a, "A 列（源）"), (self.text_b, "B 列")):
            t.tag_configure("colhead", foreground="#FFFFFF",
                            background="#4A6DA7", justify="center")
            t.insert("1.0", head, "colhead")

        # 多色标注 tag：头行按状态着色，差异字符红/绿底
        for t in (self.text_a, self.text_b):
            for st, color in self.STATUS_COLOR.items():
                t.tag_configure(f"st-{st}", foreground=color)
            t.tag_configure("seg-del", background="#FFC4C4")
            t.tag_configure("seg-ins", background="#C4E8C4")
            t.tag_configure("hit", background="#FFF3B0")   # 搜索命中当前条目

        # 同步滚动：垂直 + 水平 双向同步
        self.vsb.config(command=self._yview_both)
        self.hsb.config(command=self._xview_both)
        self.text_a.config(yscrollcommand=self._on_yscroll_a,
                           xscrollcommand=self._on_xscroll_a)
        self.text_b.config(yscrollcommand=self._on_yscroll_b,
                           xscrollcommand=self._on_xscroll_b)
        for t in (self.text_a, self.text_b):
            t.bind("<MouseWheel>", self._on_wheel)
            t.bind("<Button-4>", self._on_wheel)
            t.bind("<Button-5>", self._on_wheel)

        self._refilter()
        win.protocol("WM_DELETE_WINDOW", win.destroy)

    # ------------------------------------------------------------------
    def _stats_text(self) -> str:
        s = self.stats
        if self.kind == "diff":
            return (f"changed {s.get('changed', 0)} / same {s.get('same', 0)}"
                    f" / only_a {s.get('only_a', 0)} / only_b {s.get('only_b', 0)}")
        return (f"translated {s.get('translated', 0)} / unchanged {s.get('unchanged', 0)}"
                f" / unknown_sources {s.get('unknown_sources', 0)}")

    def _display(self, s: str) -> str:
        """单行显示安全：换行转为可见标记，保持双栏行数对齐。"""
        return s.replace("\r", " ").replace("\n", "\\n")

    def _refilter(self):
        mode = self._filter_map.get(self.var_filter.get(), "diff")
        if self.kind == "diff":
            if mode == "all":
                self.filtered = list(self.rows)
            elif mode == "diff":
                self.filtered = [r for r in self.rows if r[4] != "same"]
            else:
                self.filtered = [r for r in self.rows if r[4] == mode]
        else:
            if mode == "all":
                self.filtered = list(self.rows)
            else:
                self.filtered = [r for r in self.rows if r[4] == mode]
        self.search_pos = 0
        for t in (self.text_a, self.text_b):
            t.config(state="normal")
            t.delete("2.0", "end")   # 保留第 1 行栏标题
            t.config(state="disabled")
        self.shown = 0
        self._load_more()

    def _load_more(self):
        chunk = self.filtered[self.shown: self.shown + self.LOAD_STEP]
        if not chunk:
            self._update_count()
            return
        self.text_a.config(state="normal")
        self.text_b.config(state="normal")
        for row in chunk:
            self._render(row)
        self.text_a.config(state="disabled")
        self.text_b.config(state="disabled")
        self.shown += len(chunk)
        self._update_count()

    def _render(self, row):
        eid, field, ta, tb, st = row
        color = self.STATUS_COLOR.get(st, "#000000")
        hdr = f"[{st}] {eid} :: {field}\n"
        da, db = self._display(ta), self._display(tb)

        # A 栏
        self.text_a.insert("end", hdr, ("st-" + st,))
        self.text_a.insert("end", "  " + da + "\n")
        if st == "changed" and len(da) + len(db) <= 2000:
            self._mark_char_diff(self.text_a, da, db, "seg-del")
        self.text_a.insert("end", "\n")
        # B 栏
        self.text_b.insert("end", hdr, ("st-" + st,))
        self.text_b.insert("end", "  " + db + "\n")
        if st == "changed" and len(da) + len(db) <= 2000:
            self._mark_char_diff(self.text_b, db, da, "seg-ins")
        self.text_b.insert("end", "\n")

    def _mark_char_diff(self, text, own, other, tag):
        """字符级精细 diff：A 栏标 delete/replace 段，B 栏标 insert/replace 段。
        own/other 均为显示态字符串（\n 已转为 \\n），口径一致。"""
        try:
            sm = difflib.SequenceMatcher(None, own, other, autojunk=False)
        except Exception:
            return
        # 当前条目文本行 = 倒数第 2 行（最后插入的是空行；end-1c 落在空行上）
        line = int(text.index("end-1c").split(".")[0]) - 1
        prefix = 2   # "  " 前缀宽
        try:
            for op, i1, i2, _j1, _j2 in sm.get_opcodes():
                if op == "equal" or i2 <= i1:
                    continue
                text.tag_add(tag, f"{line}.{i1 + prefix}", f"{line}.{i2 + prefix}")
        except tk.TclError:
            pass

    def _update_count(self):
        self.lbl_count.config(
            text=f"已显示 {self.shown} / 共 {len(self.filtered)} 条"
                 + ("（继续点「加载更多」）" if self.shown < len(self.filtered) else ""))

    # ------------------------------------------------------------------
    # 同步滚动
    def _yview_both(self, *args):
        self.text_a.yview(*args)
        self.text_b.yview(*args)

    def _xview_both(self, *args):
        self.text_a.xview(*args)
        self.text_b.xview(*args)

    def _sync_pair(self, src, dst, first):
        if self._syncing:
            return
        self._syncing = True
        try:
            if dst.yview()[0] != first:
                dst.yview_moveto(first)
        finally:
            self._syncing = False

    def _on_yscroll_a(self, first, last):
        self.vsb.set(first, last)
        self._sync_pair(self.text_a, self.text_b, float(first))

    def _on_yscroll_b(self, first, last):
        self.vsb.set(first, last)
        self._sync_pair(self.text_b, self.text_a, float(first))

    def _on_xscroll_a(self, first, last):
        self.hsb.set(first, last)
        if not self._syncing and self.text_b.xview()[0] != float(first):
            self._syncing = True
            self.text_b.xview_moveto(float(first))
            self._syncing = False

    def _on_xscroll_b(self, first, last):
        self.hsb.set(first, last)
        if not self._syncing and self.text_a.xview()[0] != float(first):
            self._syncing = True
            self.text_a.xview_moveto(float(first))
            self._syncing = False

    def _on_wheel(self, event):
        delta = getattr(event, "delta", 0)
        if delta:                       # Windows
            step = -1 if delta > 0 else 1
        else:
            step = -1 if event.num == 4 else 1
        self.text_a.yview_scroll(step * 3, "units")
        self.text_b.yview_scroll(step * 3, "units")
        return "break"

    # ------------------------------------------------------------------
    def do_find(self, *args):
        q = self.var_find.get().strip().lower()
        if not q or not self.filtered:
            return
        n = len(self.filtered)
        for k in range(self.search_pos, self.search_pos + n):
            i = k % n
            eid, field, ta, tb, st = self.filtered[i]
            hay = f"{eid} {field} {ta} {tb}".lower()
            if q in hay:
                self._goto(i)
                self.search_pos = i + 1
                return
        messagebox.showinfo("查找", f"全部 {n} 条中未找到：{q}", parent=self.win)

    def _goto(self, idx: int):
        while self.shown <= idx:
            self._load_more()
        line = idx * 3 + 2   # 每条 3 行 + 第 1 行栏标题
        self.text_a.tag_remove("hit", "1.0", "end")
        self.text_b.tag_remove("hit", "1.0", "end")
        try:
            self.text_a.tag_add("hit", f"{line}.0", f"{line + 2}.end")
            self.text_b.tag_add("hit", f"{line}.0", f"{line + 2}.end")
        except tk.TclError:
            pass
        self.text_a.see(f"{line}.0")
        self.text_b.yview_moveto(self.text_a.yview()[0])

    # ------------------------------------------------------------------
    def do_export(self):
        out = filedialog.asksaveasfilename(
            parent=self.win, defaultextension=".txt",
            initialfile=self.export_initial, filetypes=[("文本报告", "*.txt")])
        if not out:
            return

        def fmt():
            lines = [f"# {self.win.title()}", f"# {self._stats_text()}", ""]
            for eid, field, ta, tb, st in self.filtered:
                lines.append(f"[{st}] {eid} :: {field}")
                if self.kind == "cover":
                    lines.append(f"    源: {ta}")
                    lines.append(f"    译: {tb}")
                else:
                    lines.append(f"    A: {ta}")
                    lines.append(f"    B: {tb}")
            if self.unknown:
                lines.append("")
                lines.append(f"# unknown_sources（译文集有、CT 无，共 {len(self.unknown)} 条）")
                lines.extend(f"  {u}" for u in self.unknown[:200])
            return "\n".join(lines)

        Path(out).write_text(fmt(), encoding="utf-8")
        self.lbl_count.config(text=f"已导出：{out}")
        os.startfile(out)


if __name__ == "__main__":
    main()
