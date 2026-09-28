# -*- coding: utf-8 -*-
"""法环 CT 翻译工具 — 可视化界面（Tkinter，标准库实现，无外部依赖）。

布局（顶部路径栏 + 主流程步骤条 / 左设置栏 + 中预览区 + 右功能页签 / 底部状态栏）：
  路径栏：选择 CT（路径 / 窗口几何 / 类型开关状态均记忆于 gui_config.json）
  主流程：① 提取词条 → ② 加载词典 ▾ → ③ 导出 XLSX → ④ AI 翻译 →
          ⑤ 人工定夺 ▾（拦截词条审核/导出/导回）→ ⑥ 批量替换
          步骤按钮按前置完成情况自动启用/置灰，始终知道下一步点哪
  左设置栏：类型开关（2 列 × 4 行，状态记忆）、最小提取长度、TXT 预览导出、
            字典状态与另存
  预览区：占满剩余高度，分批渲染（万级行不冻结），含原文/译文对照列
  功能页签：对照（DIFF/COVER，可清除所选即不 diff）| 工具 | 设置（AI 参数）
  状态栏：进度条 + 分阶段状态文字

批量替换生成新 CT（源文件只读，绝不覆盖；.bak 备份在首次提取时生成）。
所有耗时操作在后台线程执行，主界面不卡顿；进度条分阶段推进。
Windows 高分屏声明 DPI 感知，字体随系统缩放保持清晰。
"""
from __future__ import annotations

import difflib
import json
import os
import queue
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
        self.ct_path: str | None = None
        self.dict_map: dict[str, str] = {}   # 合并后的本地字典
        self.last_xlsx: str | None = None
        self._review: "FlaggedReviewWindow | None" = None   # 打开的定夺窗口（退出前查未存译文）
        self._pending_rows: list[dict] = []  # 预览渲染数据（与 tree 行序一致）
        self._busy = False                   # 后台任务互斥：防连点产生并发竞争
        self._ai_cancel = threading.Event()  # AI 翻译协作式取消（批次间检查）
        self._ai_stop_mode = False           # ④ 按钮当前处于「停止翻译」态
        # 跨线程 UI 回调唯一通道：工作线程绝不直接碰 Tk（root.after 在部分
        # 环境抛 "main thread is not in main loop" 且静默丢回调），改由
        # 主线程 30ms 轮询消费——AI 进度条曾因此全程不动（"没反应"的根因）
        self._ui_queue: queue.Queue = queue.Queue()

        self._build_ui()
        self._load_dict_file()   # 启动自动恢复上次字典（此前只存不载，重启即丢）
        if self.dict_map:
            self.lbl_dict.config(text=f"本地字典：{len(self.dict_map)} 条")
            self._refresh_step_states()
        self._poll_ui_queue()   # 启动跨线程 UI 回调轮询
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
        self.btn_dict = ttk.Menubutton(steps, text="② 加载词典 ▾")
        menu_dict = tk.Menu(self.btn_dict, tearoff=0)
        menu_dict.add_command(label="导入 TSV 字典（EN→ZH 两列）", command=self.do_import_tsv)
        menu_dict.add_command(label="导入 XLSX（8 列正典，含哈希校验）", command=self.do_import_xlsx)
        menu_dict.add_command(label="导入 TXT（原文/译文交替行）", command=self.do_import_txt)
        menu_dict.add_separator()
        menu_dict.add_command(label="清除字典（无字典全翻译）", command=self.do_clear_dict)
        self.btn_dict.config(menu=menu_dict)
        self.btn_export = ttk.Button(steps, text="③ 导出 XLSX",
                                     command=self.do_export_xlsx, state="disabled")
        self.btn_ai = ttk.Button(steps, text="④ AI 翻译",
                                 command=self._on_ai_button, state="disabled")
        self.btn_review = ttk.Menubutton(steps, text="⑤ 人工定夺 ▾", state="disabled")
        menu_rev = tk.Menu(self.btn_review, tearoff=0)
        menu_rev.add_command(label="打开定夺窗口（逐条审核 / 采纳 AI 建议）",
                             command=self.do_open_review)
        menu_rev.add_command(label="导出拦截词条 XLSX（供人工翻译）",
                             command=self.do_export_flagged_xlsx)
        menu_rev.add_command(label="导入人工翻译 XLSX（导回）", command=self.do_import_xlsx)
        self.btn_review.config(menu=menu_rev)
        self.btn_apply = ttk.Button(steps, text="⑥ 批量替换 → 新文件",
                                    command=self.do_apply, state="disabled")
        for i, b in enumerate((self.btn_extract, self.btn_dict, self.btn_export,
                               self.btn_ai, self.btn_review, self.btn_apply)):
            b.pack(side="left", fill="x", expand=True, padx=(6 if i == 0 else 3, 3))
        ttk.Label(steps, text="流程：提取 → 加载词典 → 导出（词典自动预填）→ AI 翻译 → "
                              "人工定夺拦截词条 → 替换输出；词典在 ③ 之前加载即自动预填",
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
        ttk.Label(lf_dict, text="导入走主流程 ②；导入/修改后自动保存到\n上次使用的字典文件（默认 dict.txt）",
                  foreground="#888", justify="left").pack(anchor="w", padx=6, pady=(0, 2))

        # 中：预览（占满剩余高度，原文/译文对照）
        prev = ttk.LabelFrame(main, text="预览（原文 / 译文 / 类型 / 出现次数 / 长度 / 上下文）")
        main.add(prev, weight=1)
        cols = ("src", "dst", "type", "line", "len", "ctx")
        self.tree = ttk.Treeview(prev, columns=cols, show="headings")
        for c, (w, t) in zip(cols, [(330, "原文"), (230, "译文"), (105, "类型"),
                                    (60, "出现次数"), (45, "长度"), (200, "上下文")]):
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
        ttk.Button(row_ct2, text="清除", command=self.do_clear_ct2).pack(side="left", padx=2)
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
        row_ai3 = ttk.Frame(lf_ai)
        row_ai3.pack(fill="x", padx=6, pady=2)
        ttk.Label(row_ai3, text="并发批次:").pack(side="left")
        self.var_maxworkers = tk.IntVar(value=self.cfg.get("ai_max_workers", 5))
        ttk.Spinbox(row_ai3, from_=1, to=16, width=4,
                    textvariable=self.var_maxworkers).pack(side="left", padx=4)
        ttk.Label(row_ai3, text="（同时请求的批次数，提速主杠杆；过大易被接口限流）",
                  foreground="#888").pack(side="left")
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
        rw = self._review
        if (rw is not None and rw.dst and rw.win.winfo_exists()
                and not messagebox.askyesno(
                    "退出", "定夺窗口存在尚未「存入字典」的人工译文，退出将丢弃。\n确定退出？")):
            return
        if self._busy and not messagebox.askyesno(
                "退出", "后台任务仍在执行，直接退出将中断本次运行。\n"
                        "（AI 翻译有断点保护，已完成的批次不会丢失，下次可续跑）\n\n"
                        "确定退出？"):
            return
        self.cfg["geometry"] = self.root.geometry()
        _save_cfg(self.cfg)
        self.root.destroy()

    def _save_types(self):
        self.cfg["types"] = {k: bool(v.get()) for k, v in self.type_vars.items()}
        _save_cfg(self.cfg)

    def _set_steps_state(self, state: str):
        for b in (self.btn_extract, self.btn_dict, self.btn_export,
                  self.btn_ai, self.btn_review, self.btn_apply):
            b.config(state=state)

    def _refresh_step_states(self):
        """主流程步骤按前置完成情况启停：③需已提取，④另需已导出 XLSX，
        ⑤需存在拦截词条文件（④ 完成后生成），⑥需字典非空。
        忙态（后台任务执行中）下六步全部锁定，杜绝并发竞争。"""
        if self._busy:
            self._set_steps_state("disabled")
            return
        self.btn_extract.config(state="normal")
        self.btn_dict.config(state="normal")
        has_terms = bool(self.terms)
        self.btn_export.config(state="normal" if has_terms else "disabled")
        self.btn_ai.config(
            state="normal" if (has_terms and self.last_xlsx) else "disabled")
        fp = self._flagged_path()
        self.btn_review.config(
            state="normal" if (self.last_xlsx and fp and fp.exists()) else "disabled")
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

    def _run_bg(self, fn, done, guard=True):
        """后台线程执行，完成后经队列通道回主线程回调。

        guard=True（默认）：主流程写操作（提取/导出/AI/替换/字典导入）——
        全程锁定六步并占用状态栏进度条。重复点击会派生并发任务，竞争同一份
        checkpoint / xlsx / 字典快照（如两个 AI 翻译同跑互相覆盖进度）。
        guard=False：旁路只读/独立输出操作（DIFF/COVER/导出 TXT/加注/拦截
        导出/对照表导出）——不锁步骤、不动全局进度条与状态栏（避免覆盖
        AI 实时进度），与进行中的主流程任务并行安全。返回 True=已启动。
        """
        if guard and self._busy:
            messagebox.showinfo("提示", "有任务正在后台执行，请等它完成或先停止。")
            return False
        if guard:
            self._busy = True
            self._refresh_step_states()   # 忙态 → 六步全锁定
            self.progress.start(12)
            self.var_status.set("处理中…")

        def worker():
            try:
                result = fn()
                err = None
            except Exception as e:      # noqa: BLE001
                result, err = None, e
            self._ui_queue.put(lambda: self._done(done, result, err, guard))

        threading.Thread(target=worker, daemon=True).start()
        return True

    def _poll_ui_queue(self):
        """主线程轮询消费工作线程的 UI 回调（30ms，空转开销可忽略）。"""
        while True:
            try:
                fn = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn()
            except tk.TclError:
                pass   # 回调目标窗口已销毁（如定夺窗口先关、导入后到）——正常竞态
            except Exception:   # noqa: BLE001
                import traceback
                traceback.print_exc()   # 单个回调失败绝不终止轮询（否则全部后续回调丢失）
        self.root.after(30, self._poll_ui_queue)

    def _done(self, done, result, err, guard=True):
        """统一出口。guard=True 时复位全局互斥态；先跑 done（可能派生新任务
        或更新 self.terms 等流程状态），再按最新状态刷新步骤启停——
        顺序反了会导致「提取完成后 ③ 仍置灰」这类状态滞留。"""
        if not guard:
            if err:
                messagebox.showerror("错误", str(err))
                return
            done(result)
            return
        self.progress.stop()
        self._ai_running = False   # 复位 AI 翻译心跳（无论成败）
        self._ai_stop_mode = False
        self._ai_cancel.clear()
        self.btn_ai.config(text="④ AI 翻译")
        self._busy = False
        if err:
            messagebox.showerror("错误", str(err))
            self.var_status.set(f"失败：{err}")
            self._refresh_step_states()
            return
        try:
            done(result)
        finally:
            if not self._busy:   # done 未派生新任务 → 按最新状态恢复启停
                self._refresh_step_states()

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
                r["src"][:120], self.dict_map.get(r["src"], ""), r["type"],
                f"×{r['count']}", len(r["src"]), r["ctx"][:80]))
        self._row_pos += PREVIEW_BATCH
        if self._row_pos < len(rows):
            self.var_status.set(f"预览加载中… {self._row_pos}/{len(rows)}")
            self.root.after(1, self._render_batch)
        else:
            self.var_status.set(self._final_status)

    def _fill_dst_column(self):
        """字典变化后回填预览「译文」列（分批刷新，万级行不卡顿）。"""
        rows = self._pending_rows
        items = self.tree.get_children()
        if not rows or len(items) != len(rows):
            return   # 预览尚未渲染完或数据已换，渲染时会按当前字典填入
        dm = self.dict_map
        pos = 0

        def step():
            nonlocal pos
            for it in items[pos: pos + PREVIEW_BATCH]:
                self.tree.set(it, "dst", dm.get(rows[pos]["src"], ""))
                pos += 1
            if pos < len(items):
                self.root.after(1, step)
        step()

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
        try:
            max_workers = max(1, min(16, int(self.var_maxworkers.get())))
        except (tk.TclError, ValueError):
            max_workers = 5
        # 记住填写的连接信息（Key 明文存本机 gui_config.json，仅本机使用）
        self.cfg.update(ai_api_key=api_key, ai_model=model, ai_base_url=base_url,
                        ai_max_workers=max_workers)
        _save_cfg(self.cfg)
        xlsx_in = self.last_xlsx
        out_json = Path(xlsx_in).with_name(Path(xlsx_in).stem + "-ai.json")

        # 进度可见性：进度条按条数推进 + 秒级心跳（首批返回前界面上不能是死的）。
        # _ai_running/tick 必须在 _run_bg 成功后才启动——忙态被拒时若已启动，
        # 心跳会永远空转并不断覆盖状态栏文字
        def tick():
            if not self._ai_running:
                return
            self.var_status.set(
                f"AI 翻译中… 已提交模型，等待批次返回（已等待 {int(time.time() - self._ai_t0)} 秒；"
                f"单批 30 条，接口超时上限 5 分钟，首批返回后每批更新进度条）")
            self.root.after(1000, tick)

        def work():
            from eldenct.aitrans import translate   # 待译总数由 translate 内部读取，勿重复解析整表

            self._ui_queue.put(lambda: (
                self.progress.stop(),   # 关闭 _run_bg 的自动步进，改为真实条数
                self.progress.configure(value=0)))

            def progress(done_n, total_n):
                def upd():
                    pct = done_n * 100 // max(total_n, 1)
                    self.var_status.set(f"AI 翻译中… {done_n}/{total_n} 条（{pct}%）")
                    self.progress.configure(value=done_n, maximum=max(total_n, 1))
                self._ui_queue.put(upd)   # 工作线程 → 主线程唯一安全通道
            return translate(xlsx_in, out_json,
                             Path(xlsx_in).with_name(Path(xlsx_in).stem + "-ai.xlsx"),
                             api_key=api_key, base_url=base_url, model=model,
                             max_workers=max_workers,
                             progress_cb=progress,
                             should_cancel=self._ai_cancel.is_set)

        def done(r):
            if r.get("cancelled"):
                # 取消不丢进度：合格批次已入 checkpoint + xlsx 副本，重跑续传
                self.var_status.set(
                    f"AI 翻译已停止：本轮合格 {r['translated']} 条已保留，"
                    f"剩余 {r['pending'] - r['translated'] - r['done_before']} 条待续"
                    f"（重新点 ④ 从断点续跑）")
                messagebox.showinfo(
                    "已停止",
                    f"已停止 AI 翻译。\n\n本轮合格 {r['translated']} 条已填入副本，"
                    f"断点已保存。\n重新点「④ AI 翻译」将从断点续跑，不重复计费。")
                return
            self.var_status.set(
                f"AI 翻译完成：{r['translated']}/{r['pending']} 条，"
                f"校验拦截 {r['flagged']} 条（详见 {Path(r['out_json']).with_suffix('.flagged.json').name}）")
            self._refresh_step_states()   # flagged 文件已落盘，⑤ 人工定夺解锁
            gl = f"\n术语表：{r['glossary_out']}" if r.get("glossary_out") else ""
            if messagebox.askyesno(
                    "AI 翻译完成",
                    f"合格 {r['translated']} 条已填入：{r['xlsx_out']}\n"
                    f"上下文补全 {r.get('ctx_filled', 0)} 行\n"
                    f"拦截 {r['flagged']} 条（留白待人工）{gl}\n\n"
                    f"是否立即导入该副本作为译文来源？"):
                self._import_ai_result(r["xlsx_out"])
            elif r["flagged"]:
                if messagebox.askyesno("人工定夺",
                                       f"有 {r['flagged']} 条被拦截的词条需要人工定夺，"
                                       "是否立即打开定夺窗口？"):
                    self.do_open_review()

        self._ai_cancel.clear()   # 清掉上一次的取消信号（先清后跑）
        if self._run_bg(work, done):
            # 六步已锁定；单独复活 ④ 为「停止翻译」——长任务必须随时可停
            self._ai_running = True
            self._ai_t0 = time.time()
            self.root.after(0, tick)
            self._ai_stop_mode = True
            self.btn_ai.config(state="normal", text="④ 停止翻译")

    def _on_ai_button(self):
        """④ 双态按钮：平时发起新翻译；翻译中点击 = 协作式取消。"""
        if self._ai_stop_mode:
            self._ai_cancel.set()
            self.btn_ai.config(state="disabled", text="④ 正在停止…")
            self.var_status.set(
                "AI 翻译：已请求停止，等待在途批次返回…（已完成批次全部保留，可续跑）")
            return
        self.do_ai_translate()

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

        self._run_bg(work, lambda o: messagebox.showinfo("完成", f"已导出：{o}"),
                     guard=False)   # 独立输出文件，与主流程并行安全

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
        self._save_dict()   # 空字典立即落盘：否则重启自动加载把清空悄悄撤销
        self.lbl_dict.config(text="本地字典：0 条（无字典模式）")
        self._refresh_step_states()
        self._fill_dst_column()   # 预览「译文」列同步清空
        messagebox.showinfo("已清除", "字典已清空：无字典全翻译模式")

    def _load_dict_file(self):
        """启动时自动加载上次的字典文件（EN;ZH 每行一条，split 首个分号）。

        此前字典只保存不加载：重启后 dict_map 恒为空，用户需重新导入——
        现与 _save_dict 形成闭环，GUI 字典跨会话持久。
        """
        p = self.cfg.get("dict_path") or str(PROJ / "dict.txt")
        if not Path(p).exists():
            return
        try:
            d: dict[str, str] = {}
            for line in Path(p).read_text(encoding="utf-8").splitlines():
                if ";" in line:
                    k, v = line.split(";", 1)
                    if k and v.strip():
                        d[k] = v
            self.dict_map = d
        except OSError:
            pass   # 加载失败不阻断启动（可手动经 ② 重新导入）

    def _write_dict_file(self, p: str) -> int:
        """字典写盘（EN;ZH 每行一条）。返回因格式不兼容被跳过的条数。

        键是替换匹配的精确依据，含 ';' / 换行的键无法以单行 EN;ZH 无损表达
        （改写会造成重载后键错配），只能跳过（内存中仍有效，本次会话可用）；
        值仅显示/写入用，换行降级为空格无损语义。
        """
        def _flat(s: str) -> str:
            return s.replace("\r", " ").replace("\n", " ")

        lines, skipped = [], 0
        for k, v in self.dict_map.items():
            if ";" in k or "\n" in k or "\r" in k:
                skipped += 1
                continue
            lines.append(f"{k};{_flat(v)}")
        Path(p).write_text("\n".join(lines), encoding="utf-8")
        return skipped

    def _save_dict(self):
        """静默持久化到上次使用的字典文件（默认 dict.txt），不打断导入流程。"""
        p = self.cfg.get("dict_path") or str(PROJ / "dict.txt")
        try:
            self._write_dict_file(p)
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
        try:
            skipped = self._write_dict_file(p)
        except OSError as e:
            messagebox.showerror("错误", f"字典保存失败：{e}")
            return
        self.cfg["dict_path"] = p
        _save_cfg(self.cfg)
        extra = f"（{skipped} 条因含分号/换行无法以 EN;ZH 格式保存，已跳过）" if skipped else ""
        messagebox.showinfo("完成", f"已保存 {len(self.dict_map) - skipped} 条 → {p}{extra}")

    def _dict_done(self, n, n_issue=0, quiet=False):
        self.lbl_dict.config(text=f"本地字典：{len(self.dict_map)} 条")
        self._refresh_step_states()
        self._fill_dst_column()   # 预览「译文」列实时对照
        self._save_dict()   # 主线程执行：写盘不允许出现在后台线程
        if quiet:   # 静默模式：定夺窗口关闭等场景不打断操作流
            return
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

        self._run_bg(work, done, guard=False)   # 只读对照，可与 AI 翻译并行

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

        self._run_bg(work, done, guard=False)   # 只读对照，可与 AI 翻译并行

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
            messagebox.showinfo(
                "完成", f"注意事项已追加：ID {res['id']} → {Path(res['out']).name}")

        self._run_bg(work, done, guard=False)   # 独立输出 -note.CT，不碰共享状态

    # ------------------------------------------------------------------
    # ⑤ 人工定夺：拦截词条审核 / 导出 / 导回
    # ------------------------------------------------------------------
    def _flagged_path(self) -> Path | None:
        """当前 XLSX 对应的 AI 拦截词条文件（terms-ai.flagged.json）。"""
        if not self.last_xlsx:
            return None
        return Path(self.last_xlsx).with_name(
            Path(self.last_xlsx).stem + "-ai.flagged.json")

    def do_open_review(self):
        fp = self._flagged_path()
        if not fp or not fp.exists():
            messagebox.showwarning(
                "提示", "未找到拦截词条文件（*-ai.flagged.json）。\n请先执行 ④ AI 翻译。")
            return
        self._review = FlaggedReviewWindow(self.root, self, fp)

    def do_export_flagged_xlsx(self):
        """拦截词条导出为正典 8 列 XLSX：AI 建议与拦截原因写入备注列（参考不填入）。"""
        fp = self._flagged_path()
        if not fp or not fp.exists():
            messagebox.showwarning("提示", "未找到拦截词条文件，请先执行 ④ AI 翻译。")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=Path(self.last_xlsx).stem + "-flagged.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        type_map = {r["src"]: r["type"].split("+") for r in self.terms}   # 主线程快照
        ctx_map = {r["src"]: r["ctxs"] for r in self.terms if r["ctxs"]}

        def work():
            from eldenct.xlsxio import export_terms_xlsx
            flagged = json.loads(fp.read_text(encoding="utf-8"))
            seen: set[str] = set()
            terms, partial = [], {}
            for f in flagged:
                src = f.get("src", "")
                if not src or src in seen:
                    continue
                seen.add(src)
                terms.append({"source": src,
                              "type": type_map.get(src, ["Description"]),
                              "count": 1, "class": "normal",
                              "contexts": [], "entries": []})
                partial[src] = (f"AI建议：{f.get('zh') or '（无）'}"
                                f"｜拦截原因：{f.get('reason', '?')}")
            return export_terms_xlsx(terms, out, auto_translations={},
                                     partial_notes=partial, contexts=ctx_map)

        def done(info):
            messagebox.showinfo(
                "完成", f"拦截词条已导出：{out}\n主表 {info['rows']} 行（全部待人工翻译）\n"
                        f"AI 建议与拦截原因在备注列，仅供参考、不自动填入。\n"
                        f"翻译完成后用 ⑤「导入人工翻译 XLSX」导回。")

        self._run_bg(work, done, guard=False)   # 独立输出 xlsx，并行安全

    def _export_compare_xlsx(self, ct_out: str):
        """替换完成后导出原文/译文对照表（词条 × 当前字典译文，空 = 未译）。"""
        terms_snapshot = list(self.terms)
        dict_snapshot = dict(self.dict_map)
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=Path(ct_out).stem + "-对照.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not out:
            return

        def work():
            import xlsxwriter
            wb = xlsxwriter.Workbook(str(Path(out).resolve()))
            ws = wb.add_worksheet("原文译文对照")
            fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
            ws.write_row(0, 0, ["原文", "译文", "类型", "出现次数"], fmt)
            for i, r in enumerate(terms_snapshot, start=1):
                ws.write_string(i, 0, r["src"])
                ws.write_string(i, 1, dict_snapshot.get(r["src"], ""))
                ws.write_string(i, 2, r["type"])
                ws.write_number(i, 3, r["count"])
            for c, w in enumerate([60, 60, 16, 10]):
                ws.set_column(c, c, w)
            wb.close()
            return len(terms_snapshot)

        def done(n):
            self.var_status.set(f"对照表已导出：{out}")
            messagebox.showinfo("完成", f"对照表 {n} 行已导出：\n{out}")

        self._run_bg(work, done, guard=False)   # 独立输出 xlsx，并行安全

    def do_clear_ct2(self):
        """清除对照文件选择（不 diff）。"""
        self.var_ct2.set("")
        self.cfg["last_ct2"] = ""
        _save_cfg(self.cfg)

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
            if messagebox.askyesno(
                    "完成", f"已生成：{out}\n\n{self.var_status.get()}\n\n"
                            f"是否导出原文/译文对照表（XLSX）？"):
                self._export_compare_xlsx(out)

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

    def _sync_pair(self, dst, first):
        """把 dst 栏的滚动位置对齐到 first（_syncing 防两栏互相触发死循环）。"""
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
        self._sync_pair(self.text_b, float(first))

    def _on_yscroll_b(self, first, last):
        self.vsb.set(first, last)
        self._sync_pair(self.text_a, float(first))

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


# ==========================================================================
# 人工定夺窗口：AI 拦截词条逐条审核 / 采纳建议 / 导出 XLSX / 导回
# ==========================================================================
class FlaggedReviewWindow:
    """拦截词条（*-ai.flagged.json）人工定夺。

    行数据：{"src", "reason", "zh"（AI 建议，可能缺失）}；同 src 去重。
    人工译文暂存 self.dst{src: zh}，「存入字典并关闭」统一走
    App._dict_done 合并 + 持久化 dict.txt + 预览译文列回填。
    """

    FILL_STEP = 500   # 分批渲染

    def __init__(self, master, app: "App", flagged_path):
        self.app = app
        self.flagged_path = Path(flagged_path)
        raw = json.loads(self.flagged_path.read_text(encoding="utf-8"))
        seen: set[str] = set()
        self.items: list[dict] = []
        for f in raw:
            if f.get("src") and f["src"] not in seen:
                seen.add(f["src"])
                self.items.append(f)
        self.dst: dict[str, str] = {}
        self._pos = 0
        self._cur_src: str | None = None

        win = tk.Toplevel(master)
        win.title(f"人工定夺 — AI 拦截词条 {len(self.items)} 条（{self.flagged_path.name}）")
        win.geometry("1150x680")
        win.transient(master)
        self.win = win

        # ---- 工具条 ----
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="全部采纳 AI 建议", command=self._accept_all_ai).pack(
            side="left", padx=2)
        ttk.Button(bar, text="导出拦截 XLSX…", command=self.app.do_export_flagged_xlsx).pack(
            side="left", padx=2)
        ttk.Button(bar, text="从 XLSX 导回…", command=self._import_back).pack(
            side="left", padx=2)
        self.lbl_stat = ttk.Label(bar, text=self._stat_text(), foreground="#555")
        self.lbl_stat.pack(side="left", padx=10)
        ttk.Button(bar, text="存入字典并关闭", command=self._save_close).pack(
            side="right", padx=2)

        # ---- 列表 ----
        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, padx=6, pady=2)
        cols = ("src", "ai", "reason", "dst")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")
        for c, (w, t) in zip(cols, [(340, "原文"), (230, "AI 建议（校验不合格）"),
                                    (150, "拦截原因"), (230, "人工译文")]):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.tree.tag_configure("done", foreground="#1E7B34")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ---- 编辑区 ----
        ed = ttk.LabelFrame(win, text="编辑（选中行后填写；Enter 保存该条；留空 = 不译）")
        ed.pack(fill="x", padx=6, pady=(2, 6))
        self.lbl_src = ttk.Label(ed, text="（未选中）", wraplength=1050, justify="left")
        self.lbl_src.pack(anchor="w", padx=8, pady=(4, 2))
        row = ttk.Frame(ed)
        row.pack(fill="x", padx=8, pady=(0, 6))
        self.var_edit = tk.StringVar()
        self.ent = ttk.Entry(row, textvariable=self.var_edit)
        self.ent.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="采纳 AI 建议", command=self._accept_ai_cur).pack(
            side="left", padx=4)
        ttk.Button(row, text="保存该条", command=self._save_one).pack(side="left")
        self.ent.bind("<Return>", lambda e: self._save_one())

        win.protocol("WM_DELETE_WINDOW", self._on_close)
        self._fill_tree()

    # ------------------------------------------------------------------
    def _stat_text(self) -> str:
        return f"已定夺 {len(self.dst)} / 共 {len(self.items)} 条"

    def _row_values(self, idx: int) -> tuple:
        f = self.items[idx]
        zh = self.dst.get(f["src"], "")
        return (f["src"], f.get("zh", ""), f.get("reason", ""), zh)

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        self._pos = 0
        self._fill_step()

    def _fill_step(self):
        for idx in range(self._pos, min(self._pos + self.FILL_STEP, len(self.items))):
            f = self.items[idx]
            zh = self.dst.get(f["src"], "")
            self.tree.insert("", "end", iid=str(idx), values=(
                f["src"], f.get("zh", ""), f.get("reason", ""), zh),
                tags=("done",) if zh else ())
        self._pos = min(self._pos + self.FILL_STEP, len(self.items))
        if self._pos < len(self.items):
            self.win.after(1, self._fill_step)

    def _refresh_rows(self):
        for idx in range(len(self.items)):
            zh = self.dst.get(self.items[idx]["src"], "")
            self.tree.item(str(idx), values=self._row_values(idx),
                           tags=("done",) if zh else ())
        self.lbl_stat.config(text=self._stat_text())

    def _on_select(self, *_):
        sel = self.tree.selection()
        if not sel:
            return
        f = self.items[int(sel[0])]
        self._cur_src = f["src"]
        self.lbl_src.config(text=f["src"])
        self.var_edit.set(self.dst.get(f["src"], f.get("zh", "") or ""))
        self.ent.focus_set()

    def _save_one(self):
        if self._cur_src is None:
            return
        v = self.var_edit.get().strip()
        if v:
            self.dst[self._cur_src] = v
        else:
            self.dst.pop(self._cur_src, None)
        for idx, f in enumerate(self.items):
            if f["src"] == self._cur_src:
                zh = self.dst.get(f["src"], "")
                self.tree.item(str(idx), values=self._row_values(idx),
                               tags=("done",) if zh else ())
                break
        self.lbl_stat.config(text=self._stat_text())

    def _accept_ai_cur(self):
        sel = self.tree.selection()
        if not sel:
            return
        f = self.items[int(sel[0])]
        if not f.get("zh"):
            messagebox.showinfo("提示", "该条没有 AI 建议（API 未返回）。", parent=self.win)
            return
        self.var_edit.set(f["zh"])
        self._save_one()

    def _accept_all_ai(self):
        for idx, f in enumerate(self.items):
            if f.get("zh"):
                self.dst[f["src"]] = f["zh"]
                self.tree.item(str(idx), values=self._row_values(idx), tags=("done",))
        self.lbl_stat.config(text=self._stat_text())   # 状态栏实时可见，不打断

    def _import_back(self):
        p = filedialog.askopenfilename(parent=self.win,
                                       filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        known = {f["src"] for f in self.items}

        def work():
            from eldenct.xlsxio import import_translations_xlsx
            translations, issues, _e = import_translations_xlsx(p)
            return translations, len(issues)

        def done(res):
            translations, n_issue = res
            hit = {k: v for k, v in translations.items() if k in known}
            self.dst.update(hit)
            self._refresh_rows()
            extra = f"，忽略 {n_issue} 条异常行" if n_issue else ""
            messagebox.showinfo(
                "完成", f"导入 {len(translations)} 条，命中拦截词条 {len(hit)} 条{extra}",
                parent=self.win)

        self.app._run_bg(work, done, guard=False)   # 只改本窗口暂存区，并行安全

    def _save_close(self):
        if not self.dst:
            if messagebox.askyesno("关闭", "未填写任何译文，确定直接关闭？", parent=self.win):
                self.win.destroy()
            return
        self.app.dict_map.update(self.dst)
        self.app._dict_done(len(self.dst), quiet=True)   # 静默合并：不打断关闭动作
        self.win.destroy()

    def _on_close(self):
        if self.dst and not messagebox.askyesno(
                "关闭", "存在尚未「存入字典」的人工译文，关闭将丢弃。\n确定关闭？",
                parent=self.win):
            return
        self.win.destroy()


if __name__ == "__main__":
    main()
