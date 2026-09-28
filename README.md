# eldenct — ELDEN RING CT 精准翻译工具

GitHub: <https://github.com/SantaChains/ctterm>

面向《艾尔登法环》Cheat Table（.CT，CheatTable XML）的翻译流水线。

工作流固定为 **导出 → 人工 / AI 翻译 → 导回 → 精准替换 → 乱序 diff 对照**。

AI 翻译（DeepSeek）为内置可选环节，也可完全离线用自己的翻译工具完成。

## 快速开始

```
双击 启动GUI.bat            ← 可视化操作（推荐）
.venv\Scripts\python.exe -m eldenct <子命令> ...   ← 命令行
```

## 仓库结构

```
ctterm/
├─ eldenct/          # 管线核心（15 个模块，见「命令」）
│  ├─ ctmodel.py     #   CT XML 模型（Description / DropDownList / Lua / Form）
│  ├─ extract.py     #   词条提取（去重 / 分类 / 无字母剔除）
│  ├─ xlsxio.py      #   xlsx 正典 8 列导出导回（哈希校验）
│  ├─ replace.py     #   精准替换（行级 + 写盘前严格校验 + 模板展开）
│  ├─ aitrans.py     #   AI 翻译（批量 / 校验 / checkpoint 断点续跑）
│  ├─ matcher.py     #   Aho-Corasick 术语匹配（最左最长 + 整词边界）
│  ├─ glossary.py    #   词表合并（多级优先级 + 退化条目过滤）
│  ├─ diffreport.py  #   乱序 diff / 覆盖核验
│  ├─ incremental.py #   跨版本增量（TM 继承 + ID 对齐参考）
│  ├─ entryops.py    #   追加注意事项词条
│  └─ forms.py / luaextract.py / cli.py / __main__.py
├─ gui.py            # 可视化界面（Tkinter 标准库，步骤条 + 三栏 + 页签）
├─ _gui_launch.pyw   # 带保护的启动器（崩溃日志 + 心跳）
├─ 启动GUI.bat       # 双击入口
├─ data/
│  └─ build_glossary_qualified.py   # 合格词表构建脚本（数据本身不入库）
├─ gui_config.example.json          # 配置模板（复制为 gui_config.json 使用）
├─ pyproject.toml / LICENSE / README.md
```

词表与字典文件（`glossary_qualified.tsv` / `combine_flat.json` /
`idlist-3.0-dlc.tsv` 等）为外部素材的衍生数据，**不入库**——按上文
「配置」节放入词表目录即可，构建方式见 `data/build_glossary_qualified.py`
与「边界决策记录」第 4 节。

## 工作流

```
Hexinton-v8.0.4.CT
   │  eldenct pipeline --glossary        （extract + export 一条龙）
   ▼
terms.xlsx ───────────────► 翻译（二选一）
   │                          A. 人工：翻译软件另一列填译文
   │                          B. AI：eldenct ai-translate（DeepSeek）
   ▼
eldenct import terms_translated.xlsx
   ▼
translations.json
   │  eldenct apply Hexinton-v8.0.4.CT translations.json -o Hexinton-zh.CT
   ▼
Hexinton-zh.CT ──────────►  eldenct cover / diff 对照核验
```

## 安装（新环境）

```
git clone https://github.com/SantaChains/ctterm.git
cd ctterm
uv venv
uv pip install --python .venv\Scripts\python.exe lxml xlsxwriter openpyxl pyahocorasick
```

注意：本 venv 由 uv 创建（无 pip），装依赖一律走 `uv pip install --python ...`。
热点全在 C 层 / 专项写入器：XML 解析用 lxml（libxml2）、多模式匹配用
pyahocorasick（C 自动机）、xlsx 写用 xlsxwriter。8 MB、1.3 万条目的 CT
全流程单线程约 3 秒，无需多线程 / CUDA。

## GUI 启动

- **双击 `启动GUI.bat`**（推荐）：cmd 拉起 pythonw，稳定存活。
- `_gui_launch.pyw`：带保护的启动器——异常写 `_gui_err.log`，另有 5 秒
  心跳线程写 `_gui_heartbeat.txt`（停跳 = 进程不在）。
- 配置（模型 / API Key / Base URL / 路径记忆）持久化在 `gui_config.json`。

GUI 流程（推荐时序，主界面为横向步骤条，按前置完成情况自动启停）：

1. 顶部路径栏选 CT → **① 提取词条**
2. **④ 导入译文 ▾** 菜单**导入字典**：TSV（EN→ZH 两列）/ XLSX / TXT，
   **在 ② 导出 XLSX 之前导入**即自动预填（用户字典优先于内置词表）；
   「清除字典」= 无字典全翻译。导入后自动保存到上次字典文件（默认 `dict.txt`），
   「另存字典…」可显式导出
3. **② 导出 XLSX** →（可选）**③ AI 翻译** → 导回译文 → **⑤ 批量替换 → 新文件**
4. 「对照」页签：DIFF 两 CT / COVER 覆盖对照，弹出**可视化对照窗口**——
   双栏同步滚动（垂直+水平）、changed 条目字符级红绿精细标注、
   按状态过滤（仅差异/仅 changed/only_a/b/same…）、查找跳转、
   分批渲染（4 万+ 条不卡）、报告可导出 txt 自动打开

其他界面元素：类型开关（2 列勾选框，状态记忆）与最小提取长度在左栏；
「工具」页签放多行注意事项输入；「设置」页签放 AI 连接参数；
窗口尺寸/位置关闭时记忆，Windows 高分屏下字体随系统缩放保持清晰。

## 配置

- 运行时配置持久化在 `gui_config.json`（含 API Key 明文——**已被 .gitignore
  排除，绝不入库**）。新环境可从 `gui_config.example.json` 复制一份改名后填写。
- **词表目录**（术语兜底，可选）：默认 `<仓库根>/eldenring/`，可用环境变量
  `ELDENCT_WORD_ROOT` 覆盖。放入以下文件即自动加载（缺失则跳过，管线照常工作）：
  `glossary_qualified.tsv`（合格术语表，见下文）、`combine_flat.json` /
  `eldenring.json` / `tarnished.json`（社区词表，格式 `{英文: 中文}` 扁平 dict）、
  `combine_conflicts.tsv`（配合 `--conflicts`）。

## 命令



| 命令                                                                | 作用                 |
| ----------------------------------------------------------------- | ------------------ |
| `extract <CT> [-o terms.json]`                                    | 解析 CT，提取去重词条与统计    |
| `export terms.json [-o terms.xlsx] [--glossary] [--skip-param]`   | 导出待翻译 xlsx         |
| `import terms.xlsx [-o translations.json]`                        | 导回译文（校验词条哈希）       |
| `apply <CT> translations.json -o <新CT> [--glossary] [--fallback] [--types]` | 执行替换（源只读）    |
| `cover <CT> translations.json [-o cover.txt]`                     | 翻译覆盖对照（乱序无关）       |
| `diff <CT_A> <CT_B> [-o report.txt]`                              | 两个 CT 按 ID 乱序对照    |
| `pipeline <CT> [-d outdir] [--glossary]`                          | 一条龙 extract+export |
| `delta <新CT> -o delta.xlsx [--zh 旧译CT] [--en 旧EN] [--tm 旧译json]`  | 增量导出：TM 继承预填 + 上下文列 |
| `merge base.json delta.json -o full.json`                         | 合并译文集（delta 优先，冲突报告） |
| `ai-translate terms.xlsx [--api-key] [--model] [--base-url] [--dry-run]` | AI 翻译待译行（DeepSeek；校验后填回 xlsx 副本） |
| `python gui.py`                                                   | 可视化界面（类型开关/进度条/路径记忆/.bak） |

**ai-translate 说明**：

```bash
# 离线演练（不调 API，验证校验/回填链路）
.venv\Scripts\python.exe -m eldenct ai-translate data\terms.xlsx --dry-run
# 真实翻译（Key 也可从环境变量 DEEPSEEK_API_KEY 读取）
.venv\Scripts\python.exe -m eldenct ai-translate data\terms.xlsx --api-key sk-xxx --model deepseek-chat
```

### 追加注意事项词条

```
:: 在根级词条末尾追加一条 GroupHeader 纯文本注意事项（ID 自动唯一，
:: 写盘前严格校验，源文件只读，默认输出 <源名>-note.CT）
.venv\Scripts\python.exe -m eldenct add-note Hexinton-v8.0.4.CT "【注意事项】先备份原表"
```

导出结构说明：DropDown 词条的上下文列为「所属条目 > 父级」（AI 消歧用）；
源含日文的行备注「建议汉化」，已含中文的行备注「建议保留原样或仅校对」。

- 内置提示词：FromSoftware 术语库锁定引擎（术语优先级 艾尔登法环 > 黑暗之魂 > 血源诅咒 > 只狼），
  R1 术语锁定（Arcane=感应 硬锁定，奥术为错译）/ R2 格式锁 / R3 数字映射 / R4 函数风格合成词
  语素拆解（唯一允许 '可能含义' 注释）/ R5 肃穆叙事腔 / R6 JSON 输出契约 + GOLDEN SAMPLES
- 上下文补全：界面位置为空的行自动挂同族邻居（'Living Pot - Large' ← 'Small | Huge'），
  函数风格行（驼峰/下划线）标 style:function-like 提示语素拆解
- 温度 0.2（术语锁定翻译要确定性）；校验不合格（{0} 数量/数字保真/术语未译/空译）一律留白
  并写 flagged 报告，绝不带病替换
- checkpoint 断点续跑；产物 terms-ai.xlsx（备注标"AI 译文（待校对）"）→ 照常 导入 → apply；
  另产 terms-ai.glossary.tsv 原文→译文术语表（人工校对后可 --user-terms 复用，dry-run 不落盘）
- 「设置」页签 AI 翻译同功能；`data/user_terms.tsv` 为官方属性名用户术语表（--user-terms 可让预填/fallback 也用感应等官方译名）

类型开关（--types，GUI 为勾选框）：`Description,DropDown,Lua:LuaComment,
Lua:AAComment,Lua:Caption,Lua:Msg,Lua:Guide,Form:String`。导出与替换严格受控，
未勾选的类型绝不参与。

## 翻译范围（三层）

1. **Description / DropDownList**：菜单标签与 `ID:名称` 下拉项（ID 前缀结构保留）。
2. **Lua/AA 脚本层**（`<AssemblerScript>` 内，含 `{$lua}` 块）：
   - `Lua:LuaComment`：`-- text` 行注释与单行/跨行 `--[[ ]]`（字符串字面量感知，
     不命中字符串内部的 `--`）；
   - `Lua:AAComment`：汇编模式 `// text`；
   - `Lua:Caption`：`Caption = "..."` 控件标题（`Name = "..."` 是控件标识符，
     Lua 按 `win.ctrl["Name"]` 引用，**禁译**）；
   - `Lua:Msg`：`messageDialog / error / showMessage` 的字符串实参（Lua 转义感知）；
   - `Lua:Guide`：长字符串 `[[ ]]` 内的引导行 `"key(s)": content`——冒号后为空/
     空白则忽略；`"OK", "Cancel": …` 多键特例只导出冒号后内容。
   脚本内提取/替换在「未转义逻辑行」域进行并维护 logical→raw 下标映射，
   写回只拼接 raw 前缀 + Lua/XML 双层转义的新文本 + raw 后缀，非目标字节零改动。
3. **Ascii85 表单块**（`<Forms>` 下 `Encoding="Ascii85"`，Hexinton 10 个）：
   CE 私有 Base85 字符集（XML 安全 85 字符）→ RFC1924 b85 → raw deflate →
   4 字节长度前缀 + TPF0 表单流。仅解析 type6 字符串属性（type7 Ident 为
   枚举名禁译），其余属性字节级透传。10/10 表单 round-trip 字节级一致，
   58 条 UI 字符串（ItemGib 26 / WorldChrForm 12…）进入词条。

术语表优先级（--glossary 时生效）：

用户自定义（--user-terms，tsv/json）> combine\_flat.json > eldenring.json > tarnished.json。

异值表 combine\_conflicts.tsv 默认不进入（歧义），--conflicts 显式启用。

## 合格术语表 glossary\_qualified

构建脚本 `data/build_glossary_qualified.py`（可复用）。当前使用的合格表在**工具根目录** `glossary_qualified.tsv`（12060 键，TSV 无引号语义解析）。

* `glossary_qualified.tsv` / `.json`（工具根目录 / eldenring 目录各一份）— EN→ZH 合格术语表（12060 键）
* `glossary_qualified_missing.tsv` — CT 真实存在但暂无可靠译文的词条清单

防错规则（宁可缺译、绝不译错）：

1. **键 100% 来自当前 CT 真实词条**（16557 去重后，剔除纯数字/纯符号 468）；CT 不出现的键一律不进。
2. 译文来源优先级：Hexinton 官方多语言 ID List xlsx > 杂乱/id列表.xlsx > 人工校对版（ct英文校对版.txt）> combine_flat.json。
3. **人工校对是 entry 级 Description 译文，绝不用作 DropDown 名称译文**——实测抓出地名 "Weeping Peninsula - Tombsward Cave" 被错配成功能注释"恩惠下拉菜单"，已按字段来源拆分修正。
4. 译文须含 CJK 且无全角 ASCII（误译信号）；无可靠译文 → 进 missing，不硬译。
5. 引擎侧已实测：12060 键全量替换 desc 8557 + dd 17363，missed 0、unknown 0、strict XML OK、行数一致、残留 0；短键（h/HP/No…）嵌于长键时由整词边界 + 最长匹配正确保护，长句不被拆。

覆盖：16508 词条中 12060 已译（73%）；缺译 4448 = 技术参数 857（建议保持英文）+ 真可译 3591（多为此前 ID 表未收录的带地理后缀 BOSS 名与 About 教程句，可走 xlsx 流程人工/AI 补齐）。

## xlsx 列约定



| 列     | 含义                              |
| ----- | ------------------------------- |
| 词条 ID | 源文本 SHA1 前 12 位，导回时校验（防源列被改）    |
| 类型    | Description / DropDown          |
| 源文本   | 英文原文（唯一键）                       |
| 译文    | 空待填；术语表整串命中的已预填（绿色行，可改）         |
| 出现次数  | 同词条在 CT 中的总出现次数（一次翻译全生效）        |
| 分类    | normal 需翻译 /param（有术语表译文，已预填） |
| 备注    | 预填来源 / 含数字符号提示                 |

另有一张副表「参数保留」：slot25 / sfxld_170 这类无术语表译文的纯标识符
与代码成分（模板合并后 717 行）——翻译纯属浪费，不进主表、不参与导回，
CT 中原样保留。主表只含需要翻译/已预填的内容，交给翻译软件零浪费。

**两项提取瘦身**：
1. 无字母词条（纯数字/空格/符号，如 '[ 0'、'---'）根本不提取——
   未提取 = CT 原样保留，效果等同留空且省审校。
2. 数字模板族合并：'Black Knight Captain Huw +1'..'+10' 这类仅数字
   不同的变体合并为一行 'Huw +{0}'，翻译一次，apply 按出现顺序回填数字
   （译文必须保留等量 {0} 占位符，缺失则该变体跳过并报告，绝不丢数字）。
   有术语表预填的变体不合并，保持原有精度。

## 翻译要素与保护

可翻译字段仅两类，其余（AssemblerScript、Address、Offsets、VariableType、

Hotkeys、GroupHeader、Options、Color、ID、Forms、Files、Comment）一律不改：



* **Description**：菜单标签、按钮、标题、状态说明。CE 以一对引号包裹字符串

  字面量，替换保持引号结构、内部做 XML 实体转义。

* **DropDownList**：`ID:名称` 每行一条。只替换冒号后的名称，ID 前缀原样。

  名称可能含 `&`（源中为 `&amp;`），替换后重新转义，保证合法。

  DropDownList 最后一行形如 `ID:名称</DropDownList>`（行尾带结束标签），

  正则显式排除，标签原样保留（随机替换实测曾抓出该边界，已修复）。

分类规则（extract 阶段）：



* `param`（建议保留）：无字母（纯数字 / 纯符号 / 数字符号混合，含空格

  形态如 `[ 0`）；无空格且含下划线 / 数字 / 驼峰（`defPlayerDmgCorrectRate_Dark`）；
  纯 Lua 注释层的代码成分（复合赋值 / 函数调用 / then-return 关键字 /
  十六进制地址）同样归 param，防误翻； Description/DropDown 面向最终
  用户，永不套用代码启发式

* `normal`：短语、物品 / 敌人 / NPC 名、UI 标签 —— 真正需要翻译的部分

自动译文可用性规则（防词表污染）：

combine 词表存在 `Dont`→`Ｄont`（全角）、`C`→`Ｃ`（单字母）等退化条目。

所有自动生成的译文（xlsx 预填 / 备注建议 /fallback 替换）统一要求

**含中文字符且不含全角 ASCII**，不满足即弃用 —— 宁缺毋滥，不产出

不可读文本。

格式要素提示（备注列自动标注）：含数字 /`+`/`%`/`[]` 的词条

（`Godrick (Head & Torso)`、`+10`、`Lv. 50`），翻译时保留数字与符号。

`--prefill-partial` 时备注列附加 "词表建议参考"（短语级最长匹配建议，

只写备注、不填译文列，保持译文列纯净）。

审计结论（对源 CT 全文本扫描，无遗漏翻译形态）：

`$var` 364 个全部位于 AssemblerScript（保护字段）内，Description 含 `$` 为 0；

XML 注释 4 个（itemgib/Get ID 等开发者内部标记，不参与翻译）；

GroupHeader 502 个内容全为 "1"；Description 无字面换行、无 DropDown 分隔线。

按键 / 引号要素核查（用户质询项，实证）：



* `<Caption>`/`<Header>`/`<Title>`/`<Save>` 元素均不存在；

* `<Hotkey>` 共 5 个，子元素仅 Action/Keys/ID，无 Description 文本；

  Keys/Key 为纯数字按键码（保护，不翻译）；

* `<DisassemblerComment>` 4 个，内容仅为空白，无翻译要素；

* Description 引号：提取层条件剥离外层包裹引号（首尾都是引号才剥，

  不与术语表键错位）；文本内嵌引号 4 条（`"save"`/`"Alive"`/`"No"`/

  `"Green Screen"`）作为普通词条导出，替换时转义为 `&quot;` 保证 XML 合法；

  唯一成对引号词条 `"WTF"`（源为双重包裹）替换层兼容匹配，e2e 零失配实证。

## 增量翻译（delta / merge）

版本升级工作流（apply 永远从原始 EN CT 全量重建，绝不打补丁到已译 CT）：

```
eldenct delta <新EN.CT> -o delta.xlsx --zh 旧译.CT --en 旧EN.CT --tm 旧translations.json
   # ① TM 层：新表源文本命中翻译记忆 → 译文列直接预填（绿）
   # ② ID 层：--en/--zh 按 (entry_id, field) 对齐，源文本已变的旧译文 → 备注列"ID继承参考"
   # ③ 词表层：术语表整串命中 + partial 建议
   # 附带 <out>.xlsx.manifest.json（输入路径/命中统计/时间戳）
eldenct import delta.xlsx -o delta.json
eldenct merge 旧translations.json delta.json -o full.json   # delta 优先，冲突报告
eldenct apply <新EN.CT> full.json -o 新译.CT
```

xlsx 新增 H 列「上下文」（词条父级 Description 链，导入不读取，仅供消歧）。

## 防错设计



1. **源文件只读**：apply 输出路径与源 resolve 后相同立即拒绝；产物始终是新文件。

2. **行级字节替换**：不整体序列化 XML。按 sourceline 只改目标行，

   其余字节（缩进、属性顺序、自闭合样式、注释）原样。

3. **行尾 / 编码保真**：读写均 `newline=""`，CRLF/LF 与 BOM 按源原样保留，

   逐字节一致。

4. **XML 合法化**：译文先剔除 XML 非法控制字符，再做实体转义

   （`& < > "`）。替换后以严格模式（recover=False）重新解析，

   结构破坏立即报错。

5. **术语表 fallback 边界**：默认关闭。开启时也只作用于 "xlsx 从未出现过的

   词条 "—— 你在 xlsx 里留空的行永远不被兜底覆盖（留空 = 明确不译）。

6. **导回防错配**：逐行校验 A 列词条 ID 哈希与源文本一致，

   不一致（源列被改）的行忽略并报告。

7. **原子写入**：先写临时文件再 os.replace，中断不留半文件。

8. **替换后校验**：重新解析产物，确认每个译文已生效、源文本零残留，

   并报告未匹配行（missed）与未知源（unknown\_sources）。

## 乱序 diff

以条目 ID（已实测全局唯一）为锚对齐两棵条目树，文件内顺序变化不影响对照：



* `diff CT_A CT_B`：跨版本对照（cn 参考版 vs en 新版），

  报告 changed /same/only\_a /only\_b。

* `cover CT translations.json`：同一 CT 在给定译文集下的翻译覆盖，

  报告 translated /unchanged/unknown。

## 已验证



* `Hexinton-v8.0.4.CT`（7.9 MB，12901 条目，CheatEngineTableVersion 52）

* pipeline 导出 17050 唯一词条：术语表锁定 5726 / 参数建议保留 6049 / 需人工 5275

  （锁定数含 "可用中文译文" 过滤：词表 1489 条全角 / 退化条目被拒，宁缺毋滥）

* apply 全量替换零遗漏（still-unreplaced = 0），行级 diff 仅 38748 行变化

  （全部为目标行），CRLF/BOM 字节级保留，严格解析通过。

* cover 按出现次数计数，translated 与行级替换数严格一致（38748）。

* cn-v4.15.1 与 Hexinton-v8.0.4 对照：changed 18813 /same 27943 /

  only\_a 298 /only\_b 1575（两版非逐条映射，cn 仅作术语参考）。

* 副本随机替换实测（data/\_random\_test.py，可复用）：复制源 → 提取 →

  随机译文（刻意含 `& < > "` 与中文，覆盖 XML 实体转义路径）填入 →

  导回 → 替换 → 逐行断言全绿：missed 0 /unknown 0 /strict 解析通过 /

  改动行数与词条目标行数精确相等（46968==46968）/ 非目标行逐字节一致 /

  DropDown ID 前缀 10/10 保留 / CRLF/BOM 一致 / 源文本零残留。

## 工程借鉴

基于真实调研（ripgrep/ast-grep/sqlite/nginx/redis/ffmpeg 内部原理 + Python

生态各领域顶尖库基准）在真实规模（8 MB XML、1.3 万条目、1.8 万词条）下的取舍：



| 借鉴对象          | 本项目对应                                 | 说明                                                                        |
| ------------- | ------------------------------------- | ------------------------------------------------------------------------- |
| ripgrep       | lxml/libxml2 解析 + pyahocorasick C 自动机 | 解析与多模式匹配都在 C 层，Python 只做编排；全流程 3 秒，并行 / SIMD 无收益（热点在 C 内部已向量化）            |
| ast-grep      | lxml AST + sourceline 源码映射            | 与 ast-grep 的 "结构化 AST 匹配 → 字节区间编辑 → 批量提交" 完全同构：只改目标行，不整体序列化               |
| xlsxwriter    | xlsx 导出写入器                            | 业界专项写入器：实测 17050 行带样式 1.4s，比 openpyxl 逐格样式（约 200s）快两个数量级，是 "不重复造轮子" 的最大落点 |
| sqlite B-tree | dict 哈希术语表（O (1) 精确命中）                | 术语表 3.8 万键用哈希即可；B-tree 是预排序场景的优化，这里无需求                                    |
| nginx 零拷贝     | 行级定点替换                                | 只触碰目标行字节，非目标字节零复制零改写                                                      |
| redis 原子性     | 临时文件 + os.replace                     | 写入要么完整要么不落盘，等同事务提交语义                                                      |
| ffmpeg 分块     | 按行分块 + sourceline 索引                  | 定位是 O (1) 的（lxml 已建行索引），无需额外分区                                            |
| SIMD / mmap   | 不适用                                   | Python 层无法直接利用；8 MB 文件 < L2 缓存量级，mmap 增加复杂度无收益                            |

明确不引入的过度工程：多线程（GIL + 全 C 热点，无收益）、CUDA（文本处理非

计算密集）、内存映射（8 MB 文件无收益）、文本 diff 算法库（difflib 最坏

O (n³)，但我们用的是 ID 对齐的语义 diff，不走行级算法）、新一代 xlsx 库

（wolfxl/fastexcel-rw 声称更快但生态新、样式兼容未验证，xlsxwriter 已是

成熟最优解）。

## 边界决策记录（实证查证结论，防重复踩坑）

以下各项均经实证查证后**有意不做**或**确认已覆盖**，后续迭代无需重复分析。

### 1. Lua 脚本内的地名表（区域选择脚本）——不做

`Hexinton-v8.0.4.CT` Lua 脚本区（约 133200 行起，Invasion Regions/区域选择脚本）的
`{Map = "...", Name = "...", PlayRegion = 1300010, BonfireFlags = {71302}, isBoss = true}` 表：

- `Name`/`Map` 值仅用于显示（复选框 Caption、地图下拉项）；功能匹配是**自引用**的
  （`wex.IndexedRegionNames[t.Caption]` 按显示文本反查同一张表），
  全表一致替换不会破坏传送功能。
- 但收益仅限脚本运行时弹窗（日常操作不走这里），且有三处硬性陷阱：
  1. `isBoss`/`isDungeon`/`isOpenWorld` 等字段名是**功能标识符**——
     `selectRegions()` 用 `v[property]` 按字段名查表，与复选框控件内部名
     （`create("TCheckBox", "isBoss", …)`）联动，绝不可译；
  2. 脚本硬编码 `"Unknown"` / `"Unknown PlayRegion "` 字面量（动态补录未知区域），
     翻译 Map 值必须同步处理，否则地图下拉会出现「未知」与 `Unknown` 两个重复类别；
  3. `Name` 是索引键，译文必须与原文一一对应——两个不同地名译成同一个词会
     覆盖 `IndexedRegionNames` 的键，直接损坏功能。
- 结论：维持 raw 透传，不提取不翻译。若未来要做，需新增「Lua 定向替换」模式
  （白名单只匹配 `Name =`/`Map =` 的值 + Unknown 字面量同步 + 一一映射防重名 +
  替换计数守恒校验），默认不开启。

### 2. DescriptionOnly 传送点下拉——已覆盖，无需动作

形如 `<DropDownList DescriptionOnly="1" DisplayValueAsItem="1">44 B3 E8 44 …:Renala Boss Room`
的下拉（Target-AOB 传送脚本）：hex 键是 `dd (float)` 坐标/指针数据（真正的功能值，
管线按 `ID:名称` 规则原样保留 ID 前缀），冒号后仅为显示名。实测 zh 版 12 条
全部译出（如 `…:蕾娜菈头目房间`、`…:艾尔登之兽`），hex 键零改动。

### 3. Lua 运行时控件 Caption——不做

CE Lua 运行时创建的控件（`createForm`/`createButton`/`win:create(...)` 等），
显示文本是 `Caption` 属性；代码引用走 `Name`（控件标识符，禁译）。
实测本表 Lua 层 Caption 赋值共 21 处，其中静态英文串仅 15 个，全部集中在
Invasion Regions 脚本的一个动态窗口（窗口标题 `Invasion Areas`、分组
`Map`/`Region Types`、过滤钮 `Open World`/`Dungeons`/`Bosses`、
`Save Template`/`Load Template` 按钮等）。这些字符串是 Lua 代码字面量，
翻译需修改 Lua 代码本身（同第 1 项的定向替换范畴），收益/风险不匹配，不做。

### 4. 补充字典 idlist-3.0-dlc.tsv

来源：`Hexinton ID List 3.0 DLC.xlsx`（29 表多语言 ID 列表）。提取 EN→ZH
共 6744 对，格式与 GUI 步骤④「导入 TSV 字典」兼容（`EN\tZH`、无表头、
无引号包裹、源串原样保留；已用 GUI 同款 QUOTE_NONE 解析验证）。

**剔除规则（与管线「宁缺毋滥」同口径）**：

1. 键无字母（数字/符号混排，可含空格）→ 剔除——CT 提取层本就不收无字母词条，
   字典键也不该有；
2. 键已含中文/假名 → 剔除（源已含中文 = 不译，增量口径一致）；
3. 译文不含 CJK → 剔除；
4. 同英文多译冲突（126 键）→ 一律剔除，**不采用首见优先**。冲突分两类，
   平铺字典都无法处理：*片段地名*（DLC Graces 的 `West`/`Eternal City`/
   `Bear Woods`，英文列是片段、完整地名 = 区域前缀 + 片段，`West` 的两个
   「译文」是两个不同地点，首见优先=随机锁死错地点）；*译名分歧*
   （`Golden Vow` → `黄金树立誓` vs `金誓`）。剔除零损失实证：126 键中
   123 个已被 `glossary_qualified.tsv` 以官方译名覆盖，其余 3 个正是片段地名。

**全角归一（增量风格一致性）**：DLC 中文列是全角排版（`【１】`/`＋３`/`Ｄ`），
而已落盘 zh.CT 与 glossary_qualified 均为半角（`[1]`/`+3`，实测全角＋ 0 条、
【】 0 条）。统一归一：全角字母/数字/加号 → 半角、`【１】` → `[1]`、
CJK 后紧跟 `[N]` 补空格（`铃珠[1]` → `铃珠 [1]`）；`（）：，！？` 等 CJK
标点是官方译名组成部分，保留。归一后残留全角 ASCII 的行剔除。
完整冲突清单见 `idlist-3.0-dlc.conflicts.tsv`，人工复核后可按需手工补回。

## 限制



* 假设 Description 为单行（实测 12901 条目全部单行）；多行 Description

  会报 missed，不会破坏结构。

* `combine_conflicts`（一词多译）默认不自动替换，人工在 xlsx 中定夺。

* diff 的字段对齐锚为 `(ID, field)`；跨版本条目 ID 不重合的部分以

  only\_a /only\_b 呈现，不强行对齐。