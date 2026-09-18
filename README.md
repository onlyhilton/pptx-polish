# pptx-polish

把一份**粗糙的、别人发来的** PowerPoint 改造到可对外使用 —— 原位改造，不重新生成。

**五步流水线 + 一条收口红线**：套模板版式 → 换高清配图 → 规范排版 → 收敛配色 → 标题单行收口。
每一步只解决一类问题，且**模板本体只读**：只能"用模板格式改 PPT"，不能"用 PPT 的现状改模板"。

> **English TL;DR** — A five-step pipeline that repairs an existing `.pptx` **in place**
> (never regenerates): apply a slide-master template, swap in high-resolution product
> images, normalize layout, collapse the palette to brand colors, and force every title
> onto a single line by scaling its rendered font size. The slide master, layouts and
> theme are treated as **read-only ground truth** and verified byte-for-byte before
> delivery. Windows + desktop PowerPoint is required for rendering and for the
> "does this title wrap" check. Proprietary license — see [LICENSE](LICENSE).

---

## 1. 它解决什么

团队收到的 PPT 通常长这样：标题用自由文本框随手拉的、字体东一个西一个、配图是截图糊图、
颜色十几二十种、标题太长折成两行压在分隔线上。这些问题**结构上全都合法**，
所以常规体检一律报 `P0 = 0` —— 但人一眼就看得出不对。

这个 skill 就是补上这一段"结构全对、观感不对"的修复能力，并且**每一步都有可判定的交付标准**。

| 步骤 | 解决什么 | 交付标准 |
|---|---|---|
| **0 结构体检**（可选前置） | 长期迭代堆积的重复 / 从未被引用的 slide layout | 版式面板干净 |
| **1 规范化** | 标题乱拉文本框、字体乱用 → 换模板必错位 | 不好看，但**规整**；模板保真 P0 = 0 |
| **2 配图** | 纯文字或乱配图；图片糊、被拉伸、残留空图 | 图文匹配、图片整齐、比例正确 |
| **3 排版** | 元素乱放、留白被侵占 | 母版安全区干净、排列有范式 |
| **4 配色** | 颜色多且杂、不符合 VI | 单页 ≤ 3 色、品牌色优先 |
| **5 标题单行收口** | 标题过长折成两行，第二行穿过标题下分隔线 | **每个标题都是一行**（放不下就缩字号） |

**Step 5 是收口步骤** —— 必须在内容与版面不再变动之后跑；任何改文字、改字号、改框宽的动作
都会让它的结论失效。

## 2. 红线（R1–R7）

这七条是硬约束，不是建议。其中 **R7 是这套流程存在的理由**。

| 编号 | 红线 |
|---|---|
| R1 | 页序与页数不变 |
| R2 | 内容事实零改写（数据、型号、参数、日期、客户名逐字沿用） |
| R3 | 信息不删（可合并同类项，不得整条丢弃） |
| R4 | 真实图优先（生图只用于概念图/氛围图，禁止替代真实数据图与产品截图） |
| R5 | 原件只读（输出新文件） |
| R6 | 留白区不可侵犯 |
| **R7** | **模板本体只读** —— `ppt/slideMaster*` / `ppt/slideLayout*` / `ppt/theme/*` 是模板格式的唯一事实源。不改版式的形状、几何、占位符槽位（含 `p:ph` 的 `sz`）、`defRPr` 的 `sz/b/i/u/spc/kern` 与 `latin/ea/cs`、颜色、背景图；不删版式、不删母版、不换主题。**交付前必须跑 `verify_template_fidelity.py` 且 P0 = 0** |

> R7 来自一次真实事故：产物里 `Section Title` 版式的 `01` 占位符从 `<a:defRPr sz="13800" i="1">`
> 被改成 `sz="13800" b="0" i="0"`（斜体没了），主题被换掉，16 个版式只剩 8 个 ——
> 而常规体检**全过、P0 = 0**（它拿"已经改过的版式"当权威值，这一维根本不查）。
> 用户第一眼就看出了"斜体被取消了"。

## 3. 安装

```bash
git clone <本仓库> ~/.workbuddy/skills/pptx-polish
```

WorkBuddy 会自动发现 `~/.workbuddy/skills/` 下的 skill。仓库根目录就是 skill 根目录，
`SKILL.md` 是给 Agent 读的完整说明书（含每一步的工具链、判定依据与二十条实战陷阱）。

### 依赖

| 依赖 | 说明 |
|---|---|
| Python 3.9+ | 无第三方安装位置要求 |
| `python-pptx`、`lxml`、`Pillow` | 核心依赖 |
| **Windows + 桌面版 PowerPoint** | `render_deck.py`（逐页渲染 PNG）、`verify_title_lines.py`（读 `TextRange.Lines().Count`）、`open_test.py`（真机打开验收）走 COM。**没有 PowerPoint 就没有"标题是不是单行"的结论** |

```bash
python -m pip install python-pptx lxml Pillow
```

### 命令前置（后面所有命令都假定这两个变量已设好）

```bash
SKILL="${SKILL_DIR:-$HOME/.workbuddy/skills/pptx-polish}"
S="$SKILL/scripts"
PY="${PY:-python}"        # 需已装 python-pptx / lxml / Pillow；缺模块就把它指到你的解释器
```

## 4. 用法

分步交付，每步一个独立可用的文件：

```
原名_1规范化.pptx → 原名_2配图.pptx → 原名_3排版.pptx → 原名_4配色.pptx → 原名_5标题单行.pptx
```

### 4.1 先判断：这份 PPT 是否已经套过模板

```bash
"$PY" "$S/inject_template.py" --source "<原件>" --outline      # 只读，列出结构与已有版式
```

### 4.2 Step 1 规范化

```bash
"$PY" "$S/inject_template.py" --source "<原件>" --out "<工作目录>/1注入.pptx"
"$PY" "$S/normalize_deck.py"  "<工作目录>/1注入.pptx" --out "<工作目录>/1规范化.pptx"
"$PY" "$S/verify_pptx.py"     "<工作目录>/1规范化.pptx" --out "<工作目录>/verify"
"$PY" "$S/verify_template_fidelity.py" "<工作目录>/1规范化.pptx"   # R7：P0 必须 = 0
"$PY" "$S/open_test.py"       "<工作目录>/1规范化.pptx"           # 真机打开验收
```

### 4.3 Step 2 配图

```bash
"$PY" "$S/match_product.py"  "<工作目录>/1规范化.pptx"            # 认出 PPT 里的产品型号
"$PY" "$S/replace_image.py"  "<工作目录>/1规范化.pptx" --out "<工作目录>/2配图.pptx" --new-model "<型号>"
"$PY" "$S/fix_geometry.py"   "<工作目录>/2配图.pptx" --fix-aspect --out "<工作目录>/2b.pptx"
```

产品图库随包分发（`assets/product_images/`，310 条 / 309 型号），索引见
`assets/product_index.json`。要跟随自己的外部图库，用 `--root` 或 `RUIJIE_PIC_ROOT`。

### 4.4 Step 3 / 4 排版与配色

```bash
"$PY" "$S/render_deck.py"    "<工作目录>/2b.pptx" --out "<工作目录>/render"   # 渲染后在图上看
"$PY" "$S/recolor_deck.py"   "<工作目录>/3排版.pptx" --report --out-dir "<dir>"   # 先普查（含角色分布）
"$PY" "$S/recolor_deck.py"   "<工作目录>/3排版.pptx" --map "<配方>" --out "<工作目录>/4配色.pptx"
```

配色替换**必须按角色分治**：`title`（标题专用色，由版式槽位决定）**不参与**"统一成品牌色"的映射。

### 4.5 Step 5 标题单行收口

```bash
"$PY" "$S/fix_title_oneline.py"  "<工作目录>/4配色.pptx" --check --json "<工作目录>/oneline.json"
"$PY" "$S/fix_title_oneline.py"  "<工作目录>/4配色.pptx" --out "<工作目录>/5标题单行.pptx"
"$PY" "$S/verify_title_lines.py" "<工作目录>/5标题单行.pptx"    # 违规必须 = 0
```

缩的是**渲染字号**（页面级 `<a:normAutofit fontScale>`），不是声明字号 ——
所以 R7 模板保真照样 P0 = 0，且以后标题改短了字号能自然恢复。

## 5. 仓库结构

```
pptx-polish/
├── SKILL.md               # Agent 说明书：五步工具链、判定依据、二十条实战陷阱
├── README.md              # 本文件
├── LICENSE                # 专有许可（保留所有权利）
├── .gitignore
├── scripts/               # 26 个脚本
│   ├── inject_template.py           # Step 1.1 注入模板母版体系 + 逐页重映射版式
│   ├── normalize_deck.py            # Step 1.2 标题归位 + 字体继承主题
│   ├── replace_image.py             # Step 2   换高清图（图片框零改动）
│   ├── match_product.py             # Step 2   认出 PPT 里那张图是哪个型号
│   ├── fix_geometry.py              # Step 2.7 消变形、删残留空图
│   ├── recolor_deck.py              # Step 4   配色普查（按角色）与品牌色替换
│   ├── pptx_title_scope.py          # Step 5   「哪些占位符算标题」的唯一定义
│   ├── fix_title_oneline.py         # Step 5   标题单行收口
│   ├── verify_title_lines.py        # Step 5   走 PowerPoint COM 数真实行数
│   ├── verify_template_fidelity.py  # R7       以模板为唯一事实源的保真体检
│   ├── verify_pptx.py               # 包完整性 + 标题字体/字号/颜色是否回归版式槽位
│   ├── open_test.py                 # 真机打开验收（抓 python-pptx 漏掉的包级错误）
│   ├── render_deck.py               # 逐页渲染 PNG
│   ├── product_lib.py               # 产品图库索引（跨机器可移植）
│   ├── pack_skill.py                # 打包 + 五项可移植性自检
│   └── …                            # 其余为诊断与辅助工具
├── references/            # 品牌色规范、版式范式、审计清单、报告模板
└── assets/
    ├── master/            # 两个 .thmx 模板（浅色 / 深色）
    └── product_images/    # 随包分发的产品图库（311 张 webp）
```

工具链里刻意保留**多层独立体检**，因为它们回答的不是同一个问题：

| 工具 | 权威值来自 | 回答什么 |
|---|---|---|
| `verify_pptx.py` | 产物**自己的**版式 | 包结构对不对、标题有没有回归版式槽位 |
| `verify_template_fidelity.py` | **模板文件** | R7：模板有没有被改 |
| `verify_title_lines.py` | **PowerPoint 排版引擎** | 标题到底换没换行 |
| `open_test.py` | **PowerPoint 本体** | 这个包能不能被打开 |

**结论相同 ≠ 产物相同。** 判断"改了什么"要看解压后的部件级字节差异，不要看 zip 体积 ——
重写 pptx 会把全部部件重新压缩，体积变化与内容变化无关。

## 6. 资产说明

仓库内含 Reyee 品牌 PPT 模板（`.thmx`）与产品官方图（311 张）。这些是**品牌资产**，
作为工具的运行资源随包分发；[LICENSE](LICENSE) 的专有条款**不授予**任何人对这些品牌资产
的提取、再分发或独立使用权。

## 7. 已知限制

- **渲染与行数复核依赖 Windows + 桌面版 PowerPoint**。缺此环境时 Step 5 只能给出
  "XML 度量预测"，**不得声称"已确认单行"**；脚本在 COM 不可用时会以退出码 1 明确报错，不会静默通过。
- Step 5 的宽度度量需要标题字体文件（`Noto Sans bold`）。找不到时退化为按字宽估算并把余量翻倍，
  报告里会标记"度量不精确" —— 不要当精确结果用。可用 `--font-map` 指定字体。
- 模板 16 个版式中的封面按行业场景拆分（零售 / 酒店 / 办公 / 校园 / 家庭 …），
  选择依据是内容场景，不是页面顺序。
- 本流程**不使用**任何"重建页面"路线（那会丢掉原件结构，与"规范化"的目标相反）。

## 8. 许可

**Proprietary — All rights reserved.** 详见 [LICENSE](LICENSE)。
本仓库公开可见**不等于**授权使用。
