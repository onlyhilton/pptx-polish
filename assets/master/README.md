# 模板资产目录

存放可直接用于 Step 1「套模板」的模板文件（`.thmx` / `.potx` / `.pptx` 均可）。

## 已入库

| 文件 | 底色 | 大小 | 版式 | 嵌入字体 |
|---|---|---|---|---|
| `2025 Ruijie Reyee PPT Template-20250530.thmx` | 浅色（默认用这份） | 1.47 MB | 16 | 7 |
| `2025 Ruijie Reyee PPT Template-Dark-20250620.thmx` | 深色 | 3.93 MB | 16 | 9 |

## 版式清单（两个模板都是 16 个）

| idx | 浅色版 | 深色版 | 归类 |
|---|---|---|---|
| 1 | General cover | General cover | 封面（默认） |
| 2 | Office cover | Office cover | 封面 |
| 3 | Hotel cover | CCTV cover | 封面 |
| 4 | Retail cover | Retail cover | 封面 |
| 5 | CCTV cover | Hotel Cover | 封面 |
| 6 | Hostel cover | Hostel cover | 封面 |
| 7 | Home cover | home wifi cover | 封面 |
| 8 | Smar home cover | smart home | 封面 |
| 9 | Starlink cover | Starlink | 封面 |
| 10 | 目录 | Index1 | 目录 |
| 11 | Workshop index | Index | 目录变体 |
| 12 | 标题和文本 | Title and content | 标题+内容 |
| 13 | 仅标题 | Title only | 仅标题 |
| 14 | Blank page | blank | 空白页 |
| 15 | Section Title | Section Title | 章节过渡 |
| 16 | 封底 | 封底 | 封底 |

## 使用方式

模板随包内嵌，**命令里不写路径**：

```bash
"$PY" "$S/inject_template.py" --source "<原件>" \
      --out "<工作目录>/1注入.pptx" --sync-size          # 默认 = 浅色底
"$PY" "$S/inject_template.py" --source "<原件>" --theme dark \
      --out "<工作目录>/1注入.pptx" --sync-size          # 深色底
```

`--theme` 依次尝试：真实路径 → 包内裸文件名 → 别名（`default`/`light`/`dark`）→ 唯一子串。
解析由 `inject_template.py` 的 `resolve_theme()` 从 `__file__` 定位完成，**不依赖 cwd，也不用写绝对路径** ——
这正是"模板随包分发"能成立的前提。找不到时会直接把本目录下的可用模板列出来。

**不要按 idx 或名称硬编码映射** —— 两个模板同名版式的 idx 并不一致，命名习惯也不同（浅色版 "General cover" vs 深色版 "smart home"）。一律让脚本读占位符结构现场判定。

## 关键结构事实

- `.thmx` 是标准 OOXML 包，内部基准目录是 `theme/`（`.pptx` 是 `ppt/`）。**从 PowerPoint「设计 → 保存当前主题」导出的 .thmx 含 slideMaster + slideLayouts；从 Word/Excel 导出的不含版式。**
- 版式背景是 **图片填充**（`blipFill` → `ppt/media/` 底图），不是纯色。注入时必须把 media 一起搬。
- 母版本身**没有占位符**，占位符全在 layout 层定义。
- 主题 `clrScheme` 是 Office 默认色，品牌蓝 `0055CD` 硬编码在 layout 装饰形状里。详见 `../../references/brand-color.md`。

## 换模板时

把新的 `.thmx` / `.potx` 放进本目录，跑一次 `analyze_template.py` 看结构，再对照更新上面的版式表即可，不需要改任何代码。

想让新模板也能用**别名**调用，再到 `scripts/inject_template.py` 的 `TEMPLATE_ALIASES` 里登记一行；
不登记也能用 —— 直接传**裸文件名**即可，`resolve_theme()` 会在本目录里找。
