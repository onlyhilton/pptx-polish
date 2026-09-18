# assets/ —— 随包分发的资产

| 路径 | 内容 | 体积 |
|---|---|---|
| `master/` | 可套用的 PPT 模板（`.thmx`）。用法见 `master/README.md` | 5.6 MB |
| `product_images/` | **产品高清图库**（311 个 webp / 17 个系列） | 37.6 MB |
| `product_index.json` | 图库索引（310 条 / 309 型号），`product_lib.py` 读它取图 | 169 KB |
| `product_index_sigs_webp.npz` | 认型号用的特征缓存（`match_product.py`，命中则秒级） | 578 KB |

## product_images/

`product_lib.py` 的默认图库根目录，**包内优先于本机外部库**。

### 目录名不要改

单库模式下 `lib` 名取自目录名。改名会让 `match_product.py` 的特征缓存
（`product_index_sigs_webp.npz`）失配，310 张图要全部重算特征。
`product_lib.py` 的 `SUBLIBS` 里已把 `product_images` 显式钉到 `lib=webp`。

### 为什么换台机器还能找到图

索引里每条记录存两个路径字段：

```
rel  = "RG-ES/RG-ES106D-P V2.webp"     相对本目录 → 跨机器可移植
path = 建库时的绝对路径                 仅本机有效
```

`load()` 返回前过 `rebase()`，用 `<SKILL_DIR>/assets/product_images/<rel>` 把 `path`
重算成本机真实位置。所以整包解压到任何路径都能直接取图，消费方一行都不用改。

判断"要不要重建索引"用 `root_key`（包内记为 `skill:assets/product_images`），
不是绝对路径 —— 否则同一份包换个机器就会被误判成"图库挪位"而重建。

### 更新图库

```bash
# 把新图放进本目录（保持 <系列>/<型号>.webp 的结构），然后重建索引
python scripts/product_lib.py --build
```

不想动包、只想本机临时用外部库：

```bash
python scripts/product_lib.py --find RG-RAP72 --root "<你的图库根目录>/Webp"
# 或设环境变量 RUIJIE_PIC_ROOT
```

未指 `--root` 时，外部库按 `OneDrive` / `OneDriveConsumer` / `OneDriveCommercial`
环境变量推导到 `<OneDrive>/PIC/产品图片`（子路径见 `product_lib.PIC_SUBPATH`），
**不写死盘符**；包内图库始终优先命中。

### 覆盖缺口

包内只有 webp 一支。以下**不在包里**，命中时需从厂商官网取图或自行补入：

- PNG 独有 8 个型号：`RG-EST300F-P`、`RG-NBR6125-E`、`OM-GE-10KM-SFP-SM1490`、
  `SG5200`、`TP-EAP225`、`U6 LITE`、`UAP-AC-M`、`UDM PRO`
- 云桌面终端整条线（CT / CPM / CPK，约 90 个型号）
