# 碳酸锂产业链

碳酸锂产业链周度数据的公网季节性图谱（独立于钢材站 juanluo-charts）。

- 站点：https://xujiahao88.github.io/lithium-charts/
- 仓库：`xujiahao88/lithium-charts`
- 本地：`C:\Users\Administrator\lithium-charts`

## 数据

来源：`C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\下游周度数据.xlsx`

| tab | 数据集 id | 源 sheet | 口径 | 单位 | 图 | 最新 |
|---|---|---|---|---|---|---|
| 碳酸锂库存 | `lc_inv` | 钢联 | Mysteel 周度库存调研（207 家，货权） | 吨 | 7 | 2026-10-21 |
| 碳酸锂产量 | `lc_output` | 碳酸锂产量 | SMM 周度产量（总计 + 辉石/云母/盐湖/回收料） | 吨 | 5 | 2026-10-15 |
| 正极材料 | `lc_cathode` | 正极 | SMM 三元/磷酸铁锂 产量 + 厂内库存 | 吨 | 4 | 2026-10-15 |
| 动力电芯 | `lc_cell` | 电芯 | SMM 动力电芯产量（合计/铁锂/三元） | GWh | 3 | 2026-10-16 |

> ⚠️ 两套口径不可混读：库存用 **Mysteel**（`lc_inv`），产量与需求端用 **SMM**。
> 「碳酸锂产量」sheet 里另有 4 列 SMM 样本库存（G~J），按需求**未纳入**站点——库存统一以 Mysteel 口径为准，
> 避免同一张图上出现两个来源的"库存"。

## 结构

```
lithium-charts/
├─ index.html
├─ assets/app.js        # 沿用钢材站渲染骨架；已加 4 个数据集的
│                       # GRID_COLS / MONTH_AXIS，以及 ch.nd（小数位）/ summary.nd
├─ assets/style.css
├─ assets/vendor/echarts.min.js
├─ data/data.js         # window.CHART_DATA（内联，file:// 可开）
├─ data/meta.json
├─ data/lc_inv.json     # 钢联库存
├─ data/lc_output.json  # SMM 碳酸锂产量
├─ data/lc_cathode.json # SMM 正极材料
├─ data/lc_cell.json    # SMM 动力电芯
├─ deploy_repo.py       # 推 GitHub（git-data API）
└─ scripts/
   ├─ build_lc_inv_charts.py        # 钢联库存 → lc_inv（含勾稽自检）
   ├─ build_lc_downstream_charts.py # SMM 三表 → lc_output/lc_cathode/lc_cell（含勾稽自检）
   ├─ shot_lithium.py               # Edge headless 出单 tab 长图
   ├─ stack_images.py               # 多张长图纵向拼接（宽度不一自动缩放）
   └─ run_lc_pipeline.py            # 一条龙：build×2 → bump → push → 出图 → 发微信
```

## 图表

**碳酸锂库存（钢联）**
- 组 1「碳酸锂库存（周）」：总量 / 贸易商 / 冶炼厂 / 材料厂及终端
- 组 2「库存结构（派生）」：各环节占总量比重（%）
- 汇总表：本期 / 上期 / 环比 / 较年初 / 较年初%

**碳酸锂产量（SMM）**：总计 / 锂辉石 / 锂云母 / 盐湖 / 回收料，勾稽 总计 = 四项之和

**正极材料（SMM）**：组 1 产量（三元 / 磷酸铁锂）；组 2 厂库存（三元厂 / 铁锂厂）

**动力电芯（SMM）**：合计 / 铁锂 / 三元，勾稽 合计 = 铁锂 + 三元

### 版式（每行几张）

| tab | 每行张数 | 说明 |
|---|---|---|
| 碳酸锂库存 | 4 | `GRID_COLS.lc_inv = 4`，≤1360px 降到 2 列 |
| 碳酸锂产量 | **3** | 用户 2026-10-21 指定；5 张图排成 3+2 |
| 正极材料 | 2 | 两组（产量 / 厂库存）各一行两张 |
| 动力电芯 | **3** | 用户 2026-10-21 指定；3 张图一行排满 |

> 2026-10-21 调整：原 `@media (max-width:1360px)` 会把 `.cols3` 一并压成 2 列，
> 导致笔记本窗口与长图截图（当时 1080 宽）都塌成 2 列。现改为 **`.cols3` 保持 3 列到 1000px**
> （仅 `.cols4` 在 1360 降 2）。长图侧同步把视口放宽到 **1440px / scale 1.5**，
> 输出仍是 2160px 宽，但每卡从 330 → 440 CSS px，X 轴月份标签不再拥挤。

### 横轴口径（两套，注意区别）

- `lc_inv`：MM-DD 季节性轴，日期即源表日期（Mysteel 周三口径，单年起步 → 单条红线）。
- `lc_output` / `lc_cathode` / `lc_cell`：源表 A 列是**下载日**，间隔在 5/6/7/8/14 天间漂移，
  跨年还会左右挪 1~2 天。若直接取 (月,日) 建槽位，4 年叠线会横向错开、轴膨胀到 160+ 槽位
  （碳酸锂产量实测 164 期 → 164 槽）。因此统一按 **ISO 周号** 映射到**基准年（最新年）同周的周四**：
  轴收敛为 52 槽，年份间可比。真实日期仍用于 asOf 与「较年初」。

## 更新

```bash
python scripts/run_lc_pipeline.py                 # 全链（4 个数据集都重建并推站，微信只发库存图）
python scripts/run_lc_pipeline.py --downstream    # 另出「产量+正极+电芯」拼接长图并发送
python scripts/run_lc_pipeline.py --no-send       # 只更新站点与长图
python scripts/run_lc_pipeline.py --no-push       # 不推公网
python scripts/run_lc_pipeline.py --check-only    # 只试算，不落盘
python scripts/run_lc_pipeline.py --force         # 忽略签名去重，强制重跑重发
```

去重签名 = 4 个数据集 asOf 的组合（`lc_inv=…|lc_output=…|…`）。任一数据集换期都会重发，
避免「只更新了 SMM 三表、库存没变」被误判为已发送。

单独重算某个 tab：

```bash
python scripts/build_lc_downstream_charts.py --only lc_cell
```

## 已知源数据问题（站点按原样呈现，未插补）

- 「钢联」sheet 缺 2026-10-09 一周；2026-02-13 与 02-20 两期四列数值完全相同。
- 「碳酸锂产量」总计与四项工艺之和有 2 期差 2 吨（2023-10-06、2023-11-04），已在图注标注。
- 「电芯」合计与铁锂+三元分项有 5 期差 ≤0.25 GWh（2025-10 前后），源表口径差异，未调整。
- 「正极」三元材料产量自 2024-12 起、磷酸铁锂产量自 2025-04 起才有数据（源表前期无该指标）。
- 「正极」2024-01-02 与 2024-01-04 属同一 ISO 周，叠线时保留较晚一期。
