# 碳酸锂库存

碳酸锂（钢联）周度库存的公网季节性图谱。

- 站点：https://xujiahao88.github.io/lithium-charts/
- 仓库：`xujiahao88/lithium-charts`

## 数据

来源：`C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\下游周度数据.xlsx` → 「钢联」sheet
（我的钢铁网 Mysteel 周度库存调研，样本 207 家：冶炼厂 74 / 贸易商 61 / 材料厂及终端 72，货权口径）

四个分项：碳酸锂库存总量、贸易商库存、冶炼厂库存、材料厂及终端库存（源表单位「吨」，本站直接按吨显示，未换算）。

> 同工作簿的「正极 / 表需 / 电芯 / 碳酸锂产量」等 sheet 是 **SMM** 口径，与本数据集（钢联）不是同一来源，刻意未纳入。

## 结构

```
lithium-charts/
├─ index.html
├─ assets/app.js        # 沿用钢材站（juanluo-charts）渲染骨架，仅加 lc_inv 的
│                       # GRID_COLS / MONTH_AXIS / ROW_META「较年初%」/ 单年tag 四处
├─ assets/style.css
├─ assets/vendor/echarts.min.js
├─ data/data.js         # window.CHART_DATA（内联，file:// 可开）
├─ data/meta.json
├─ data/lc_inv.json     # 唯一数据集
├─ deploy_repo.py       # 推 GitHub（git-data API）
└─ scripts/
   ├─ build_lc_inv_charts.py   # 解析「钢联」→ 数据集（含勾稽自检）
   ├─ shot_lithium.py          # Edge headless 出长图
   └─ run_lc_pipeline.py       # 一条龙：build → bump → push → 出图 → 发微信
```

## 图表

- 组 1「碳酸锂库存（周）」：总量 / 贸易商 / 冶炼厂 / 材料厂及终端，4 张
- 组 2「库存结构（派生）」：冶炼厂 / 材料厂 / 贸易商 占总量比重（%），3 张
- 顶部汇总表：本期 / 上期 / 环比 / 较年初 / 较年初%（单位吨）

横轴为 MM-DD 季节性轴：2026 年起步只有 1 年 → 单条红线；明年起自动变为 2 年对比，无需改代码。

## 更新

```bash
python scripts/run_lc_pipeline.py                 # 全链（默认发微信）
python scripts/run_lc_pipeline.py --no-send       # 只更新站点与长图
python scripts/run_lc_pipeline.py --check-only    # 只试算，不落盘
```
