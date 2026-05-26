---
name: maijia-menu-analyse
description: Export Meituan POS report-center menu cost/gross-profit data and generate a Maijia-style menu analysis workbook. Use when the user asks for `/maijia-menu-analysis`, Maijia menu analysis, Meituan menu cost export, past month/quarter menu analysis, or a 5A-like report from Meituan POS data.
---

# Maijia Menu Analyse

Use this skill to automate the Maijia workflow:

1. Open Meituan POS web backend.
2. Navigate to 报表中心 -> 菜品报表 -> 菜品成本毛利统计.
3. Set the requested 营业日期 range, query, and export the `.xlsx`.
4. Generate a menu analysis workbook from the exported cost/gross-profit table.

The generated report should mirror the structure of the Maijia 5A example only for dimensions supported by the exported data. Do not invent missing dimensions such as 食材分类, 味型归类, or 烹调技法 unless the user provides a mapping table or explicitly asks for inferred labels.

## Browser Export Workflow

Use Chrome when the user needs their existing Meituan login session.

1. Open `https://pos.meituan.com/web/operation/main#/`.
2. Click top navigation `报表中心`.
3. In the left sidebar, click `菜品报表`.
4. In the menu flyout, click `菜品成本毛利统计`.
5. Set `营业日期`:
   - Interpret "过去三个月" as the last three calendar months ending today unless the user specifies exact dates.
   - Use absolute dates in the page fields, e.g. `2026/02/26` to `2026/05/26`.
6. Click `查询`; wait until the result table updates.
7. Click `导出`.
8. If macOS shows a Save panel, save the file, then move it into the current workspace `documents/` folder. Create the folder if missing.

Always verify that the exported file exists and that the visible query range matches the requested date range before generating the report.

## Report Generation

Use `scripts/generate_menu_report.py` after export.

```bash
python3 /path/to/maijia-menu-analyse/scripts/generate_menu_report.py \
  --input /path/to/成本毛利表.xlsx \
  --output /path/to/documents/菜单分析报告.xlsx \
  --title "永杰厚道 菜单分析报告"
```

Prefer the agent's bundled Python runtime when available. Otherwise use `python3`:

```bash
python3 scripts/generate_menu_report.py --input ... --output ...
```

Make sure `openpyxl` is installed in whichever Python runtime is used.

## Output Structure

Generate these sheets:

- `2A_总览`: KPI row plus grouped summaries.
- `A1_价格区间分析`: Bucket by `折后均价(元)`.
- `A2_毛利区间分析`: Bucket by `折后毛利率`.
- `完整菜品清单`: Source rows enriched with price bucket, margin bucket, revenue share, and key metrics.

Use these source columns from the exported cost/gross-profit sheet:

- `品项名称`
- `品项类型`
- `菜品编码`
- `单位`
- `销售数量`
- `销售收入(元)`
- `折后均价(元)`
- `菜品预估成本(元)`
- `折后毛利(元)`
- `折后毛利率`
- `优惠金额(元)`

If a column is missing, stop and explain which required column is absent.

## Bucketing Rules

Price buckets use `折后均价(元)`:

`未定价`, `<10元`, `10-20元`, `20-30元`, `30-40元`, `40-50元`, `50-60元`, `60-80元`, `80-100元`, `100-150元`, `>=150元`.

Margin buckets use `折后毛利率`:

`无毛利率`, `<50%`, `50-60%`, `60-70%`, `70-80%`, `80-85%`, `85-90%`, `90-95%`, `>=95%`.

Core totals:

- `销售收入 = sum(销售收入(元))`
- `菜品预估成本 = sum(菜品预估成本(元))`
- `折后毛利 = sum(折后毛利(元))`
- `折后毛利率 = 折后毛利 / 销售收入`
- `销售占比 = group 销售收入 / total 销售收入`

## Validation

Before finishing:

1. Confirm the report workbook exists under `documents/`.
2. Open/read the workbook and confirm the sheets are present.
3. Confirm total `销售收入`, `折后毛利`, and row count match the source data excluding the `合计` row.
4. Mention missing dimensions were intentionally omitted when relevant.
