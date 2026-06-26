---
name: maijia-menu-analyse
description: Use when the user asks to fetch Meituan self-service dish data, export complete dish/menu details, analyze Maijia menu performance, run dish-level attribution, or prepare dish catalog and stall mapping from Meituan POS data.
---

# Maijia Menu Analyse

Use this skill to automate the Maijia workflow:

1. Open Meituan POS web backend.
2. Export the correct dish/menu dataset from 报表中心.
3. Validate the `.xlsx` export before analysis.
4. Prepare dish/stall analysis inputs when the source columns support it.

For complete dish-level operating details, use `自助取数 -> 自助菜品取数`.

For stall/档口 attribution, fetch the dish catalog from `运营中心 -> 菜品管理 -> 菜品库`. In Maijia analysis, `档口 = 基础分类`; use `基础分类` from `总部菜品` as the management stall dimension.

The generated report should mirror the structure of the Maijia 5A example only for dimensions supported by the exported data. Do not invent missing dimensions such as 食材分类, 味型归类, or 烹调技法 unless the user provides a mapping table or explicitly asks for inferred labels.

## Browser Export Workflow

Use Chrome when the user needs their existing Meituan login session.

## Raw Export Naming

Save all raw downloaded files under `documents/raw_exports/` with these names:

- Dish sales export: `maijia_dishes_YYYYMMDD_YYYYMMDD.xlsx`
- Dish catalog export: `maijia_dish_catalog_YYYYMMDD.xlsx`
- If a date range is split across multiple exports, append `_part01`, `_part02`, etc. before `.xlsx`.

### Complete Dish Data: 自助菜品取数

Use this workflow when the user asks for "菜品完整信息", "菜品取数", "菜品主题数据", "穿透到菜品", dish-level attribution, or an export such as `maijia_dishes_YYYYMMDD_YYYYMMDD.xlsx`.

1. Open `https://pos.meituan.com/web/report/main#/rms-report/home`.
2. Click top navigation `报表中心`.
3. In the left sidebar, click `自助取数`.
4. In the menu flyout, click `自助菜品取数`.
5. Set the requested date range. Prefer complete business days; avoid current partial-day data unless the user explicitly asks for it.
6. Click `展开筛选`.
7. Select every available field group, including:
   - `查询维度`
   - `菜品销量信息`
   - `菜品关联信息`
   - `菜品成本及服务`
8. Click `查询`; wait until the table refreshes and confirm the date range and row count.
9. Click the report-page `导出` button.
10. In the export confirmation dialog, click `前往下载清单`.
11. In `下载清单记录`, find the matching `菜品主题数据(日期【...】)` row by date range and request time.
12. Wait for `状态` to become `导出完成`, then click that row's far-right `操作` column `下载` link. This is not the report-page `导出` button.
13. If macOS shows a Save panel, save into `documents/raw_exports/` and rename with the standard pattern, for example `documents/raw_exports/maijia_dishes_20260614_20260620.xlsx`.

### Dish Catalog / Stall Mapping: 菜品库基础信息

Use this workflow when the user asks for "菜品库", "菜品基础信息", "档口归因", "基础分类归因", a latest dish catalog, or a stall mapping table to connect with `自助菜品取数`.

1. Open `https://pos.meituan.com/web/operation/main#/`, or click top navigation `运营中心` in the logged-in Meituan session.
2. In the left sidebar, click `菜品管理`.
3. In the menu flyout, click `菜品库`.
4. Select the target brand in the left brand tree:
   - Use `麦家小馆` by default.
   - Use another sub-brand only when the user explicitly requests it.
5. Click the top-right `菜品导出` button.
6. In the export dialog, choose `导出菜品基础信息`.
7. Select `全部字段` so the export includes `基础分类`, `打印出品档口`, `出品部门`, `设置出品部门`, and `预估成本`.
8. Click the bottom-right `确定` button.
9. If macOS shows a Save panel, save into `documents/raw_exports/` and rename with the standard pattern, for example `documents/raw_exports/maijia_dish_catalog_20260626.xlsx`.

This workflow is from `运营中心`, not `报表中心`. It produces a dish dimension table rather than a dated sales fact table.

## Complete Dish Export Shape

Meituan `自助菜品取数` exports may use this worksheet layout:

- Row 1: report title, e.g. `菜品主题数据`.
- Row 2: filter description, including date range and selected fields.
- Row 3: actual header row.
- Row 4 onward: data rows.

When profiling these files, skip the first two rows and treat row 3 as headers. Some exports contain an incorrect OOXML worksheet dimension, so `openpyxl` `max_row` / `max_column` can under-report the sheet as `1 x 1`; if that happens, stream the worksheet XML directly or iterate rows rather than trusting the dimension metadata.

Typical complete dish fields include `下单时间`, `门店`, `菜品名称`, `时段`, `订单分类`, `营业日`, `营业周`, `营业月`, `菜品销售数量`, `菜品销售额`, `菜品优惠`, `菜品收入`, `菜品关联正向订单量`, `退菜数量`, `退菜金额`, `出餐订单数`, and other optional fields when enabled.

## Dish Catalog Export Shape

Meituan `菜品库 -> 导出菜品基础信息` exports commonly include:

- `总部菜品`: the core catalog sheet. Row 1 is usually a title, row 2 is filter/context text, row 3 is the real header row, and row 4 onward contains dish rows.
- `总部套餐`: package composition rows. Use it only when decomposing set meals into component dishes; otherwise report 套餐 separately.

Important `总部菜品` fields:

- `菜品编码（SPUID）`, `菜品编码（SKUID）`, `菜品名称`, `规格名称`
- `品牌`
- `基础分类`
- `售卖价`, `会员价`, `预估成本`
- `打印出品档口`, `出品部门`, `设置出品部门`
- `当前售卖状态`, `创建时间`, `修改时间`

Use `基础分类` as `档口` for management reporting. `打印出品档口`, `出品部门`, and `设置出品部门` are operational routing fields and should not replace `基础分类` unless the user explicitly asks for kitchen routing analysis.

When joining the catalog to `自助菜品取数`, prefer stable code fields such as SPU/SKU or dish code when both files contain them. If codes are unavailable, join on normalized `菜品名称` plus `规格名称` when possible and disclose potential ambiguity from duplicate names, packages, or renamed dishes.

## Analysis Guidance

Use `自助菜品取数` as the dated dish-sales fact table and `菜品库` as the dish dimension table. Join them before making dish/stall conclusions.

Useful report outputs include:

- Store-week dish sales ranking.
- Store-week stall/档口 contribution, using `基础分类`.
- Channel or daypart drilldown by dish and stall.
- Top positive and negative dish contributors for同比/环比 changes.

Do not invent unsupported dimensions. If a requested field is absent from both exports, state the gap and keep the report at the dimensions actually present.

## Validation

After exporting any `.xlsx`:

```bash
file path/to/export.xlsx
unzip -t path/to/export.xlsx
```

For `自助菜品取数`, also confirm the title row, date range row, header row, data row count, and selected field groups. Report only aggregate structure such as row count and columns; do not paste raw dish-level records into chat.

For `菜品库` exports, confirm the presence of `总部菜品`, locate the real header row, and verify that `基础分类` exists before claiming stall/档口 attribution is available. If `总部套餐` is present, mention whether 套餐 rows were decomposed or kept separate.

Before finishing a generated workbook:

1. Confirm the report workbook exists under `documents/`.
2. Open/read the workbook and confirm the sheets are present.
3. Confirm total `销售收入`, `折后毛利`, and row count match the source data excluding the `合计` row.
4. Mention missing dimensions were intentionally omitted when relevant.
