# maijia-menu-analyse

Codex skill for exporting menu cost/gross-profit data from Meituan POS and generating a Maijia-style menu analysis workbook.

It is intended for prompts such as:

```text
/maijia-menu-analysis 帮我分析过去三个月的菜单，生成 5A 报告。
```

The skill automates the workflow from Meituan POS report export to a workbook with these available dimensions:

- `2A_总览`
- `A1_价格区间分析`
- `A2_毛利区间分析`
- `完整菜品清单`

It does not invent missing 5A dimensions such as ingredient category, flavor type, or cooking technique unless the user provides a mapping table or explicitly asks for inferred labels.

## Install

Clone this repository, then copy the skill folder into your Codex skills directory:

```bash
git clone https://github.com/YinXiaoyu-1998/maijia-menu-analyse.git
mkdir -p ~/.codex/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse ~/.codex/skills/
```

Restart Codex or refresh the skills list after installing.

## Requirements

- A Meituan POS account already logged in through Chrome when browser export is needed.
- Python 3 with `openpyxl` for local workbook generation.
- A Meituan exported `.xlsx` from `报表中心 -> 菜品报表 -> 菜品成本毛利统计`.

## Generate From An Existing Export

```bash
python3 ~/.codex/skills/maijia-menu-analyse/scripts/generate_menu_report.py \
  --input /path/to/成本毛利表.xlsx \
  --output /path/to/documents/菜单分析报告.xlsx
```

The script creates the output folder when it does not already exist.

## Expected Source Columns

The Meituan export must contain:

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

## License

MIT
