# maijia-menu-analyse

Agent Skill for exporting complete dish/menu data from Meituan POS and preparing Maijia-style dish and stall attribution inputs when the source fields support them.

The installable skill lives at:

```text
skills/maijia-menu-analyse/
```

That directory is the portable unit: its root contains `SKILL.md`, plus the script files the agent can use.

It is intended for prompts such as:

```text
/maijia-menu-analysis 帮我分析过去三个月的菜单，生成 5A 报告。
帮我去美团自助菜品取数下载完整菜品信息，文件名按规范保存为 maijia_dishes_YYYYMMDD_YYYYMMDD.xlsx。
帮我导出最新菜品库，用基础分类做档口映射。
基于菜品主题数据帮我做菜品穿透分析。
```

The skill automates the workflow from Meituan POS report export to a workbook with these available dimensions:

- `2A_总览`
- `A1_价格区间分析`
- `A2_毛利区间分析`
- `完整菜品清单`

It does not invent missing 5A dimensions such as ingredient category, flavor type, or cooking technique unless the user provides a mapping table or explicitly asks for inferred labels.

## Install

Clone this repository first:

```bash
git clone https://github.com/YinXiaoyu-1998/maijia-menu-analyse.git
```

Then install the `skills/maijia-menu-analyse` folder into the skills directory used by your agent.

### OpenClaw

```bash
openclaw skills install ./maijia-menu-analyse/skills/maijia-menu-analyse --global
openclaw skills check
```

To install for one configured OpenClaw agent instead of globally:

```bash
openclaw skills install ./maijia-menu-analyse/skills/maijia-menu-analyse --agent <agent-id>
```

### Claude Code

Personal install:

```bash
mkdir -p ~/.claude/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse ~/.claude/skills/
```

Project install:

```bash
mkdir -p .claude/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse .claude/skills/
```

### Cursor

User-level install:

```bash
mkdir -p ~/.cursor/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse ~/.cursor/skills/
```

Project-level install:

```bash
mkdir -p .cursor/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse .cursor/skills/
```

### Codex

```bash
mkdir -p ~/.codex/skills
cp -R maijia-menu-analyse/skills/maijia-menu-analyse ~/.codex/skills/
```

For other agents, copy `skills/maijia-menu-analyse` into the agent's configured skills directory, preserving this structure:

```text
maijia-menu-analyse/
├── SKILL.md
├── agents/
│   └── openai.yaml
└── scripts/
    └── generate_menu_report.py
```

Restart or refresh the agent's skills list after installing if the agent does not hot-reload skills.

## Compatibility

This is a standard `SKILL.md`-based Agent Skill package. It is not limited to Codex.

- Anthropic describes Agent Skills as an open standard and notes that the same format can work across tools that adopt it.
- Claude Code supports skills stored at `~/.claude/skills/<skill-name>/SKILL.md` or `.claude/skills/<skill-name>/SKILL.md`.
- OpenClaw supports installing a local skill directory whose root contains `SKILL.md`.
- Cursor supports Agent Skills in the editor and CLI, with skills defined in `SKILL.md` files.

Agent implementations still differ in their exact install paths, available browser automation tools, and Python/runtime setup. The workbook generator itself is a plain Python script, so it can also be run manually from any checkout.

References:

- Anthropic Skills overview: https://support.claude.com/en/articles/12512176-what-are-skills
- Claude Code Skills docs: https://docs.claude.com/en/docs/claude-code/skills
- OpenClaw skills CLI docs: https://docs.openclaw.ai/cli/skills
- Cursor 2.4 changelog: https://cursor.com/changelog/2-4

## Requirements

- A Meituan POS account already logged in through Chrome when browser export is needed.
- Python 3 with `openpyxl` for local workbook validation and analysis.
- A Meituan exported `.xlsx` from `报表中心 -> 自助取数 -> 自助菜品取数` for complete dish-level data.
- A Meituan exported `.xlsx` from `运营中心 -> 菜品管理 -> 菜品库 -> 菜品导出 -> 导出菜品基础信息` when stall/档口 mapping is needed. In this workflow, `档口 = 基础分类`.

## Raw Export Naming

Save all raw downloaded files under `documents/raw_exports/`:

| Export | File name |
|---|---|
| `自助菜品取数` / `菜品主题数据` | `maijia_dishes_YYYYMMDD_YYYYMMDD.xlsx` |
| `菜品库` / `导出菜品基础信息` | `maijia_dish_catalog_YYYYMMDD.xlsx` |

For split downloads, append `_part01`, `_part02`, etc. before `.xlsx`.

## Fetch Complete Dish Data

Use Chrome with an already logged-in Meituan session:

1. Open `https://pos.meituan.com/web/report/main#/rms-report/home`.
2. Go to `自助取数 -> 自助菜品取数`.
3. Set the date range, expand filters, and select all field groups.
4. Query, export, go to `下载清单记录`, and click the matching row's far-right `下载`.
5. Save under `documents/raw_exports/`, for example `documents/raw_exports/maijia_dishes_20260614_20260620.xlsx`.

The export usually has two metadata rows before the real header: row 1 is the title, row 2 is the filter description, row 3 is the actual header row.

## Fetch Dish Catalog / Stall Mapping

Use this when a report needs stall/档口 attribution:

1. Open `https://pos.meituan.com/web/operation/main#/` or click `运营中心`.
2. Go to `菜品管理 -> 菜品库`.
3. Select the target brand, usually `麦家小馆`.
4. Click `菜品导出`.
5. Choose `导出菜品基础信息`.
6. Select `全部字段`.
7. Confirm and save under `documents/raw_exports/`, for example `documents/raw_exports/maijia_dish_catalog_20260626.xlsx`.

The catalog usually contains `总部菜品` and `总部套餐`. Use `总部菜品.基础分类` as the stall/档口 dimension; keep `打印出品档口` and `出品部门` as kitchen-routing fields unless the user asks for that view.

## Analysis Outputs

Use `自助菜品取数` as the dated dish-sales fact table and `菜品库` as the dish dimension table. Typical outputs include dish sales ranking, store-week stall contribution using `基础分类`, and dish/stall同比 or 环比 attribution.

## License

MIT
