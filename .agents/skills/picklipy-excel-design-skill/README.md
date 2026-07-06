# PickliPy Excel Design Skill

A Codex/OpenAI Agent Skill for designing and generating Excel design files for `PickliPy.Assay` and `PickliPy.Screen` / `PicklPy.Screen`.

## Install locally for Codex

Copy this folder to either a repository-scoped location:

```text
<repo>/.agents/skills/picklipy-excel-design/
```

or a user-scoped location:

```text
$HOME/.agents/skills/picklipy-excel-design/
```

Restart Codex if the skill is not discovered.

## Main files

- `SKILL.md`: trigger description and operating instructions for Codex.
- `scripts/picklipy_design_builder.py`: workbook builder, validator, volume estimator, and blacklist CSV generator.
- `references/`: schema, layout recipes, and troubleshooting.
- `examples/`: JSON specs for Assay and Screen designs.

## Smoke test

```bash
python scripts/picklipy_design_builder.py assay-demo --output /tmp/assay_demo.xlsx
python scripts/picklipy_design_builder.py screen-demo --output /tmp/screen_demo.xlsx --slots 12 --dst-barcodes P1,P2,P3
python scripts/picklipy_design_builder.py validate /tmp/assay_demo.xlsx --mode assay
python scripts/picklipy_design_builder.py validate /tmp/screen_demo.xlsx --mode screen
```
