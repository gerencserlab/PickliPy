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

## Run a randomized Screen

Screen workbooks may contain an optional DST `Layout:` map. Identical layout labels are runtime-randomization boundaries and are separate from `DST_Blacklist!Groups:`.

```bash
python -m PickliPy.Screen -f /tmp/screen_demo.xlsx --randomize-wells true --no-pause
```

The builder's `slot_layout` JSON accepts `randomization_layout_mode` values `none`, `all`, `halves`, or `quadrants` and pins generated controls by default; custom layouts accept a `randomization_layout` well map or 16 × 24 grid. The presets split the full 16 × 24 canvas, so use a custom map for 96-well half/quadrant boundaries.

## Run with operational backfill

An unused `Backfill | vehicle | mock final concentration` Labels row derives a target total transfer volume. A paired CLI override takes priority:

```bash
python -m PickliPy.Assay -f assay.xlsx --backfill-name vehicle --backfill-volume 100 --no-pause
python -m PickliPy.Screen -f screen.xlsx --backfill-name vehicle --backfill-volume 100 --no-pause
```

The builder examples `assay_backfill_spec.json` and `screen_backfill_spec.json` demonstrate workbook configuration. Screen `slot_layout` specs can append the reserved row with `layout.extra_labels`.
