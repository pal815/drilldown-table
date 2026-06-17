# drilldown-table — Rust port SPEC (implementation contract)

## Goal
Reimplement the drilldown-table engine in **Rust** as a single static CLI binary so the Claude/Codex
**skill** can generate tables **without Python**. Scope is **xlsx + docx ONLY** (NOT pptx).

## Source of truth
`../drilldown_table.py` — the Python engine. **READ IT FULLY and port it faithfully.** Output must match
the IR fixtures (below). Do NOT read/port `../app/excel_io.py` (Excel COM, clipboard, depth-autodetect,
autosum) — that is app-only and OUT OF SCOPE. The skill invokes
`py drilldown_table.py data.yaml out.ext orient theme`; you reproduce exactly that path.

## CLI
```
drilldown-table <data.{yaml,json}> <out.{xlsx,docx}> [row|column|both] [color|grey|mono]
```
- Defaults: `orient=row`, `theme=color` (match Python `__main__`).
- Output extension picks the renderer. `.pptx` → exit non-zero with:
  `pptx is not supported in the Rust build; use the Python engine.`
- Hidden debug flag **`--dump-ir`**: instead of rendering, print the post-`layout`+`themed` Grid as JSON
  (schema below) to stdout. Used by the test harness.

## What to port (functions in drilldown_table.py)
- **Input**: `load_data` (YAML/JSON) → model dict. Use `serde` + `serde_yaml` + `serde_json`.
- **IR**: `GCell` (fields: `value,text,numfmt,fill,bold,size,halign,indent,wrap,color,top,left,right,bottom`),
  `Grid` (`g`: map (r,c)→GCell, plus `merges, soft_merges, widths, heights, outline_rows, outline_cols, freeze, nr, nc, _depth`).
- **Formatting**: `_disp`/`_fmt` (int/float → `#,##0` via `{:,}`; percent string `"21%"`→`0.21` with fmt
  `0%`/`0.0%`; else text), `_looks_num`. (These are already baked into the IR fixtures.)
- **Color/theme**: `_interp` (hex interpolation, `round`, `%02X`), `_level_fill`, `themed`. Copy the exact
  palette constants (HEAD_FILL, SUB_FILL, C_NAVY, C_TEXT, C_MUTE, C_GRID, C_HEAVY, FONT="맑은 고딕",
  THIN="thin", HEAVY="medium", …) from the engine verbatim.
- **Layout**: `layout` + `_build_row`, `_build_column` (`_build_column_ndepth` / `_build_column_2level`),
  `_build_both`, and helpers `_normalize_items, _nodes_of, _max_depth, _tree_size, _vband, _hband,
  _titles, _merge, _style, _lbl, _edges, _box, _clear_left, _outer, _Al, _grid_borders`.
- **Border reconcile (render-time)**: `reconcile, _region_border, _heaviest, _merge_maps, _is_dark`.
- **Render**: `render_xlsx` → **rust_xlsxwriter**; `render_docx` → **docx-rs**.

## IR JSON schema (`--dump-ir` output AND the oracle in `fixtures/ir/`)
```
{ "nr": int, "nc": int, "depth": int|null,
  "freeze": [r,c]|null,
  "merges": [[r1,c1,r2,c2], ...], "soft_merges": [...],
  "widths": {"<col>": number}, "heights": {"<row>": number},
  "outline_rows": {"<row>": level}, "outline_cols": {...},
  "cells": { "r,c": { "value":num|str|bool|null, "text":str, "numfmt":str|null,
                       "fill":hex6|null, "bold":bool, "size":int, "halign":str|null,
                       "indent":int, "wrap":bool, "color":hex6, "top":"thin"|"medium"|null,
                       "left":..., "right":..., "bottom":... } } }
```
- `--dump-ir` dumps the grid **after `layout`+`themed` but BEFORE `reconcile`** (the fixtures are
  pre-reconcile). `reconcile` runs only inside the renderers.
- `fill`/`color` are 6-hex uppercase strings without `#`, or null. Borders are `"thin"|"medium"|null`.

## Crates
`rust_xlsxwriter`, `docx-rs`, `serde`, `serde_yaml`, `serde_json`. Pure-Rust only (musl-static friendly).

## Rendering specifics
- **xlsx**: per-side borders `thin`/`medium` (colors = C_GRID / C_HEAVY); solid fills (ARGB from `fill`);
  number formats; font `맑은 고딕`; bold; halign + indent + wrap; merge ranges from `grid.merges`;
  freeze panes from `grid.freeze`; column widths from `grid.widths`; hide gridlines.
- **docx**: table with per-cell borders, cell shading (fill), `grid_span`/`v_merge` for merges,
  `RunFonts::east_asia("Malgun Gothic")`, sizes, bold, alignment + indent, fixed layout + column widths.
  (`w:noWrap` is unsupported by docx-rs — OK to skip; note it in TEST_REPORT.md.)

## DONE criteria (verify, don't claim)
1. `cargo build --release` → `target/release/drilldown-table[.exe]`.
2. **All 15 IR fixtures MATCH**: for each `fixtures/ir/<name>_<orient>_<theme>.json`, run
   `target/release/drilldown-table fixtures/inputs/<name>.yaml out.xlsx <orient> <theme> --dump-ir > dump.json`
   then `py fixtures/compare_ir.py dump.json fixtures/ir/<tag>.json` → must print `MATCH`.
   (This proves layout + theming + formatting parity. The names: twolevel{row,column}, both{both},
   multilevel{row,column}, × {color,grey,mono}.)
3. Rendering every tag to `.xlsx` and `.docx` succeeds with no error.
4. Write `TEST_REPORT.md`: pass/fail per fixture + any intentional deviations.

## Constraints
- Release profile: `opt-level="z"`, `lto=true`, `codegen-units=1`, `panic="abort"`, `strip=true`.
- One crate, binary name `drilldown-table`. No pptx / COM / clipboard / excel_io.
- Work autonomously; do not ask questions. Prefer matching the IR exactly over "close enough".
