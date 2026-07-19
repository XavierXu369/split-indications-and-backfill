---
name: split-indications-and-backfill
description: Inspect and process a complete Excel molecule, product, or asset pool when approved indications must be mechanically split into one row per nonblank item and every original field must be formula-backfilled. Use when Codex must preserve a complete source sheet, pause after validation and 1–2 previews, then create and verify a three-sheet XLOOKUP workbook without mapping, normalization, or medical interpretation.
---

# Split Indications And Backfill

Create a traceable row-level indication workbook from a complete source pool.

Before acting, read [references/input-output-contract.md](references/input-output-contract.md). Use [scripts/create_split_workbook.ps1](scripts/create_split_workbook.ps1) for inspection, preview, generation, and verification.

## Workflow

Keep the three stages separate. Do not generate the workbook during inspection or preview.

1. Obtain the input workbook, complete source-sheet name, output directory, stable key field, entity-name field, indication field, and approved delimiters.
2. Run the script with `-Mode Inspect`. Report the source fingerprint, effective dimensions, required-field checks, key checks, blank indications, delimiter anomalies, duplicate items, extra/hidden sheets, merged cells, and source-sheet dependencies.
3. Stop on any blocker. Do not invent a key, delimiter, or source field.
4. Run `-Mode Preview -PreviewCount 2`. Show 1–2 representative split examples and the formula example. Preserve source wording; do not map, normalize, deduplicate, or medically interpret indications.
5. Wait for explicit user approval.
6. Run `-Mode Generate` with the approved output path plus the input SHA-256 and source signature returned by inspection. Never overwrite an existing file.
7. Use the script's reopened-workbook verification result. Report counts, warnings, formula coverage, source preservation, and the output path.

## Core rules

- Require a complete source sheet. A three-column helper sheet is review-only and cannot supply all backfilled fields.
- Default fields are `序号`, `药品`, and `获批适应症`. Accept alternatives only when explicitly supplied or unambiguously confirmed.
- Default delimiters are English and full-width semicolons (`;` and `；`). Split nothing else without approval.
- Retain one blank split row when the original indication is blank.
- Preserve duplicate indication items as separate rows and flag them; do not silently deduplicate.
- Preserve the stable key's underlying Excel value type and number format. Do not convert numeric keys to display text.
- Treat duplicate or blank keys, blank/duplicate headers, merged source cells, delimiter-only content, ambiguous delimiters, source dependencies on sheets that would be deleted, fingerprint drift, existing output, or verification failure as hard stops.
- Never modify the input workbook.

## Output

Generate exactly three sheets:

1. The complete source sheet, unchanged.
2. `拆分结果`: stable key, entity name, and `适应症拆分结果`.
3. `更新版拆分底稿`: the same first three fields, followed by every remaining source field in source order. Retain the original full indication field.

Backfill every field after column C with XLOOKUP formulas. Preserve blanks and number formats. Do not replace formulas with static values.

## Commands

```powershell
$script = ".\scripts\create_split_workbook.ps1"

& $script -Mode Inspect `
  -InputPath "C:\path\input.xlsx" `
  -SourceSheet "完整分子池" `
  -IdColumn "序号" -NameColumn "药品" -IndicationColumn "获批适应症" |
  Format-List *

& $script -Mode Preview `
  -InputPath "C:\path\input.xlsx" `
  -SourceSheet "完整分子池" `
  -IdColumn "序号" -NameColumn "药品" -IndicationColumn "获批适应症" `
  -PreviewCount 2 -CompleteSourceConfirmed |
  Format-List *

& $script -Mode Generate `
  -InputPath "C:\path\input.xlsx" `
  -OutputPath "C:\path\output.xlsx" `
  -SourceSheet "完整分子池" `
  -IdColumn "序号" -NameColumn "药品" -IndicationColumn "获批适应症" `
  -CompleteSourceConfirmed `
  -ExpectedInputSha256 "<inspect result>" `
  -ExpectedSourceSignature "<inspect result>" |
  Format-List *
```

If desktop Excel COM automation is unavailable, report that the deterministic formula-backfill workflow is blocked. Do not silently substitute another output design.
