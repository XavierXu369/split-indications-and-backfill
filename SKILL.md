---
name: split-indications-and-backfill
description: Validate and process an Excel molecule, product, or asset pool when approved indications must be split from semicolon-separated text into one row per indication. Use when Codex must preserve a complete source sheet, show 1–2 examples for approval, and create a three-sheet workbook that formula-backfills all original fields by XLOOKUP.
---

# Split Indications And Backfill

Use this skill to create a traceable, row-level indication workbook without manually recreating the full source dataset.

Before acting, read [references/input-output-contract.md](references/input-output-contract.md).

## Required interaction

Follow this sequence. Do not skip the confirmation gates.

1. **Request the input.** Ask the user for one Excel workbook, the full source-sheet name, and the output directory. The source sheet must contain a unique identifier, an entity-name field, an approved-indication field, and all fields that need to survive downstream.
2. **Validate read-only.** Confirm the workbook and source sheet open; required fields exist; the identifier is nonblank and unique; entity names exist; count blank indications and rows containing the confirmed delimiter; and confirm that the source sheet is a complete pool rather than a three-column extract.
3. **Report validation and pause.** State whether the input passes, list blank-indication entities, and say which source sheet will provide the formula backfill. If a blocking issue exists, stop and ask for a decision.
4. **Show 1–2 samples.** Show representative split rows and one XLOOKUP example. Preserve all original wording; do not map, normalize, or medically interpret an indication at this stage.
5. **Wait for explicit confirmation.** Only after the user approves the examples, generate the workbook.
6. **Generate and verify.** Run [scripts/create_split_workbook.ps1](scripts/create_split_workbook.ps1). Reopen the output to check the three-sheet structure, row/column counts, formula coverage, one returned value, and blank-indication retention.
7. **Report completion.** Give a clickable output path and the counts required by the reference contract.

## Input rules

- Default field names are `序号`, `药品`, and `获批适应症`; accept alternative names only when the user explicitly specifies them.
- Split only confirmed delimiters. The script defaults to English and full-width semicolons (`;` and `；`). Do not treat commas, slashes, parentheses, or line breaks as split delimiters unless the user authorizes this.
- A three-column helper sheet may be used for review, but it is not sufficient for formula backfill. Require a complete source sheet.
- If the identifier is duplicated, do not run XLOOKUP. Ask the user to provide a stable composite or technical key first.
- If an indication is blank, retain one row with a blank split-result cell. Do not delete or infer content.

## Output contract

The generated workbook contains exactly three sheets:

1. The complete source sheet, unchanged and used as the formula source.
2. `拆分结果`: identifier, entity name, and `适应症拆分结果` only.
3. `更新版拆分底稿`: identifier, entity name, `适应症拆分结果`, then every other source field in source order. The original full indication field remains in this sheet.

All fields after the first three columns are XLOOKUP formulas filled through every generated row. The source workbook is never modified.

## Run the script

On Windows with desktop Excel installed, run the script after confirmation. Use explicit paths and avoid overwriting files.

```powershell
& .\scripts\create_split_workbook.ps1 `
  -InputPath "C:\path\input.xlsx" `
  -OutputPath "C:\path\output.xlsx" `
  -SourceSheet "完整分子池" `
  -IdColumn "序号" `
  -NameColumn "药品" `
  -IndicationColumn "获批适应症"
```

If Excel COM automation is unavailable, report that the deterministic formula-backfill step is blocked; do not silently substitute a different output design.

## Hard stops

Stop and ask for human direction when the key is missing or duplicated, the full source sheet is unavailable, delimiters are ambiguous, the requested output would overwrite an existing file, source fields conflict, or formula calculation/verification fails.

## Completion report

Report: source record and field counts; output row and column counts; number of records split; blank-indication entities retained; number of formulas filled; verification status; any exception; and the output path.
