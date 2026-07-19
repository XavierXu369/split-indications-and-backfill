# Input and output contract

## 1. Required input

Provide one Excel workbook and identify the complete source sheet.

| Role | Default field | Requirement |
|---|---|---|
| Stable key | `序号` | Nonblank and unique; underlying Excel value type must be retained. |
| Entity name | `药品` | Nonblank display field copied to every split row. |
| Split field | `获批适应症` | Text split only by approved delimiters. |
| Full source fields | All remaining columns | Required for the formula-backed updated base. |

A helper sheet containing only the three key fields may be used for review, but never as the formula source.

## 2. Staged execution

### Inspect

Run read-only and report:

- workbook path, SHA-256, size, and last-write time;
- source-sheet name and source signature;
- UsedRange surface dimensions and effective nonblank dimensions;
- source record and field counts;
- workbook sheet list, extra sheets, hidden sheets, and defined-name count;
- required-field presence;
- blank and duplicate header counts;
- blank and duplicate key counts;
- blank entity-name count;
- blank-indication count and entity list;
- records containing approved delimiters and records that actually create multiple rows;
- consecutive/trailing delimiters or other empty segments;
- repeated indication items within one source record;
- delimiter-only records;
- merged cells in the effective source range;
- formulas in the source sheet that depend directly on another workbook sheet.

Inspection must not create a workbook. A run passes only when `CanProceed` is true.

### Preview

Preview must reuse the same validation and requires explicit confirmation that the named source sheet is the complete pool. Return 1–2 representative examples, prioritizing a multi-item row and then a blank-indication row. Also return one exact XLOOKUP formula example. Do not create a workbook.

### Generate

Generate only after explicit approval and complete-source confirmation. Pass `ExpectedInputSha256` and `ExpectedSourceSignature` from the approved inspection/preview so a changed source cannot be used silently.

## 3. Hard stops and warnings

### Hard stops

- input or source sheet unavailable;
- incomplete source pool;
- required field missing or duplicated;
- blank or duplicate header;
- blank or duplicate stable key;
- blank entity name;
- merged cells in the effective source range;
- a nonblank indication contains only delimiters/whitespace;
- delimiter is ambiguous or unapproved;
- source formulas directly depend on a sheet that the standard three-sheet output would delete;
- input/output extensions differ;
- output would overwrite the input or an existing file;
- approved file or source signature has changed;
- formula calculation, formula coverage, source preservation, or reopened-workbook verification fails.

### Warnings that do not change data automatically

- blank indication: retain one blank split row;
- consecutive or trailing delimiters: ignore blank segments but report the affected stable keys;
- repeated indication items within one source record: retain every occurrence and report it;
- hidden or extra sheets: report them before the standard output deletes non-source sheets;
- workbook-defined names or formulas: report them for awareness.

## 4. Split rules

1. Default delimiter pattern: `[;；]`.
2. Trim whitespace around each item.
3. Create one row for every nonblank item, retaining original order.
4. Do not deduplicate repeated items.
5. If the original indication is blank, create one row with a blank `适应症拆分结果`.
6. Preserve the full original indication in the source sheet and as a formula-backed field in `更新版拆分底稿`.
7. Do not infer diseases, remove risk/end-point terms, standardize wording, or perform TA Mapping.

## 5. Three-sheet output

| Sheet | Required contents |
|---|---|
| Source | Original complete source sheet, unchanged in values, formulas, and number formats. |
| `拆分结果` | Stable key, entity name, `适应症拆分结果`. |
| `更新版拆分底稿` | Stable key, entity name, `适应症拆分结果`, then all source fields except duplicated key and entity-name fields. |

The output contains exactly these three sheets in this order. Output columns equal source columns plus one.

## 6. Key and formula contract

- Read stable keys through Excel `Value2` and write the same underlying typed value; never use display `.Text` as the lookup key.
- Preserve key number formats in both output sheets.
- Distinguish text keys from numeric keys during uniqueness checks.
- Use the output key in column A for every XLOOKUP.
- Use absolute source ranges and quote/escape the source-sheet name.
- Preserve true blanks rather than returning zero for blank source cells.
- Fill formulas from row 2 through the final split row.
- Preserve the matched source row's number format for every returned field.

Example:

```excel
=IFERROR(LET(_v,XLOOKUP($A2,'完整分子池'!$A$2:$A$75,'完整分子池'!$C$2:$C$75,""),IF(_v="","",_v)),"")
```

## 7. Required verification

After saving, close and reopen the output read-only. Verify:

- exactly three required sheets and no others;
- source signature unchanged;
- split row count equals the precomputed expectation;
- one output row for every nonblank item and every blank source indication;
- output columns = source columns + 1;
- exact formula count = split data rows × (source columns - 2);
- every cell in the formula region contains a formula;
- no formula error cells;
- at least one nonblank returned value equals its source value;
- the original indication field remains present;
- input remains untouched and output path is new.

## 8. Completion report

Report:

- source file SHA-256 and source signature;
- source records/fields and split records/fields;
- added rows;
- records containing delimiters and records actually split;
- blank indications retained;
- empty-segment and repeated-item warnings;
- exact formula count;
- three-sheet, source-preservation, return-value, and formula-error status;
- exceptions and output path.
