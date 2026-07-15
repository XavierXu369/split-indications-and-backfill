# Input and output contract

## Required source contract

Provide one Excel workbook and identify the complete source sheet.

The source sheet must contain:

| Role | Default field | Requirement |
|---|---|---|
| Stable key | `序号` | Nonblank and unique for every source record. |
| Entity name | `药品` | Nonblank display field copied to every split row. |
| Split field | `获批适应症` | Text to split by confirmed delimiters. |
| Full source fields | All remaining columns | Required for the updated split base. |

The user may provide another review sheet containing only the three key fields. Treat it as a cross-check only; do not use it as the formula source.

## Validation response

Before samples, report:

- source-sheet name, source record count, and field count;
- presence of the three required fields;
- blank or duplicate key count;
- blank entity-name count;
- blank-indication count and entity list;
- count of source records containing the confirmed delimiter;
- whether a complete source pool is available.

Pass only when the key is unique and the full source pool exists.

## Split rules

1. Default delimiter pattern: `[;；]`.
2. Trim whitespace around each split item.
3. Create one row for each nonblank split item, preserving the key and entity name.
4. If the original split field is blank, create one row with a blank `适应症拆分结果`.
5. Preserve original text in the source sheet and in the formula-backed original indication field.
6. Do not infer a disease, remove a risk/end-point term, standardize wording, or perform mapping.

## Three-sheet output

| Sheet | Required contents |
|---|---|
| Source | Original complete source sheet; unchanged. |
| `拆分结果` | Key, entity name, `适应症拆分结果`. |
| `更新版拆分底稿` | Key, entity name, `适应症拆分结果`, then all source fields except duplicated key and entity-name fields. |

In `更新版拆分底稿`, the original approved-indication column must be retained as a separate formula-backed field. The new split value appears only in column C.

## Formula contract

The output key is in column A. Each source field after column C receives an XLOOKUP formula pointing to the source key and the matching source return column. Example:

```excel
=IFERROR(XLOOKUP($A2,'完整分子池'!$A$2:$A$75,'完整分子池'!$C$2:$C$75,""),"")
```

- `$A2` locks the lookup column but lets the row change when filling downward.
- Both source ranges use absolute references.
- Fill all formulas from row 2 through the final split row.
- Preserve source number formats for returned columns.
- Keep the source and formula-backed sheets in the same workbook.

## Quality checks

Before delivery, verify:

- the source sheet is unchanged;
- every nonblank semicolon item becomes one row;
- blank indications remain one row;
- output columns = source columns + 1;
- formula count = output data rows × (source columns - 2);
- formula range is fully populated;
- one formula returns a value correctly and the original indication field remains available;
- the output file does not overwrite the input or an existing delivery.

## Completion report template

```text
已完成：[文件名及路径]

- 原始记录数：X；原始字段数：Y
- 拆分后记录数：Z；新增行数：Z - X
- 含分隔符并拆分的原始记录数：A
- 原始适应症为空但保留的记录数：B（列出实体标识）
- 更新版拆分底稿：Z 行、Y + 1 列
- 已填充回填公式：Z × (Y - 2) = N 个
- 已核验：三 Sheet、公式覆盖、返回值、原始适应症保留

异常/待确认事项：无；或列明实体标识与建议。
```
