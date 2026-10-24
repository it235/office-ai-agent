---
name: excel-table-agent
description: Use for Excel tasks that need table understanding, cleanup, calculation, formula repair, transpose, pivot tables, charts, reporting, or multi-step spreadsheet automation.
application: Excel
default_for_application: true
tags: excel, spreadsheet, table, formula, chart, pivot, clean, transpose, statistics, data-analysis
allowed-tools: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData, SortData, FilterData, RemoveDuplicates, ConditionalFormat, MergeCells, AutoFit, FindReplace, CreatePivotTable, CreateSheet, DeleteSheet, RenameSheet, CopySheet, InsertRowCol, DeleteRowCol, HideRowCol, ProtectSheet, TransformData, DataAnalysis, GenerateReport
intent_types: data_analysis, formula, chart, table_format, data_clean, transform
---

# Excel Table Agent

Use this skill when the user asks Excel to do real spreadsheet work from natural language, especially:

- organize or beautify a table
- calculate totals, averages, rankings, rates, commissions, KPIs, or custom formulas
- repair an incorrect formula or algorithm
- generate a chart from selected data
- transpose rows and columns
- clean duplicates, blanks, spaces, inconsistent text, or obvious data quality issues
- create a pivot table, summary table, or report sheet

## Operating Rules

1. First read the current Excel context: workbook, sheet, selection address, used range, headers, sample rows, data types, numeric columns, text columns, formula cells, blanks, duplicates, and candidate target area.
2. Prefer native JSON tools. This skill cannot execute VBA: if no registered tool can express the task, report a capability gap instead of planning around VBA.
3. If the user selected a range, treat it as the primary working range. If there is no selection, infer the current table from `UsedRange`.
4. Do not ask the user for data that the add-in can observe from Excel.
5. For write operations, produce a plan with target ranges and expected effects before acting when the task is medium or risky.
6. After each operation, observe the sheet state: changed range, formulas, chart count, row/column shape, or generated summary.
7. If execution fails, repair the tool parameters using the observation. Do not fall back to a long chat explanation unless execution is impossible.

## Tool Preferences

- Table formatting: `FormatRange`, `AutoFit`, `ConditionalFormat`
- Formula generation or repair: `ApplyFormula`
- Summary/statistics: `DataAnalysis`, `ApplyFormula`, `CreatePivotTable`
- Chart generation: `CreateChart` after selecting chart type from data shape
- Row/column transpose: `TransformData` with `operation=transpose`
- Cleaning: `CleanData`, `RemoveDuplicates`, `FindReplace`
- Report output: `GenerateReport`, `WriteData`, `CreateChart`
- Worksheet structure: `CreateSheet`, `DeleteSheet`, `RenameSheet`, `CopySheet`
- Row/column structure: `InsertRowCol`, `DeleteRowCol`, `HideRowCol`
- Sheet protection: `ProtectSheet`

## Chart Selection Heuristics

- Trend over time: line chart
- Category comparison: column or bar chart
- Share of total with few categories: pie chart
- Relationship between two numeric columns: scatter chart
- Multiple numeric series by category: clustered column chart

## Formula Repair Heuristics

When asked to fix an algorithm or formula:

1. Inspect the source columns and the current formula if present.
2. Identify references, relative/absolute addressing, row offsets, and likely fill-down range.
3. Generate the corrected formula for the first data row.
4. Apply it to the target range with fill-down when appropriate.
5. Observe whether Excel reports formula errors such as `#VALUE!`, `#REF!`, `#N/A`, or empty results.
