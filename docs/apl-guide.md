# NPOIHelpers — Dyalog APL User Guide

This guide covers using the `ArrayEWE.Helpers` library from Dyalog APL via the .NET bridge.

## Setup

Load the DLL once per session before using any functions:

```apl
⎕USING←'ArrayEWE.Helpers,path\to\NPOIHelpers.dll'
WorksheetExtension.Initialize
```

`Initialize` must be called at least once — it registers the encoding provider needed to read `.xls` files.

Assign `WorksheetExtension` to a shorter name for convenience:

```apl
H←WorksheetExtension
```

---

## Opening and Creating Workbooks

```apl
wb ← H.Open 'C:\data\report.xlsx' ''          ⍝ open existing file (.xls or .xlsx)
wb ← H.Open 'C:\data\encrypted.xlsx' 'secret' ⍝ open password-protected file
wb ← H.New  'C:\data\output.xlsx'             ⍝ create new (format inferred from extension)
```

`New` does not write a file — it creates an in-memory workbook. Call `Save` to persist it.

`Open` returns an `IWorkbook`. Because Dyalog APL does not apply C# default parameter values, the password argument must always be supplied — pass `''` when there is no password.

Encrypted `.xlsx` files are fully supported. Encrypted `.xls` files may fail to open depending on the encryption variant used by Excel (NPOI 2.x does not support CryptoAPI RC4, which Excel 2007+ uses for `.xls` passwords).

---

## Working with Sheets

```apl
⍝ list all sheet names in a workbook
names ← H.AllSheets wb              ⍝ returns a .NET string array

⍝ get a sheet by name
sheet ← H.GetSheet wb 'Sheet1'

⍝ add a new sheet (appended at the end)
sheet ← H.AddSheet wb 'NewSheet' 0

⍝ add a sheet as the first tab
sheet ← H.AddSheet wb 'Cover' 1

⍝ delete a sheet (throws if name not found)
H.DeleteSheet wb 'OldSheet'

⍝ get sheet by position (0-based, direct NPOI call)
sheet ← wb.GetSheetAt 0
```

---

## Reading Cell Ranges

`GetRange` returns a `RangeResult` object. Coordinates are **1-based** (row 1, column 1 = cell A1). Pass a vector of property names to control which fields are populated; omitted properties are absent from the result.

```apl
⍝ entire used range — all properties
r ← H.GetUsedRange sheet

⍝ specific properties, specific rectangle (top left rows cols — 1-based)
r ← H.GetRange sheet ('Values' 'NumberFormat') 3 2 10 5

⍝ all five properties, specific rectangle
r ← H.GetRange sheet ('Values' 'Error' 'NumberFormat' 'HasFormula' 'Formula') 3 2 10 5
```

Available property names (case-insensitive): `Values`, `Error`, `NumberFormat`, `HasFormula`, `Formula`.

If the requested range does not overlap the sheet's data, the result contains no properties.

---

## The RangeResult Object

`RangeResult` contains only the properties that were requested. It has two parallel members:

| Member | Type | Contents |
|---|---|---|
| `Names` | `string[]` | Names of the properties present in this result |
| `Values` | `object[]` | Parallel array of 2D data arrays, one per name |

Access a specific property by name using the indexer:

```apl
vals ← r['Values']        ⍝ returns the Values 2D array (object[,])
fmts ← r['NumberFormat']  ⍝ returns the NumberFormat 2D array (string[,])
```

Returns `⎕NULL` if the named property was not requested.

The five available property names and their 2D array types:

| Name | Array type | Contents |
|---|---|---|
| `Values` | `object[,]` | Cell values: `Double`, `String`, `Boolean`, or `null` |
| `Error` | `bool[,]` | `1` where the cell contains an Excel error (e.g. `#DIV/0!`) |
| `NumberFormat` | `string[,]` | Excel format string, or `null` for General/no format |
| `HasFormula` | `bool[,]` | `1` where the cell is a formula |
| `Formula` | `string[,]` | Formula text (e.g. `"A1+B1"`) for formula cells |

Formula cells: `Values` contains the **cached result**; `HasFormula` and `Formula` give the expression.

### Converting to APL Arrays

The 2D array returned by the indexer is a .NET array. Convert to a native APL matrix with `↑`:

```apl
vals ← ↑r['Values']       ⍝ APL matrix of boxed .NET objects
errs ← ↑r['Error']        ⍝ Boolean matrix — 1 = cell is an error
fmts ← ↑r['NumberFormat'] ⍝ format strings or ⎕NULL
```

Null cells in `Values` appear as `⎕NULL`. To replace with a fill value:

```apl
vals ← 0@(⎕NULL∘≡¨)↑r['Values']
```

---

## Writing Cell Ranges

`PutRange` writes one property at a time. Both the `Values` matrix and the `NumberFormat` matrix are passed as plain APL matrices — the .NET bridge converts them to `Object[,]` automatically.

```apl
data ← 2 3⍴'Name' 'Score' 'Pass' 'Alice' 95.5 1
fmts ← 2 3⍴'' '' '' '' '#,##0.00' ''

⍝ write values starting at row 1, column 1
H.PutRange sheet 'Values' data 1 1

⍝ apply number formats at the same position
H.PutRange sheet 'NumberFormat' fmts 1 1

⍝ write values starting at row 5, column 3
H.PutRange sheet 'Values' data 5 3
```

Supported property names (case-insensitive): `Values`, `NumberFormat`. Each call writes exactly one property; call twice to write both.

For `Values`: supported cell types are `String`, `Double` (and other numeric types), `Boolean`, `DateTime`, `⎕NULL` (leaves cell blank).

For `NumberFormat`: each cell should contain an Excel format string (e.g. `'#,##0.00'`) or an empty string to leave that cell's format unchanged.

### Appending Data

`AppendRange` writes below the last row of existing data, with one blank row separating the blocks:

```apl
H.AppendRange sheet data 1       ⍝ append starting at column 1
H.AppendRange sheet data 2       ⍝ append starting at column 2
```

---

## Housekeeping

```apl
⍝ remove all rows from a sheet
H.ClearSheet sheet

⍝ remove rows with ZeroHeight=1 and blank hidden-column cells
H.ClearHidden sheet

⍝ auto-size all columns
H.AutoFitColumns sheet 1 2147483647   ⍝ 2147483647 = int.MaxValue = all columns

⍝ auto-size columns 2–5
H.AutoFitColumns sheet 2 4            ⍝ left=2, cols=4

⍝ save workbook (password argument is required — pass '' for no encryption)
H.Save wb 'C:\data\output.xlsx' ''
H.Save wb 'C:\data\secret.xlsx' 'password'   ⍝ save with file-level encryption

⍝ protect / unprotect workbook structure (prevents sheet add/delete/rename)
H.Protect wb 'password'   ⍝ protect
H.Protect wb ''           ⍝ unprotect

⍝ protect / unprotect an individual sheet (prevents cell edits, row/col changes, etc.)
H.ProtectSheet sheet 'password'   ⍝ protect
H.ProtectSheet sheet ''           ⍝ unprotect

⍝ read column widths (in Excel character-width units, same as Format > Column Width)
widths ← H.GetColumnWidths sheet 1 2147483647   ⍝ all columns in used range
widths ← H.GetColumnWidths sheet 2 4            ⍝ columns 2–5 only
```

---

## Typical Read Workflow

```apl
H ← WorksheetExtension
WorksheetExtension.Initialize

wb    ← H.Open 'C:\data\report.xlsx' ''
sheet ← H.GetSheet wb 'Data'
r     ← H.GetUsedRange sheet

vals  ← ↑r['Values']          ⍝ APL matrix (may contain ⎕NULL for empty cells)
errs  ← ↑r['Error']           ⍝ Boolean matrix of Excel errors
```

## Typical Write Workflow

```apl
wb    ← H.New 'C:\data\output.xlsx'
sheet ← H.AddSheet wb 'Results'

data ← 3 2⍴'Name' 'Score' 'Alice' 95.5 'Bob' 82.0
fmts ← 3 2⍴'' '' '' '#,##0.00' '' '#,##0.00'

H.PutRange sheet 'Values' data 1 1
H.PutRange sheet 'NumberFormat' fmts 1 1
H.Save wb 'C:\data\output.xlsx' ''
```

---

## Notes

- Workbooks opened with `Open` should be disposed after use: `wb.Dispose`. (The `HelpedBook` APL class, used by the higher-level workspace wrappers, disposes automatically on destruction.)
- `.xls` (Excel 97–2003) and `.xlsx` (Excel 2007+) are both supported for reading and writing.
- `GetRange` and `GetUsedRange` never throw for out-of-range coordinates — they return an empty result instead.
- Number format strings follow Excel's format syntax (e.g. `"dd/mm/yyyy"`, `"#,##0.00"`, `"0.00%"`).
- Dyalog APL does not resolve extension methods on instance objects: all `ISheet` methods (`GetRange`, `PutRange`, `GetUsedRange`, `ClearSheet`, etc.) must be called via `H.MethodName sheet ...`.
- Dyalog APL does not apply C# default parameter values: all parameters must be supplied explicitly. Use `2147483647` (`int.MaxValue`) where the C# default means "unbounded". Pass `''` where the C# default is `null` (e.g. the `password` arguments to `Open`, `Save`, and `Protect`).
- `Protect` controls workbook *structure* protection (prevents adding, deleting, or renaming sheets) — it is separate from file-level encryption set via the `password` argument to `Save`.
- `GetColumnWidths` returns a `Double[]` of column widths in points, matching Excel's `Column.Width` COM property. For unset columns it returns the sheet's default column width. On an empty sheet (no rows) it returns an empty array.
- `ProtectSheet` controls sheet-level protection (prevents cell edits, row/column insertion, formatting changes, etc.) — it is independent of both workbook structure protection and file encryption. Pass `''` to remove sheet protection.
