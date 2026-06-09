# Plan: June Feature Additions to NPOIHelpers

All changes go in `NPOIHelpers/Helpers.cs`, `NPOIHelpers/RangeResult.cs`, `NPOITests/UnitTest1.cs`, and APL source files under `APLSource/`.

---

## What was implemented

### Group A — Sheet-level manipulation

**`ClearSheet(this ISheet sheet)`**
Removes all rows from the sheet.

**`ClearHidden(this ISheet sheet)`**
Removes hidden rows (`row.ZeroHeight == true`) and blanks cells in hidden columns.

**`AutoFitColumns(this ISheet sheet, int left = 1, int cols = int.MaxValue)`**
Calls `sheet.AutoSizeColumn(i)` for each column in range.

---

### Group B — Write

**`AppendRange(this ISheet sheet, object[,] data, int left = 1)`**
Finds the last used row via `GetUsedRangeAddress()` and delegates to `PutRange`.

**`PutRange(this ISheet sheet, string name, object[,] value, int top, int left)`**
Single public overload — one property name and one matrix per call. Both the `Values` matrix and the `NumberFormat` matrix are `object[,]`; the bridge converts APL matrices automatically. Supported names:
- `"Values"` — writes cell values (string, double, bool, DateTime, null)
- `"NumberFormat"` — applies Excel format strings; empty or null cells are skipped

To write both values and formats, call twice:
```apl
H.PutRange sheet 'Values' data 1 1
H.PutRange sheet 'NumberFormat' fmts 1 1
```

Note: all five parameters are required — Dyalog APL does not apply C# default values.

---

### Group C — Read

**`RangeResult` — refactored to dynamic**
Instead of fixed properties, `RangeResult` holds two parallel arrays and a string indexer:
- `string[] Names` — names of the properties present in this result
- `object[] Values` — parallel array of 2D data arrays
- `this[string name]` — returns the 2D array for the named property, or `null` if absent

**`GetRange(this ISheet sheet, string[] names, int top, int left, int rows, int cols)`**
Named-property overload — allocates and fills only the requested arrays. Unknown names throw. Order of `Names` in the result matches the order requested.

**`GetRange(this ISheet sheet, int top, int left, int rows, int cols)`**
Positional overload — returns all five properties.

Five available property names: `Values` (`object[,]`), `Error` (`bool[,]`), `NumberFormat` (`string[,]`), `HasFormula` (`bool[,]`), `Formula` (`string[,]`).

**`GetUsedRange(this ISheet sheet)`**
Convenience wrapper: calls `GetUsedRangeAddress()` then the positional `GetRange`.

---

### Group D — Sheet management

**`AddSheet(IWorkbook workbook, string name, bool first = false)`**
Calls `SetSheetOrder(name, 0)` when `first == true`.

**`DeleteSheet(IWorkbook workbook, string name)`**
Throws `ArgumentException` if sheet not found.

---

### APL layer

**`docs/apl-guide.md`** — new user guide for Dyalog APL callers. Key notes:
- Dyalog does not resolve extension methods on instance objects — use `H.MethodName sheet ...`
- Dyalog does not apply C# default parameter values — all parameters must be supplied explicitly
- Use `2147483647` (`int.MaxValue`) where the C# default means "unbounded"

**`APLSource/PutRange.aplf`** — updated wrapper. Dispatches on argument shape:
- `PutRange sheet data` → Values only, top=1 left=1
- `PutRange sheet data topleft` → Values only with coords
- `PutRange sheet values names` → named-property write, loops once per name

**`APLSource/GetRange.aplf`** — new wrapper:
- `GetRange sheet` → returns raw values matrix (via `GetUsedRange`)
- `GetRange sheet names` → returns `RangeResult` for selected properties

**`APLSource/TestHelpers.aplf`** — APL-level integration test for the new API.

---

## Tests

42 tests in `UnitTest1.cs`, all passing. Local helpers `PutValues` and `PutWithFormats` wrap the single-overload `PutRange` for test convenience.
