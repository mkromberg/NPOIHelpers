# Plan: Rename Tests → NPOITests and Add GemBox Projects

## Context

The project currently has NPOI-based helpers and tests. We want to evaluate GemBox.Spreadsheet
as an alternative to NPOI. This plan renames the existing test folder for clarity, then
creates parallel GemBoxHelpers and GemBoxTests projects with the same structure and API
shape as their NPOI counterparts, but using GemBox types internally.

---

## Step 1: Rename Tests → NPOITests

Use `git mv` to preserve history.

```
git mv Tests NPOITests
git mv NPOITests\Tests.csproj NPOITests\NPOITests.csproj
```

Update `NPOIHelpers.sln`:
- Change project entry from `"Tests", "Tests\Tests.csproj"` → `"NPOITests", "NPOITests\NPOITests.csproj"`

Update `NPOITests\NPOITests.csproj`:
- Change `<AssemblyName>Tests</AssemblyName>` → `<AssemblyName>NPOITests</AssemblyName>`
- Change `<RootNamespace>Tests</RootNamespace>` → `<RootNamespace>NPOITests</RootNamespace>`

No source file changes are needed — the `namespace Tests` declarations in .cs files can stay as-is
since they don't need to match the project name.

---

## Step 2: Create GemBoxHelpers project

New folder: `GemBoxHelpers\`

### GemBoxHelpers\GemBoxHelpers.csproj
Copy NPOIHelpers.csproj with these changes:
- Assembly name: `GemBox.SCIta.Helpers`
- Root namespace: `SCIta.GemBoxHelpers`
- New GUID
- Replace `<PackageReference Include="NPOI" Version="2.7.5" />` with `<PackageReference Include="GemBox.Spreadsheet" Version="*" />`
- Remove `System.Text.Encoding.CodePages` (not needed by GemBox)
- Keep `System.Configuration.ConfigurationManager` only if needed

### GemBoxHelpers\RangeResult.cs
Identical to NPOIHelpers\RangeResult.cs — no changes needed.

### GemBoxHelpers\WorksheetExtension.cs
Replaces `ISheetExtension.cs`. Key API mapping:

| NPOI | GemBox |
|---|---|
| `ISheet` | `ExcelWorksheet` |
| `CellRangeAddress` | `CellRange` |
| `sheet.GetRow(r)?.GetCell(c)` | `worksheet.Cells[r, c]` |
| `cell.CellType` switch | `cell.ValueType` switch |
| `cell.NumericCellValue` | `(double)cell.Value` |
| `cell.StringCellValue` | `(string)cell.Value` |
| `cell.BooleanCellValue` | `(bool)cell.Value` |
| `cell.CachedFormulaResultType` | not needed — GemBox evaluates formulas automatically |
| `WorkbookFactory.Create(path)` | `ExcelFile.Load(path)` |

Methods to implement:
- `GetUsedCellRange(ExcelWorksheet sheet)` — replaces `GetUsedRangeAddress()`
  - Use `sheet.GetUsedCellRange()` (GemBox built-in) then translate to the 0-based row/col bounds
- `GetRange(ExcelWorksheet, top, left, rows, cols)` — same 1-based coordinate convention
- `InitEncodings()` — replace encoding setup with `SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")`
- `PutRange(ExcelWorksheet, object[,], int[])` — use `worksheet.Cells[r, c].Value = v`

### GemBoxHelpers\CellRangeExtension.cs
Replaces `CellRangeAddressExtension.cs`. Key mapping:

| NPOI | GemBox |
|---|---|
| `CellRangeAddress` | `CellRange` |
| `.FirstRow` / `.LastRow` | `.FirstRowIndex` / `.LastRowIndex` |
| `.FirstColumn` / `.LastColumn` | `.FirstColumnIndex` / `.LastColumnIndex` |

Methods:
- `Intersect(CellRange x, CellRange y)` — same logic, adapted field names
- `GetSize(CellRange x)` — same logic, adapted field names

---

## Step 3: Create GemBoxTests project

New folder: `GemBoxTests\`

### GemBoxTests\GemBoxTests.csproj
Copy NPOITests\NPOITests.csproj with these changes:
- Assembly name: `GemBoxTests`
- Root namespace: `GemBoxTests`
- New GUID
- Replace NPOI PackageReference with `GemBox.Spreadsheet`
- Remove MSTest reference to NPOITests; project reference points to `GemBoxHelpers\GemBoxHelpers.csproj`

### GemBoxTests\TestDataGenerator.cs
Replace NPOI workbook API with GemBox:

| NPOI | GemBox |
|---|---|
| `new XSSFWorkbook()` | `new ExcelFile()` |
| `wb.CreateSheet("Sheet1")` | `ef.Worksheets.Add("Sheet1")` |
| `sheet.CreateRow(r).CreateCell(c).SetCellValue(v)` | `ws.Cells[r, c].Value = v` |
| `wb.Write(fileStream)` | `ef.Save(path)` |

### GemBoxTests\UnitTest1.cs
Replace NPOI types in assertions:

| NPOI | GemBox |
|---|---|
| `CellRangeAddress` | `CellRange` |
| `.FirstRow` / `.LastRow` | `.FirstRowIndex` / `.LastRowIndex` |
| `ISheetExtension.GetUsedRangeAddress(sheet)` | `WorksheetExtension.GetUsedCellRange(sheet)` |
| `ISheetExtension.GetRange(...)` | `WorksheetExtension.GetRange(...)` |
| Open file: `WorkbookFactory.Create(path)` | `ExcelFile.Load(path)` |

The 6 test methods and their assertions stay structurally identical.

### GemBoxTests\App.config
GemBox has fewer dependency conflicts than NPOI, so the binding redirects will be simpler.
Start with an empty `<assemblyBinding>` block and add redirects only if build errors require them.

---

## Step 4: Add new projects to solution

Add to `NPOIHelpers.sln`:
```
Project("{FAE04EC0-301F-11D3-BF4B-00C0F79EFBC4}") = "GemBoxHelpers", "GemBoxHelpers\GemBoxHelpers.csproj", "{<new-guid>}"
EndProject
Project("{FAE04EC0-301F-11D3-BF4B-00C0F79EFBC4}") = "GemBoxTests", "GemBoxTests\GemBoxTests.csproj", "{<new-guid>}"
EndProject
```
Add corresponding build configuration entries for Debug|AnyCPU and Release|AnyCPU.

---

## Files to create / modify

| File | Action |
|---|---|
| `NPOIHelpers.sln` | Update Tests path; add GemBoxHelpers and GemBoxTests |
| `Tests\` → `NPOITests\` | `git mv` |
| `NPOITests\Tests.csproj` → `NPOITests\NPOITests.csproj` | `git mv` + edit |
| `GemBoxHelpers\GemBoxHelpers.csproj` | Create |
| `GemBoxHelpers\RangeResult.cs` | Copy unchanged |
| `GemBoxHelpers\WorksheetExtension.cs` | Create (refactored ISheetExtension) |
| `GemBoxHelpers\CellRangeExtension.cs` | Create (refactored CellRangeAddressExtension) |
| `GemBoxHelpers\Properties\AssemblyInfo.cs` | Create |
| `GemBoxTests\GemBoxTests.csproj` | Create |
| `GemBoxTests\UnitTest1.cs` | Create (refactored) |
| `GemBoxTests\TestDataGenerator.cs` | Create (refactored) |
| `GemBoxTests\App.config` | Create |
| `GemBoxTests\Properties\AssemblyInfo.cs` | Create |

---

## Verification

1. `msbuild NPOIHelpers.sln` — all 4 projects build without errors
2. `vstest.console NPOITests\bin\Debug\NPOITests.dll` — existing tests still pass
3. `vstest.console GemBoxTests\bin\Debug\GemBoxTests.dll` — GemBox tests pass
