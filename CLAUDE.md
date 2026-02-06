# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NPOIHelpers is a .NET Framework 4.7.2 library that provides extension methods for working with Excel files via NPOI. The library simplifies extracting cell ranges from Excel sheets into 2D arrays.

## Build Commands

```bash
# Build the solution
msbuild NPOIHelpers.sln

# Build in Release mode
msbuild NPOIHelpers.sln /p:Configuration=Release

# Run tests (requires VS Test Console or dotnet test with appropriate tooling)
vstest.console Tests\bin\Debug\Tests.dll
```

## Architecture

The library consists of three main components in the `NPOI.SCIta.Helpers` namespace:

- **ISheetExtension** (`NPOIHelpers/ISheetExtension.cs`): Extension methods for `ISheet`
  - `GetUsedRangeAddress()`: Calculates the bounding box of all non-empty cells
  - `GetRange(top, left, rows, cols)`: Extracts cell values into a 2D array with 1-based coordinates

- **CellRangeAddressExtension** (`NPOIHelpers/CellRangeAddressExtension.cs`): Extension methods for `CellRangeAddress`
  - `Intersect()`: Computes intersection of two cell ranges
  - `GetSize()`: Returns (rows, columns) tuple for a range

- **RangeResult** (`NPOIHelpers/RangeResult.cs`): Return type for `GetRange()`
  - `Values`: 2D object array containing cell values (strings, numbers, booleans, or null)
  - `Error`: Parallel 2D bool array indicating cells with errors

## Key Design Notes

- `GetRange()` uses **1-based coordinates** for top/left parameters (Excel-style), but internally NPOI uses 0-based indexing
- The library handles formula cells by reading their cached result values
- Empty cells and cells outside the sheet's used range are returned as null
- Tests use hardcoded file paths (`k:\users\stefano\`) and require specific test Excel files
