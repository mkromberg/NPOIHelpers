# Build and Test Instructions

## Prerequisites

- Visual Studio 2022 (or later) with .NET Framework 4.7.2 targeting pack
- MSBuild (included with Visual Studio)

## Building the Project

Use Visual Studio's MSBuild to build the solution:

```bash
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" NPOIHelpers.sln -verbosity:minimal -restore
```

For a release build:

```bash
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" NPOIHelpers.sln -verbosity:minimal -restore -p:Configuration=Release
```

Note: The `dotnet build` command does not work reliably with this .NET Framework project due to GitVersionTask compatibility issues.

## Running Tests

After building, run tests using vstest.console:

```bash
"C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" Tests\bin\Debug\Tests.dll
```

### Test File Dependencies

Some tests require Excel files at specific paths:
- `k:\users\stefano\bigtest.xlsx`
- `k:\users\stefano\bigtest1.xlsx`

Tests that don't require external files (e.g., `TestCellRange`, `TestIntersect`) will pass without these files.
