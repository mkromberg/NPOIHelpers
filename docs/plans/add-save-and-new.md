# Plan: Add Save and New Methods to Helpers

## Context

The APL layer already has `Save.aplf` and `New.aplf` that implement workbook saving and creation by calling library-specific APIs directly. Adding `Save` and `New` static methods to both helper libraries exposes this functionality at the C# level, making the helpers self-contained and consistent with the other workbook-level methods (`AllSheets`, `AddSheet`, `GetSheet`).

- **Save**: NPOI requires a `FileStream` + `IWorkbook.Write(stream)`; GemBox simply calls `ExcelFile.Save(path)`
- **New**: NPOI must select `XSSFWorkbook` (.xlsx) or `HSSFWorkbook` (.xls) based on the file extension; GemBox always returns `new ExcelFile()` regardless of extension

The APL `New.aplf` confirms this: it dispatches on the extension to choose XSSF vs HSSF for NPOI, and uses `ExcelFile` for both extensions in GemBox.

## Files to Modify

| File | Change |
|------|--------|
| `NPOIHelpers/Helpers.cs` | Add `New(string fileName)` and `Save(IWorkbook wb, string fileName)`; add `using System.IO;` |
| `GemBoxHelpers/Helpers.cs` | Add `New(string fileName)` and `Save(ExcelFile ef, string fileName)` |
| `NPOITests/UnitTest1.cs` | Add `TestNew()` and `TestSave()` test methods; add `using System.IO;` |
| `GemBoxTests/UnitTest1.cs` | Add `TestNew()` and `TestSave()` test methods; add `using System.IO;` |

## Implementation

### NPOIHelpers/Helpers.cs

Add `using System.IO;` at the top. Add both static methods inside `WorksheetExtension`:

```csharp
public static SS.UserModel.IWorkbook New(string fileName)
{
    var ext = Path.GetExtension(fileName).ToLowerInvariant();
    if (ext == ".xls")
        return new NPOI.HSSF.UserModel.HSSFWorkbook();
    return new NPOI.XSSF.UserModel.XSSFWorkbook();
}

public static void Save(SS.UserModel.IWorkbook wb, string fileName)
{
    using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
    {
        wb.Write(fs);
    }
}
```

### GemBoxHelpers/Helpers.cs

Add both static methods inside `WorksheetExtension`:

```csharp
public static ExcelFile New(string fileName)
{
    return new ExcelFile();
}

public static void Save(ExcelFile ef, string fileName)
{
    ef.Save(fileName);
}
```

### NPOITests/UnitTest1.cs

Add `using System.IO;` at the top, then add:

```csharp
[TestMethod]
public void TestNew()
{
    var path = Path.Combine(TestDataGenerator.TestDataPath, "new_test.xlsx");
    var wb = WorksheetExtension.New(path);
    Assert.IsInstanceOfType(wb, typeof(XSSFWorkbook));

    var pathXls = Path.Combine(TestDataGenerator.TestDataPath, "new_test.xls");
    var wbXls = WorksheetExtension.New(pathXls);
    Assert.IsInstanceOfType(wbXls, typeof(NPOI.HSSF.UserModel.HSSFWorkbook));
}

[TestMethod]
public void TestSave()
{
    var path = Path.Combine(TestDataGenerator.TestDataPath, "save_test.xlsx");
    var wb = WorksheetExtension.New(path);
    wb.CreateSheet("Test").PutRange(new object[,] { { "Hello", "World" } }, 1, 1);
    WorksheetExtension.Save(wb, path);

    Assert.IsTrue(File.Exists(path));

    var result = WorkbookFactory.Create(path).GetSheetAt(0).GetRange(1, 1, 1, 2);
    Assert.AreEqual("Hello", result.Values[0, 0]);
    Assert.AreEqual("World", result.Values[0, 1]);

    File.Delete(path);
}
```

### GemBoxTests/UnitTest1.cs

Add `using System.IO;` at the top, then add:

```csharp
[TestMethod]
public void TestNew()
{
    var path = Path.Combine(TestDataGenerator.TestDataPath, "new_test.xlsx");
    var ef = WorksheetExtension.New(path);
    Assert.IsNotNull(ef);
    Assert.IsInstanceOfType(ef, typeof(ExcelFile));
}

[TestMethod]
public void TestSave()
{
    var path = Path.Combine(TestDataGenerator.TestDataPath, "save_test.xlsx");
    var ef = WorksheetExtension.New(path);
    ef.Worksheets.Add("Test").PutRange(new object[,] { { "Hello", "World" } }, 1, 1);
    WorksheetExtension.Save(ef, path);

    Assert.IsTrue(File.Exists(path));

    var result = ExcelFile.Load(path).Worksheets[0].GetRange(1, 1, 1, 2);
    Assert.AreEqual("Hello", result.Values[0, 0]);
    Assert.AreEqual("World", result.Values[0, 1]);

    File.Delete(path);
}
```

## Notes

- Both `New` and `Save` follow the existing non-extension static pattern of `AllSheets`, `AddSheet`, and `GetSheet`
- NPOI's `New` defaults to XSSF for any unrecognised extension (e.g. `.xlsx`, `.xlsm`)
- `TestSave` uses `New` internally, so it also exercises the `New` path end-to-end
- The NPOI `UnitTest1.cs` does not currently have `using System.IO;` — it will need to be added; `TestDataGenerator.cs` already has it
- Test files are written to the existing `TestData` directory and cleaned up afterwards

## Verification

Build and run all tests:
```bash
msbuild NPOIHelpers.sln
vstest.console NPOITests\bin\Debug\NPOITests.dll
vstest.console GemBoxTests\bin\Debug\GemBoxTests.dll
```

Confirm `TestNew` and `TestSave` pass in both test suites.
