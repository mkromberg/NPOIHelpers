# PutRange Implementation Plan

## Overview

Add a `PutRange` extension method to `ISheet` that writes a 2D array of values to a specified location in the sheet. This is the inverse of `GetRange`.

## Method Signature

```csharp
public static void PutRange(this ISheet sheet, object[,] values, int top = 1, int left = 1)
```

**Parameters:**
- `values`: 2D array of values to write (supports `string`, `double`, `bool`, `DateTime`, and `null`)
- `top`: 1-based row number for the top-left corner (default: 1)
- `left`: 1-based column number for the top-left corner (default: 1)

## Behavior

### Value Type Handling

| Input Type | Cell Action |
|------------|-------------|
| `string` | `cell.SetCellValue(string)` |
| `double`, `int`, `float`, etc. | `cell.SetCellValue(double)` |
| `bool` | `cell.SetCellValue(bool)` |
| `DateTime` | `cell.SetCellValue(DateTime)` |
| `null` | `cell.SetBlank()` |

### Row/Cell Creation

- Create rows if they don't exist (`sheet.CreateRow` or `sheet.GetRow`)
- Create cells if they don't exist (`row.CreateCell` or `row.GetCell`)

### Coordinate System

- Uses **1-based coordinates** to match `GetRange` and Excel conventions
- `top=1, left=1` corresponds to cell A1 (0-indexed row 0, column 0)

## Implementation

Add to `ISheetExtension.cs`:

```csharp
public static void PutRange(this SS.UserModel.ISheet sheet, object[,] values, int top = 1, int left = 1)
{
  int rows = values.GetLength(0);
  int cols = values.GetLength(1);

  for (int r = 0; r < rows; r++)
  {
    int rowIndex = (top - 1) + r;
    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

    for (int c = 0; c < cols; c++)
    {
      int colIndex = (left - 1) + c;
      var cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
      var value = values[r, c];

      switch (value)
      {
        case null:
          cell.SetBlank();
          break;
        case string s:
          cell.SetCellValue(s);
          break;
        case bool b:
          cell.SetCellValue(b);
          break;
        case DateTime dt:
          cell.SetCellValue(dt);
          break;
        case double d:
          cell.SetCellValue(d);
          break;
        default:
          // Convert other numeric types to double
          cell.SetCellValue(Convert.ToDouble(value));
          break;
      }
    }
  }
}
```

## Test Plan

### Test File Strategy

Tests will create in-memory workbooks, write data with `PutRange`, then verify with `GetRange`. This avoids file I/O and keeps tests fast.

### Test Cases

#### 1. Basic Round-Trip Test
Write values and read them back to verify they match.

```csharp
[TestMethod]
public void TestPutRangeRoundTrip()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  var input = new object[,] {
    { "A", "B", "C" },
    { 1.0, 2.0, 3.0 },
    { true, false, null }
  };

  sheet.PutRange(input, 1, 1);
  var result = sheet.GetRange(1, 1, 3, 3);

  Assert.AreEqual("A", result.Values[0, 0]);
  Assert.AreEqual("B", result.Values[0, 1]);
  Assert.AreEqual(1.0, result.Values[1, 0]);
  Assert.AreEqual(true, result.Values[2, 0]);
  Assert.AreEqual(null, result.Values[2, 2]);
}
```

#### 2. Offset Position Test
Verify writing to non-origin positions works correctly.

```csharp
[TestMethod]
public void TestPutRangeOffset()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  var input = new object[,] { { "X" } };
  sheet.PutRange(input, 5, 3);  // Row 5, Column C

  var result = sheet.GetRange(5, 3, 1, 1);
  Assert.AreEqual("X", result.Values[0, 0]);

  // Verify cells before are empty
  var before = sheet.GetRange(1, 1, 4, 2);
  Assert.AreEqual(null, before.Values[0, 0]);
}
```

#### 3. Overwrite Existing Data Test
Verify that PutRange overwrites existing cell values.

```csharp
[TestMethod]
public void TestPutRangeOverwrite()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  // Write initial data
  sheet.PutRange(new object[,] { { "Old" } }, 1, 1);

  // Overwrite
  sheet.PutRange(new object[,] { { "New" } }, 1, 1);

  var result = sheet.GetRange(1, 1, 1, 1);
  Assert.AreEqual("New", result.Values[0, 0]);
}
```

#### 4. Data Types Test
Verify all supported data types are handled correctly.

```csharp
[TestMethod]
public void TestPutRangeDataTypes()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  var input = new object[,] {
    { "text", 42.5, true, null },
    { 100, -3.14, false, "" }
  };

  sheet.PutRange(input, 1, 1);
  var result = sheet.GetRange(1, 1, 2, 4);

  Assert.AreEqual("text", result.Values[0, 0]);
  Assert.AreEqual(42.5, result.Values[0, 1]);
  Assert.AreEqual(true, result.Values[0, 2]);
  Assert.AreEqual(null, result.Values[0, 3]);
  Assert.AreEqual(100.0, result.Values[1, 0]);  // int -> double
  Assert.AreEqual(-3.14, result.Values[1, 1]);
  Assert.AreEqual(false, result.Values[1, 2]);
  Assert.AreEqual("", result.Values[1, 3]);
}
```

#### 5. Large Range Test
Verify performance and correctness with larger data sets.

```csharp
[TestMethod]
public void TestPutRangeLarge()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  int rows = 100, cols = 20;
  var input = new object[rows, cols];
  for (int r = 0; r < rows; r++)
    for (int c = 0; c < cols; c++)
      input[r, c] = $"R{r}C{c}";

  sheet.PutRange(input, 1, 1);
  var result = sheet.GetRange(1, 1, rows, cols);

  Assert.AreEqual(rows, result.Values.GetLength(0));
  Assert.AreEqual(cols, result.Values.GetLength(1));
  Assert.AreEqual("R0C0", result.Values[0, 0]);
  Assert.AreEqual("R99C19", result.Values[99, 19]);
}
```

#### 6. Empty Array Test
Verify behavior with empty or zero-dimension arrays.

```csharp
[TestMethod]
public void TestPutRangeEmpty()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  var input = new object[0, 0];
  sheet.PutRange(input, 1, 1);  // Should not throw

  // Sheet should remain empty
  Assert.AreEqual(-1, sheet.FirstRowNum);
}
```

#### 7. Integer Types Test
Verify various integer types are converted to double.

```csharp
[TestMethod]
public void TestPutRangeIntegerTypes()
{
  var wb = new XSSFWorkbook();
  var sheet = wb.CreateSheet("Test");

  var input = new object[,] {
    { (int)1, (long)2, (short)3, (byte)4, (float)5.5 }
  };

  sheet.PutRange(input, 1, 1);
  var result = sheet.GetRange(1, 1, 1, 5);

  Assert.AreEqual(1.0, result.Values[0, 0]);
  Assert.AreEqual(2.0, result.Values[0, 1]);
  Assert.AreEqual(3.0, result.Values[0, 2]);
  Assert.AreEqual(4.0, result.Values[0, 3]);
  Assert.AreEqual(5.5, (double)result.Values[0, 4], 0.01);
}
```

### Test Organization

Add tests to a new file `Tests/PutRangeTests.cs` to keep them separate from existing tests. Add the file to `Tests.csproj`:

```xml
<Compile Include="PutRangeTests.cs" />
```

## Future Considerations

- **Formula support**: Could add an overload that accepts formula strings
- **Formatting**: Could preserve or set cell styles
- **Error values**: Could add ability to write error values (matching the Error array in RangeResult)
- **ClearRange**: Could add a method to clear a range of cells
