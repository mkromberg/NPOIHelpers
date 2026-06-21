using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ArrayEWE.Helpers;
using System.Linq;

namespace Tests
{
  [TestClass]
  public class UnitTest1
  {
    [ClassInitialize]
    public static void ClassInit(TestContext context)
    {
      TestDataGenerator.EnsureTestDataExists();
    }

    static object[,] Vals(RangeResult r)    => (object[,])r["Values"];
    static bool[,]   Errs(RangeResult r)    => (bool[,])r["Error"];
    static string[,] Formats(RangeResult r) => (string[,])r["NumberFormat"];
    static bool[,]   HF(RangeResult r)      => (bool[,])r["HasFormula"];
    static string[,] FM(RangeResult r)      => (string[,])r["Formula"];

    static void PutValues(ISheet sheet, object[,] data, int top = 1, int left = 1)
      => sheet.PutRange("Values", data, top, left);

    static void PutWithFormats(ISheet sheet, object[,] data, object[,] formats, int top = 1, int left = 1)
    {
      sheet.PutRange("Values", data, top, left);
      sheet.PutRange("NumberFormat", formats, top, left);
    }

    [TestMethod]
    public void TestGetUsedRangeAddress()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTestPath);
      var wh = wb.GetSheetAt(0);
      var u = wh.GetUsedRangeAddress();
      Assert.AreEqual(0, u.FirstRow);
      Assert.AreEqual(0, u.FirstColumn);
      Assert.AreEqual(9999, u.LastRow);
      Assert.AreEqual(19, u.LastColumn);
    }

    [TestMethod]
    public void TestCellRange()
    {
      var u = new NPOI.SS.Util.CellRangeAddress(0, 9999, 0, 19);
      Assert.AreEqual(0, u.FirstRow);
      Assert.AreEqual(0, u.FirstColumn);
      Assert.AreEqual(9999, u.LastRow);
      Assert.AreEqual(19, u.LastColumn);
    }

    [TestMethod]
    public void TestIntersect()
    {
      var u = new NPOI.SS.Util.CellRangeAddress(0, 9999, 0, 19);
      var d = new NPOI.SS.Util.CellRangeAddress(3, 30, 3, 30);
      var i = u.Intersect(d);
      Assert.AreEqual(3, i.FirstRow);
      Assert.AreEqual(3, i.FirstColumn);
      Assert.AreEqual(30, i.LastRow);
      Assert.AreEqual(19, i.LastColumn);
    }

    [TestMethod]
    public void TestSheetGet11()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTestPath);
      var wh = wb.GetSheetAt(0);
      int rows = 3, cols = 4;
      int y = 1, x = 1;

      var r = wh.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, Vals(r).GetLength(0));
      Assert.AreEqual(cols, Vals(r).GetLength(1));

      var b = false;
      foreach (bool eVal in Errs(r))
        b |= eVal;
      Assert.IsFalse(b);

      Assert.AreEqual($"Cell: Row-{y - 1};CellNo:{x - 1}", Vals(r)[0, 0]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", Vals(r)[rows - 1, cols - 1]);
    }

    [TestMethod]
    public void TestSheetGet22()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTestPath);
      var wh = wb.GetSheetAt(0);
      int rows = 3, cols = 4;
      int y = 2, x = 2;
      var r = wh.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, Vals(r).GetLength(0));
      Assert.AreEqual(cols, Vals(r).GetLength(1));

      var b = false;
      foreach (bool eVal in Errs(r))
        b |= eVal;
      Assert.IsFalse(b);

      Assert.AreEqual($"Cell: Row-{y - 1};CellNo:{x - 1}", Vals(r)[0, 0]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", Vals(r)[rows - 1, cols - 1]);
    }

    public void TestSheetGet1_53()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      var r = wh.GetRange(5, 3, 3, 4);
      Assert.IsNull(r["Values"]);
    }

    [TestMethod]
    public void TestSheetGet1_73()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      int rows = 3, cols = 4;
      int y = 8, x = 3;
      var r = wh.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, Vals(r).GetLength(0));
      Assert.AreEqual(cols, Vals(r).GetLength(1));

      var b = false;
      foreach (bool eVal in Errs(r))
        b |= eVal;
      Assert.IsFalse(b);

      Assert.AreEqual(null, Vals(r)[0, 0]);
      Assert.AreEqual(null, Vals(r)[1, 0]);
      Assert.AreEqual(null, Vals(r)[1, 1]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", Vals(r)[rows - 1, cols - 1]);
    }

    [TestMethod]
    public void TestSheetGet1_()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      var r = wh.GetRange();
      Assert.AreEqual(10000, Vals(r).GetLength(0));
      Assert.AreEqual(20, Vals(r).GetLength(1));

      var b = false;
      foreach (bool eVal in Errs(r))
        b |= eVal;
      Assert.IsFalse(b);

      Assert.AreEqual(null, Vals(r)[0, 0]);
      Assert.AreEqual(null, Vals(r)[1, 0]);
      Assert.AreEqual(null, Vals(r)[1, 1]);
      Assert.AreEqual(null, Vals(r)[22, 4]);
      Assert.AreEqual($"Cell: Row-9999;CellNo:19", Vals(r)[9999, 19]);
    }

    [TestMethod]
    public void TestCreateWorkbooks()
    {
      var (bigTestMs, bigTest1Ms) = TestDataGenerator.BenchmarkCreation();
      Console.WriteLine($"bigtest.xlsx:  {bigTestMs} ms");
      Console.WriteLine($"bigtest1.xlsx: {bigTest1Ms} ms");
    }

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

      PutValues(sheet, input, 1, 1);
      var result = sheet.GetRange(1, 1, 3, 3);

      Assert.AreEqual("A", Vals(result)[0, 0]);
      Assert.AreEqual("B", Vals(result)[0, 1]);
      Assert.AreEqual("C", Vals(result)[0, 2]);
      Assert.AreEqual(1.0, Vals(result)[1, 0]);
      Assert.AreEqual(2.0, Vals(result)[1, 1]);
      Assert.AreEqual(3.0, Vals(result)[1, 2]);
      Assert.AreEqual(true, Vals(result)[2, 0]);
      Assert.AreEqual(false, Vals(result)[2, 1]);
      Assert.AreEqual(null, Vals(result)[2, 2]);
    }

    [TestMethod]
    public void TestPutRangeOffset()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      var input = new object[,] { { "X" } };
      PutValues(sheet, input, 5, 3);

      var result = sheet.GetRange(5, 3, 1, 1);
      Assert.AreEqual("X", Vals(result)[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeOverwrite()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      PutValues(sheet, new object[,] { { "Old" } }, 1, 1);
      PutValues(sheet, new object[,] { { "New" } }, 1, 1);

      var result = sheet.GetRange(1, 1, 1, 1);
      Assert.AreEqual("New", Vals(result)[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeDataTypes()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      var input = new object[,] {
        { "text", 42.5, true, null },
        { 100, -3.14, false, "" }
      };

      PutValues(sheet, input, 1, 1);
      var result = sheet.GetRange(1, 1, 2, 4);

      Assert.AreEqual("text", Vals(result)[0, 0]);
      Assert.AreEqual(42.5, Vals(result)[0, 1]);
      Assert.AreEqual(true, Vals(result)[0, 2]);
      Assert.AreEqual(null, Vals(result)[0, 3]);
      Assert.AreEqual(100.0, Vals(result)[1, 0]);
      Assert.AreEqual(-3.14, Vals(result)[1, 1]);
      Assert.AreEqual(false, Vals(result)[1, 2]);
      Assert.AreEqual("", Vals(result)[1, 3]);
    }

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

      PutValues(sheet, input, 1, 1);
      var result = sheet.GetRange(1, 1, rows, cols);

      Assert.AreEqual(rows, Vals(result).GetLength(0));
      Assert.AreEqual(cols, Vals(result).GetLength(1));
      Assert.AreEqual("R0C0", Vals(result)[0, 0]);
      Assert.AreEqual("R99C19", Vals(result)[99, 19]);
    }

    [TestMethod]
    public void TestPutRangeEmpty()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      PutValues(sheet, new object[0, 0], 1, 1);

      Assert.AreEqual(0, sheet.PhysicalNumberOfRows);
    }

    [TestMethod]
    public void TestPutRangeIntegerTypes()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      var input = new object[,] {
        { (int)1, (long)2, (short)3, (byte)4, (float)5.5 }
      };

      PutValues(sheet, input, 1, 1);
      var result = sheet.GetRange(1, 1, 1, 5);

      Assert.AreEqual(1.0, Vals(result)[0, 0]);
      Assert.AreEqual(2.0, Vals(result)[0, 1]);
      Assert.AreEqual(3.0, Vals(result)[0, 2]);
      Assert.AreEqual(4.0, Vals(result)[0, 3]);
      Assert.AreEqual(5.5, (double)Vals(result)[0, 4], 0.01);
    }

    [TestMethod]
    public void TestAllSheets()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Alpha");
      wb.CreateSheet("Beta");
      wb.CreateSheet("Gamma");
      var names = WorksheetExtension.AllSheets(wb);
      Assert.AreEqual(3, names.Length);
      Assert.AreEqual("Alpha", names[0]);
      Assert.AreEqual("Beta", names[1]);
      Assert.AreEqual("Gamma", names[2]);
    }

    [TestMethod]
    public void TestAddSheet()
    {
      var wb = new XSSFWorkbook();
      var sheet = WorksheetExtension.AddSheet(wb, "NewSheet");
      Assert.IsNotNull(sheet);
      Assert.AreEqual("NewSheet", sheet.SheetName);
      Assert.AreEqual(1, wb.NumberOfSheets);
    }

    [TestMethod]
    public void TestGetSheet()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Alpha");
      wb.CreateSheet("Beta");
      var sheet = WorksheetExtension.GetSheet(wb, "Beta");
      Assert.IsNotNull(sheet);
      Assert.AreEqual("Beta", sheet.SheetName);
    }

    [TestMethod]
    public void TestGetSheetCaseInsensitive()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Alpha");
      wb.CreateSheet("Beta");
      Assert.AreEqual("Alpha", WorksheetExtension.GetSheet(wb, "alpha").SheetName);
      Assert.AreEqual("Beta",  WorksheetExtension.GetSheet(wb, "BETA").SheetName);
      Assert.IsNull(WorksheetExtension.GetSheet(wb, "Gamma"));
    }

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
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "Hello", "World" } }, 1, 1);
      WorksheetExtension.Save(wb, path);

      Assert.IsTrue(File.Exists(path));

      var result = WorkbookFactory.Create(path).GetSheetAt(0).GetRange(1, 1, 1, 2);
      Assert.AreEqual("Hello", Vals(result)[0, 0]);
      Assert.AreEqual("World", Vals(result)[0, 1]);

      File.Delete(path);
    }

    [TestMethod]
    public void TestClearSheet()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B" }, { "C", "D" } }, 1, 1);
      Assert.AreNotEqual(0, sheet.PhysicalNumberOfRows);
      sheet.ClearSheet();
      Assert.AreEqual(0, sheet.PhysicalNumberOfRows);
    }

    [TestMethod]
    public void TestClearHiddenRows()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A" }, { "B" }, { "C" } }, 1, 1);
      sheet.GetRow(1).ZeroHeight = true;
      sheet.ClearHidden();
      Assert.IsNull(sheet.GetRow(1));
      Assert.IsNotNull(sheet.GetRow(0));
      Assert.IsNotNull(sheet.GetRow(2));
    }

    [TestMethod]
    public void TestClearHiddenColumns()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B", "C" } }, 1, 1);
      sheet.SetColumnHidden(1, true);
      sheet.ClearHidden();
      var cell = sheet.GetRow(0)?.GetCell(1);
      Assert.IsTrue(cell == null || cell.CellType == NPOI.SS.UserModel.CellType.Blank);
    }

    [TestMethod]
    public void TestAutoFitColumns()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "Short", "A much longer string value here" } }, 1, 1);
      sheet.AutoFitColumns();
      Assert.IsTrue(sheet.GetColumnWidth(0) > 0);
      Assert.IsTrue(sheet.GetColumnWidth(1) > 0);
    }

    [TestMethod]
    public void TestAppendRange()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A" }, { "B" } }, 1, 1);
      sheet.AppendRange(new object[,] { { "C" }, { "D" } });
      var result = sheet.GetRange(1, 1, 4, 1);
      Assert.AreEqual("A", Vals(result)[0, 0]);
      Assert.AreEqual("B", Vals(result)[1, 0]);
      Assert.AreEqual("C", Vals(result)[2, 0]);
      Assert.AreEqual("D", Vals(result)[3, 0]);
    }

    [TestMethod]
    public void TestAppendRangeEmptySheet()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      sheet.AppendRange(new object[,] { { "A" } });
      var result = sheet.GetRange(1, 1, 1, 1);
      Assert.AreEqual("A", Vals(result)[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeWithFormats()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      var values = new object[,] { { 1234.5, 0.75 } };
      var formats = new object[,] { { "#,##0.00", "0.00%" } };
      PutWithFormats(sheet, values, formats, 1, 1);
      var result = sheet.GetRange(1, 1, 1, 2);
      Assert.AreEqual(1234.5, Vals(result)[0, 0]);
      Assert.AreEqual(0.75, Vals(result)[0, 1]);
      Assert.AreEqual("#,##0.00", Formats(result)[0, 0]);
      Assert.AreEqual("0.00%", Formats(result)[0, 1]);
    }

    [TestMethod]
    public void TestGetRangeFormulas()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      var row = sheet.CreateRow(0);
      row.CreateCell(0).SetCellValue(10.0);
      row.CreateCell(1).SetCellValue(20.0);
      row.CreateCell(2).SetCellFormula("A1+B1");
      wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

      var result = sheet.GetRange(1, 1, 1, 3);
      Assert.IsFalse(HF(result)[0, 0]);
      Assert.IsFalse(HF(result)[0, 1]);
      Assert.IsTrue(HF(result)[0, 2]);
      Assert.AreEqual("A1+B1", FM(result)[0, 2]);
      Assert.AreEqual(30.0, Vals(result)[0, 2]);
    }

    [TestMethod]
    public void TestGetUsedRange()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B" }, { "C", "D" } }, 1, 1);
      var used = sheet.GetUsedRange();
      Assert.AreEqual(2, Vals(used).GetLength(0));
      Assert.AreEqual(2, Vals(used).GetLength(1));
      Assert.AreEqual("A", Vals(used)[0, 0]);
      Assert.AreEqual("D", Vals(used)[1, 1]);
    }

    [TestMethod]
    public void TestGetUsedRangeEmpty()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      var used = sheet.GetUsedRange();
      Assert.IsNull(used["Values"]);
    }

    [TestMethod]
    public void TestAddSheetFirst()
    {
      var wb = new XSSFWorkbook();
      WorksheetExtension.AddSheet(wb, "Alpha");
      WorksheetExtension.AddSheet(wb, "Beta");
      WorksheetExtension.AddSheet(wb, "First", first: true);
      Assert.AreEqual("First", wb.GetSheetName(0));
      Assert.AreEqual(3, wb.NumberOfSheets);
    }

    [TestMethod]
    public void TestAddSheetLast()
    {
      var wb = new XSSFWorkbook();
      WorksheetExtension.AddSheet(wb, "Alpha");
      WorksheetExtension.AddSheet(wb, "Beta");
      Assert.AreEqual("Beta", wb.GetSheetName(1));
    }

    [TestMethod]
    public void TestDeleteSheet()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Alpha");
      wb.CreateSheet("Beta");
      wb.CreateSheet("Gamma");
      WorksheetExtension.DeleteSheet(wb, "Beta");
      Assert.AreEqual(2, wb.NumberOfSheets);
      Assert.AreEqual("Alpha", wb.GetSheetName(0));
      Assert.AreEqual("Gamma", wb.GetSheetName(1));
    }

    [TestMethod]
    public void TestDeleteSheetNotFound()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Alpha");
      Assert.ThrowsException<ArgumentException>(() => WorksheetExtension.DeleteSheet(wb, "NotExist"));
    }

    [TestMethod]
    public void TestGetRangeNamedPropertiesSelectsCorrectly()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", 1.0 } }, 1, 1);

      var result = sheet.GetRange(new[] { "Values" });
      Assert.AreEqual(1, Vals(result).GetLength(0));
      Assert.IsNull(result["Error"]);
      Assert.IsNull(result["NumberFormat"]);
      Assert.IsNull(result["HasFormula"]);
      Assert.IsNull(result["Formula"]);
    }

    [TestMethod]
    public void TestGetRangeNamedPropertiesMultiple()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutWithFormats(sheet, new object[,] { { 1234.5 } }, new object[,] { { "#,##0.00" } }, 1, 1);

      var result = sheet.GetRange(new[] { "Values", "NumberFormat" });
      Assert.AreEqual(1234.5, Vals(result)[0, 0]);
      Assert.AreEqual("#,##0.00", Formats(result)[0, 0]);
      Assert.IsNull(result["Error"]);
    }

    [TestMethod]
    public void TestGetRangeNamedPropertiesWithCoords()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "X", "Y" }, { "Z", "W" } }, 1, 1);

      var result = sheet.GetRange(new[] { "Values" }, top: 2, left: 2, rows: 1, cols: 1);
      Assert.AreEqual(1, Vals(result).GetLength(0));
      Assert.AreEqual(1, Vals(result).GetLength(1));
      Assert.AreEqual("W", Vals(result)[0, 0]);
    }

    [TestMethod]
    public void TestGetRangeNamedPropertiesUnknownThrows()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A" } }, 1, 1);
      Assert.ThrowsException<ArgumentException>(() => sheet.GetRange(new[] { "Unknown" }));
    }

    [TestMethod]
    public void TestRangeResultNamesPreserveOrder()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A" } }, 1, 1);

      var result = sheet.GetRange(new[] { "Formula", "Values", "Error" });
      Assert.AreEqual("Formula",  result.Names[0]);
      Assert.AreEqual("Values",   result.Names[1]);
      Assert.AreEqual("Error",    result.Names[2]);
    }

    [TestMethod]
    public void TestPutRangeNamed()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      sheet.PutRange("Values", new object[,] { { "Hello", 42.0 } }, 1, 1);

      var result = sheet.GetRange(1, 1, 1, 2);
      Assert.AreEqual("Hello", Vals(result)[0, 0]);
      Assert.AreEqual(42.0, Vals(result)[0, 1]);
    }

    [TestMethod]
    public void TestPutRangeNamedWithFormats()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      sheet.PutRange("Values", new object[,] { { 0.5 } }, 1, 1);
      sheet.PutRange("NumberFormat", new object[,] { { "0.00%" } }, 1, 1);

      var result = sheet.GetRange(new[] { "Values", "NumberFormat" });
      Assert.AreEqual(0.5, Vals(result)[0, 0]);
      Assert.AreEqual("0.00%", Formats(result)[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeNamedUnknownThrows()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      Assert.ThrowsException<ArgumentException>(() =>
        sheet.PutRange("Unknown", new object[,] { { 1.0 } }, 1, 1));
    }

    static double ExpectedPoints(ISheet sheet, int col) =>
      (sheet.GetColumnWidthInPixels(col) + 5.0) * 0.75;

    [TestMethod]
    public void TestGetColumnWidthsDefault()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B", "C" } }, 1, 1);

      var widths = sheet.GetColumnWidths();
      Assert.AreEqual(3, widths.Length);
      for (int c = 0; c < 3; c++)
        Assert.AreEqual(ExpectedPoints(sheet, c), widths[c], 0.01);
    }

    [TestMethod]
    public void TestGetColumnWidthsExplicit()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B", "C" } }, 1, 1);
      sheet.SetColumnWidth(0, (int)(10.0 * 256));
      sheet.SetColumnWidth(1, (int)(20.0 * 256));

      var widths = sheet.GetColumnWidths();
      Assert.AreEqual(3, widths.Length);
      Assert.AreEqual(ExpectedPoints(sheet, 0), widths[0], 0.01);
      Assert.AreEqual(ExpectedPoints(sheet, 1), widths[1], 0.01);
      // Wider column should produce larger point value
      Assert.IsTrue(widths[1] > widths[0]);
    }

    [TestMethod]
    public void TestGetColumnWidthsRange()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      PutValues(sheet, new object[,] { { "A", "B", "C", "D", "E" } }, 1, 1);
      sheet.SetColumnWidth(1, (int)(15.0 * 256));
      sheet.SetColumnWidth(2, (int)(25.0 * 256));

      var widths = sheet.GetColumnWidths(left: 2, cols: 2);
      Assert.AreEqual(2, widths.Length);
      Assert.AreEqual(ExpectedPoints(sheet, 1), widths[0], 0.01);
      Assert.AreEqual(ExpectedPoints(sheet, 2), widths[1], 0.01);
    }

    [TestMethod]
    public void TestGetColumnWidthsEmptySheet()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");
      var widths = sheet.GetColumnWidths();
      Assert.AreEqual(0, widths.Length);
    }

    [TestMethod]
    public void TestProtectWorkbook()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Sheet1");
      Assert.IsFalse(wb.IsStructureLocked());
      WorksheetExtension.Protect(wb, "secret");
      Assert.IsTrue(wb.IsStructureLocked());
      WorksheetExtension.Protect(wb, null);
      Assert.IsFalse(wb.IsStructureLocked());
    }

    [TestMethod]
    public void TestProtectWorkbookPasswordHashSet()
    {
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Sheet1");
      WorksheetExtension.Protect(wb, "test");
      var prot = wb.GetCTWorkbook().workbookProtection;
      Assert.IsNotNull(prot);
      Assert.IsTrue(prot.lockStructure);
      Assert.IsNotNull(prot.workbookPassword);
      Assert.AreEqual(2, prot.workbookPassword.Length);
    }

    [TestMethod]
    public void TestProtectWorkbookRoundTrip()
    {
      var path = Path.Combine(TestDataGenerator.TestDataPath, "protect_wb_test.xlsx");
      var wb = new XSSFWorkbook();
      wb.CreateSheet("Sheet1");
      WorksheetExtension.Protect(wb, "secret");
      WorksheetExtension.Save(wb, path);

      var wb2 = (XSSFWorkbook)WorksheetExtension.Open(path);
      Assert.IsTrue(wb2.IsStructureLocked());

      File.Delete(path);
    }

    [TestMethod]
    public void TestProtectSheet()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Sheet1");
      Assert.IsFalse(sheet.Protect);
      WorksheetExtension.ProtectSheet(sheet, "secret");
      Assert.IsTrue(sheet.Protect);
      WorksheetExtension.ProtectSheet(sheet, null);
      Assert.IsFalse(sheet.Protect);
    }

    [TestMethod]
    public void TestProtectSheetRoundTrip()
    {
      var path = Path.Combine(TestDataGenerator.TestDataPath, "protect_sheet_test.xlsx");
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Sheet1");
      PutValues(sheet, new object[,] { { "Protected data" } }, 1, 1);
      WorksheetExtension.ProtectSheet(sheet, "secret");
      WorksheetExtension.Save(wb, path);

      var wb2 = WorksheetExtension.Open(path);
      Assert.IsTrue(wb2.GetSheetAt(0).Protect);

      File.Delete(path);
    }

  }
}
