using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GemBox.Spreadsheet;
using ArrayEWE.Helpers;
using System.Linq;

namespace GemBoxTests
{
  [TestClass]
  public class UnitTest1
  {
    [ClassInitialize]
    public static void ClassInit(TestContext context)
    {
      TestDataGenerator.EnsureTestDataExists();
    }

    [TestMethod]
    public void TestGetUsedRangeAddress()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTestPath);
      var ws = ef.Worksheets[0];
      var u = ws.GetUsedCellRange(false);
      Assert.AreEqual(0, u.FirstRowIndex);
      Assert.AreEqual(0, u.FirstColumnIndex);
      Assert.AreEqual(9999, u.LastRowIndex);
      Assert.AreEqual(19, u.LastColumnIndex);
    }

    [TestMethod]
    public void TestCellRange()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTestPath);
      var ws = ef.Worksheets[0];
      var u = ws.GetUsedCellRange(false);
      Assert.AreEqual(0, u.FirstRowIndex);
      Assert.AreEqual(0, u.FirstColumnIndex);
      Assert.AreEqual(9999, u.LastRowIndex);
      Assert.AreEqual(19, u.LastColumnIndex);
    }

    [TestMethod]
    public void TestIntersect()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTestPath);
      var ws = ef.Worksheets[0];
      var u = ws.GetUsedCellRange(false);
      var d = ws.Cells.GetSubrangeAbsolute(3, 3, 30, 30);
      var i = u.Intersect(d);
      Assert.AreEqual(3, i.FirstRowIndex);
      Assert.AreEqual(3, i.FirstColumnIndex);
      Assert.AreEqual(30, i.LastRowIndex);
      Assert.AreEqual(19, i.LastColumnIndex);
    }

    [TestMethod]
    public void TestSheetGet11()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTestPath);
      var ws = ef.Worksheets[0];
      int rows = 3, cols = 4;
      int y = 1, x = 1;

      var r = ws.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, r.Values.GetLength(0));
      Assert.AreEqual(cols, r.Values.GetLength(1));

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual($"Cell: Row-{y - 1};CellNo:{x - 1}", r.Values[0, 0]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", r.Values[rows - 1, cols - 1]);
    }

    [TestMethod]
    public void TestSheetGet22()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTestPath);
      var ws = ef.Worksheets[0];
      int rows = 3, cols = 4;
      int y = 2, x = 2;
      var r = ws.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, r.Values.GetLength(0));
      Assert.AreEqual(cols, r.Values.GetLength(1));

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual($"Cell: Row-{y - 1};CellNo:{x - 1}", r.Values[0, 0]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", r.Values[rows - 1, cols - 1]);
    }

    public void TestSheetGet1_53()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTest1Path);
      var ws = ef.Worksheets[0];
      int rows = 3;
      int cols = 4;
      var r = ws.GetRange(5, 3, rows, cols);
      Assert.AreEqual(0, r.Values.GetLength(0));
      Assert.AreEqual(0, r.Values.GetLength(1));
    }

    [TestMethod]
    public void TestSheetGet1_73()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTest1Path);
      var ws = ef.Worksheets[0];
      int rows = 3, cols = 4;
      int y = 8, x = 3;
      var r = ws.GetRange(y, x, rows, cols);
      Assert.AreEqual(rows, r.Values.GetLength(0));
      Assert.AreEqual(cols, r.Values.GetLength(1));

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual(null, r.Values[0, 0]);
      Assert.AreEqual(null, r.Values[1, 0]);
      Assert.AreEqual(null, r.Values[1, 1]);
      Assert.AreEqual($"Cell: Row-{(y - 1) + (rows - 1)};CellNo:{(x - 1) + (cols - 1)}", r.Values[rows - 1, cols - 1]);
    }

    [TestMethod]
    public void TestSheetGet1_()
    {
      var ef = ExcelFile.Load(TestDataGenerator.BigTest1Path);
      var ws = ef.Worksheets[0];
      var r = ws.GetRange();
      Assert.AreEqual(10000, r.Values.GetLength(0));
      Assert.AreEqual(20, r.Values.GetLength(1));

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual(null, r.Values[0, 0]);
      Assert.AreEqual(null, r.Values[1, 0]);
      Assert.AreEqual(null, r.Values[1, 1]);
      Assert.AreEqual(null, r.Values[22, 4]);
      Assert.AreEqual($"Cell: Row-9999;CellNo:19", r.Values[9999, 19]);
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
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      var input = new object[,] {
        { "A", "B", "C" },
        { 1.0, 2.0, 3.0 },
        { true, false, null }
      };

      ws.PutRange(input, 1, 1);
      var result = ws.GetRange(1, 1, 3, 3);

      Assert.AreEqual("A", result.Values[0, 0]);
      Assert.AreEqual("B", result.Values[0, 1]);
      Assert.AreEqual("C", result.Values[0, 2]);
      Assert.AreEqual(1.0, result.Values[1, 0]);
      Assert.AreEqual(2.0, result.Values[1, 1]);
      Assert.AreEqual(3.0, result.Values[1, 2]);
      Assert.AreEqual(true, result.Values[2, 0]);
      Assert.AreEqual(false, result.Values[2, 1]);
      Assert.AreEqual(null, result.Values[2, 2]);
    }

    [TestMethod]
    public void TestPutRangeOffset()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      var input = new object[,] { { "X" } };
      ws.PutRange(input, 5, 3);

      var result = ws.GetRange(5, 3, 1, 1);
      Assert.AreEqual("X", result.Values[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeOverwrite()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      ws.PutRange(new object[,] { { "Old" } }, 1, 1);
      ws.PutRange(new object[,] { { "New" } }, 1, 1);

      var result = ws.GetRange(1, 1, 1, 1);
      Assert.AreEqual("New", result.Values[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeDataTypes()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      var input = new object[,] {
        { "text", 42.5, true, null },
        { 100, -3.14, false, "" }
      };

      ws.PutRange(input, 1, 1);
      var result = ws.GetRange(1, 1, 2, 4);

      Assert.AreEqual("text", result.Values[0, 0]);
      Assert.AreEqual(42.5, result.Values[0, 1]);
      Assert.AreEqual(true, result.Values[0, 2]);
      Assert.AreEqual(null, result.Values[0, 3]);
      Assert.AreEqual(100.0, result.Values[1, 0]);
      Assert.AreEqual(-3.14, result.Values[1, 1]);
      Assert.AreEqual(false, result.Values[1, 2]);
      Assert.AreEqual("", result.Values[1, 3]);
    }

    [TestMethod]
    public void TestPutRangeLarge()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      int rows = 100, cols = 20;
      var input = new object[rows, cols];
      for (int r = 0; r < rows; r++)
        for (int c = 0; c < cols; c++)
          input[r, c] = $"R{r}C{c}";

      ws.PutRange(input, 1, 1);
      var result = ws.GetRange(1, 1, rows, cols);

      Assert.AreEqual(rows, result.Values.GetLength(0));
      Assert.AreEqual(cols, result.Values.GetLength(1));
      Assert.AreEqual("R0C0", result.Values[0, 0]);
      Assert.AreEqual("R99C19", result.Values[99, 19]);
    }

    [TestMethod]
    public void TestPutRangeEmpty()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      var input = new object[0, 0];
      ws.PutRange(input, 1, 1);

      Assert.IsNull(ws.GetUsedCellRange(false));
    }

    [TestMethod]
    public void TestPutRangeIntegerTypes()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Test");

      var input = new object[,] {
        { (int)1, (long)2, (short)3, (byte)4, (float)5.5 }
      };

      ws.PutRange(input, 1, 1);
      var result = ws.GetRange(1, 1, 1, 5);

      Assert.AreEqual(1.0, result.Values[0, 0]);
      Assert.AreEqual(2.0, result.Values[0, 1]);
      Assert.AreEqual(3.0, result.Values[0, 2]);
      Assert.AreEqual(4.0, result.Values[0, 3]);
      Assert.AreEqual(5.5, (double)result.Values[0, 4], 0.01);
    }

    [TestMethod]
    public void TestAllSheets()
    {
      var ef = new ExcelFile();
      ef.Worksheets.Add("Alpha");
      ef.Worksheets.Add("Beta");
      ef.Worksheets.Add("Gamma");
      var names = WorksheetExtension.AllSheets(ef);
      Assert.AreEqual(3, names.Length);
      Assert.AreEqual("Alpha", names[0]);
      Assert.AreEqual("Beta", names[1]);
      Assert.AreEqual("Gamma", names[2]);
    }

    [TestMethod]
    public void TestAddSheet()
    {
      var ef = new ExcelFile();
      var ws = WorksheetExtension.AddSheet(ef, "NewSheet");
      Assert.IsNotNull(ws);
      Assert.AreEqual("NewSheet", ws.Name);
      Assert.AreEqual(1, ef.Worksheets.Count);
    }

    [TestMethod]
    public void TestGetSheet()
    {
      var ef = new ExcelFile();
      ef.Worksheets.Add("Alpha");
      ef.Worksheets.Add("Beta");
      var ws = WorksheetExtension.GetSheet(ef, "Beta");
      Assert.IsNotNull(ws);
      Assert.AreEqual("Beta", ws.Name);
    }

  }
}
