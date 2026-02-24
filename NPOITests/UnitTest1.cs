using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SCIta.Helpers;
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
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTestPath);
      var wh = wb.GetSheetAt(0);
      int rows = 3, cols = 4;
      int y = 2, x = 2;
      var r = wh.GetRange(y, x, rows, cols);
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
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      int rows = 3;
      int cols = 4;
      var r = wh.GetRange(5, 3, rows, cols);
      Assert.AreEqual(0, r.Values.GetLength(0));
      Assert.AreEqual(0, r.Values.GetLength(1));
    }


    [TestMethod]
    public void TestSheetGet1_73()
    {
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      int rows = 3, cols = 4;
      int y = 8, x = 3;
      var r = wh.GetRange(y, x, rows, cols);
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
      var wb = WorkbookFactory.Create(TestDataGenerator.BigTest1Path);
      var wh = wb.GetSheetAt(0);
      var r = wh.GetRange();
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
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      var input = new object[,] { { "X" } };
      sheet.PutRange(input, 5, 3);

      var result = sheet.GetRange(5, 3, 1, 1);
      Assert.AreEqual("X", result.Values[0, 0]);
    }

    [TestMethod]
    public void TestPutRangeOverwrite()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      sheet.PutRange(new object[,] { { "Old" } }, 1, 1);
      sheet.PutRange(new object[,] { { "New" } }, 1, 1);

      var result = sheet.GetRange(1, 1, 1, 1);
      Assert.AreEqual("New", result.Values[0, 0]);
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

      sheet.PutRange(input, 1, 1);
      var result = sheet.GetRange(1, 1, 2, 4);

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

    [TestMethod]
    public void TestPutRangeEmpty()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Test");

      var input = new object[0, 0];
      sheet.PutRange(input, 1, 1);

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

      sheet.PutRange(input, 1, 1);
      var result = sheet.GetRange(1, 1, 1, 5);

      Assert.AreEqual(1.0, result.Values[0, 0]);
      Assert.AreEqual(2.0, result.Values[0, 1]);
      Assert.AreEqual(3.0, result.Values[0, 2]);
      Assert.AreEqual(4.0, result.Values[0, 3]);
      Assert.AreEqual(5.5, (double)result.Values[0, 4], 0.01);
    }

  }
}
