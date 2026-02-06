using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SCIta.Helpers;

namespace Tests
{
  [TestClass]
  public class PutRangeTests
  {
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
      sheet.PutRange(input, 5, 3);  // Row 5, Column C

      var result = sheet.GetRange(5, 3, 1, 1);
      Assert.AreEqual("X", result.Values[0, 0]);
    }

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
      sheet.PutRange(input, 1, 1);  // Should not throw

      // Sheet should remain empty (no rows created)
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
