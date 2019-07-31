using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.SCIta.Helpers;
using System.Linq;

namespace Tests
{
  [TestClass]
  public class UnitTest1
  {

    [TestMethod]
    public void TestGetUsedRangeAddress()
    {
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest.xlsx");
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
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest.xlsx");
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
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest.xlsx");
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
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest1.xlsx");
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
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest1.xlsx");
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
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest1.xlsx");
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

  }
}
