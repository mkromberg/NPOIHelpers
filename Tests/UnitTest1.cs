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
      int rows = 3;
      int cols = 4;
      var r = wh.GetRange(1, 1, rows, cols);
      Assert.AreEqual(r.Values.GetLength(0), rows);
      Assert.AreEqual(r.Values.GetLength(1), cols);

      var b = false;
      foreach(var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual(r.Values[0,0], "Cell: Row-0;CellNo:0");
      Assert.AreEqual(r.Values[2, 3], "Cell: Row-2;CellNo:3");
    }

    [TestMethod]
    public void TestSheetGet22()
    {
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest.xlsx");
      var wh = wb.GetSheetAt(0);
      int rows = 3;
      int cols = 4;
      var r = wh.GetRange(2, 2, rows, cols);
      Assert.AreEqual(r.Values.GetLength(0), rows);
      Assert.AreEqual(r.Values.GetLength(1), cols);

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual(r.Values[0, 0], "Cell: Row-1;CellNo:1");
      Assert.AreEqual(r.Values[2, 3], "Cell: Row-3;CellNo:4");
    }

    [TestMethod]
    public void TestSheetGet1_53()
    {
      var wb = WorkbookFactory.Create(@"k:\users\stefano\bigtest1.xlsx");
      var wh = wb.GetSheetAt(0);
      int rows = 3;
      int cols = 4;
      var r = wh.GetRange(5, 3, rows, cols);
      Assert.AreEqual(r.Values.GetLength(0), rows);
      Assert.AreEqual(r.Values.GetLength(1), cols);

      var b = false;
      foreach (var e in r.Error) {
        b |= e;
      }
      Assert.IsFalse(b);

      Assert.AreEqual(r.Values[0, 0], "Cell: Row-4;CellNo:2");
      Assert.AreEqual(r.Values[2, 3], "Cell: Row-6;CellNo:5");
    }
  }
}
