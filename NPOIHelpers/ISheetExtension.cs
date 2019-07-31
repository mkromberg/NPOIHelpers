using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SCIta.Helpers
{
  public static class ISheetExtension
  {
    public static SS.Util.CellRangeAddress GetUsedRangeAddress(this SS.UserModel.ISheet sheet)
    {
      int lastcol = 0, firstcol = int.MaxValue;
      
      for (int n = sheet.FirstRowNum; n <= sheet.LastRowNum; n++) {
        var row = sheet.GetRow(n);
        lastcol = Math.Max(lastcol, row.LastCellNum);
        firstcol = Math.Min(firstcol, row.FirstCellNum);
      }

      return new SS.Util.CellRangeAddress(sheet.FirstRowNum, sheet.LastRowNum, firstcol, lastcol);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet,int top=1,int left=1, int rows = int.MaxValue, int cols = int.MaxValue)
    {
      var u = sheet.GetUsedRangeAddress();
      var (h, w) = u.GetSize();
      rows = Math.Min(h, rows);
      cols = Math.Min(w, cols);
      var wanted = new SS.Util.CellRangeAddress(top - 1, (top - 1) + (rows - 1), left - 1, (left - 1) + (cols - 1));
      var rng = u.Intersect(wanted);
      if(rng.NumberOfCells == 0) {
        return new RangeResult();
      }

      rng.FirstRow = top - 1;
      rng.FirstColumn = left - 1;

      (rows, cols) = rng.GetSize();

      var r = new object[rows, cols];
      var e = new bool[r.GetLength(0), r.GetLength(1)];

      rng = u.Intersect(wanted);

      for (int n = rng.FirstRow; n <= rng.LastRow; n++) {
        var rr = n - (rng.FirstRow - (top - 1));
        var row = sheet.GetRow(n);
        for(int c = rng.FirstColumn; c <= rng.LastColumn; c++) {
          var rc = c - (rng.FirstColumn - (left-1));
          var cell = row.GetCell(c);
          switch(cell.CellType) {
            case SS.UserModel.CellType.Numeric:
              r[rr,rc] = cell.NumericCellValue;
              break;
            case SS.UserModel.CellType.Boolean:
              r[rr, rc] = cell.BooleanCellValue;
              break;
            case SS.UserModel.CellType.String:
              r[rr, rc] = cell.StringCellValue;
              break;
            case SS.UserModel.CellType.Error:
              e[rr, rc] = true;
              break;
            case SS.UserModel.CellType.Formula:
              switch (cell.CachedFormulaResultType) {
                case SS.UserModel.CellType.String:
                  r[rr, rc] = cell.StringCellValue;
                  break;
                case SS.UserModel.CellType.Boolean:
                  r[rr, rc] = cell.BooleanCellValue;
                  break;
                case SS.UserModel.CellType.Numeric:
                  r[rr, rc] = cell.NumericCellValue;
                  break;
                case SS.UserModel.CellType.Error:
                  e[rr, rc] = true;
                  break;
              }
              break;
          }
        }
      }
      return new RangeResult {
        Error = e,
        Values = r
      };
    }
  }
}
