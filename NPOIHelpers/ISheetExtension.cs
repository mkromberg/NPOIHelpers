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
        firstcol = Math.Min(firstcol, row.FirstCellNum);
        lastcol = Math.Max(lastcol, row.LastCellNum - 1);
      }

      return new SS.Util.CellRangeAddress(sheet.FirstRowNum, sheet.LastRowNum, firstcol, lastcol);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet)
    {
      return GetRange(sheet, 1);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, int top)
    {
      return GetRange(sheet, top, 1);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, int top, int left)
    {
      return GetRange(sheet, top, left, int.MaxValue);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, int top, int left, int rows)
    {
      return GetRange(sheet, top, left, rows, int.MaxValue);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, int top = 1, int left = 1, int rows = int.MaxValue, int cols = int.MaxValue)
    {
      var u = sheet.GetUsedRangeAddress();

      int Clamp(int x, int w, int max)
      {
        if (w == int.MaxValue)
          return max;
        else
          return (x - 1) + (w - 1);
      }
      int bottom = Clamp(top, rows, u.LastRow), 
        right = Clamp(left, cols, u.LastColumn);

      var wanted = new SS.Util.CellRangeAddress(top - 1, bottom, left - 1, right);

      var rng = u.Intersect(wanted);
      if (rng == null) {
        return new RangeResult();
      }

      var rrng = u.Intersect(wanted);

      rrng.FirstRow = top - 1;
      rrng.FirstColumn = left - 1;

      var (rrows, rcols) = rrng.GetSize();

      var r = new object[rrows, rcols];
      var e = new bool[r.GetLength(0), r.GetLength(1)];

      (rows, cols) = rng.GetSize();

      int or = rng.FirstRow - rrng.FirstRow, 
        oc = rng.FirstColumn - rrng.FirstColumn;

      for (int n = 0; n < rows; n++) {
        var rr = n + or;
        var row = sheet.GetRow(n + rng.FirstRow);
        for (int c = 0; c < cols; c++) {
          int x = c + rng.FirstColumn;
          if (x < row.FirstCellNum)
            continue;

          var cell = row.GetCell(x);
          var rc = c + oc;
          switch (cell.CellType) {
            case SS.UserModel.CellType.Numeric:
              r[rr, rc] = cell.NumericCellValue;
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
