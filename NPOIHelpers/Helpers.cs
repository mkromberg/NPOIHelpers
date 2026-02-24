using System;
using SS = NPOI.SS;

namespace ArrayEWE.Helpers
{
  public static class WorksheetExtension
  {
    public static void Initialize()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    public static SS.Util.CellRangeAddress GetUsedRangeAddress(this SS.UserModel.ISheet sheet)
    {
      int lastcol = 0, firstcol = int.MaxValue;

      for (int n = sheet.FirstRowNum; n <= sheet.LastRowNum; n++) {
        var row = sheet.GetRow(n);
        if (row != null) {
          firstcol = Math.Min(firstcol, row.FirstCellNum);
          lastcol = Math.Max(lastcol, row.LastCellNum - 1);
        }
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
        if (row == null)
          continue;

        for (int c = 0; c < cols; c++) {
          int x = c + rng.FirstColumn;
          if (x < row.FirstCellNum || x>=row.LastCellNum)
            continue;

          var cell = row.GetCell(x);
          if (cell == null)
            continue;

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

    public static string[] AllSheets(SS.UserModel.IWorkbook wb)
    {
      var names = new string[wb.NumberOfSheets];
      for (int i = 0; i < wb.NumberOfSheets; i++)
        names[i] = wb.GetSheetName(i);
      return names;
    }

    public static SS.UserModel.ISheet AddSheet(SS.UserModel.IWorkbook wb, string name)
    {
      return wb.CreateSheet(name);
    }

    public static SS.UserModel.ISheet GetSheet(SS.UserModel.IWorkbook wb, string name)
    {
      return wb.GetSheet(name);
    }

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
              cell.SetCellValue(Convert.ToDouble(value));
              break;
          }
        }
      }
    }
  }
}
