using System;
using GemBox.Spreadsheet;

namespace ArrayEWE.Helpers
{
  public static class WorksheetExtension
  {
    public static void Initialize()
    {
      var key = GetEnvironmentVariable("GEMBOX_LICENCE_KEY") ?? GetEnvironmentVariable("GEMBOX_LICENSE_KEY") ?? "FREE-LIMITED-KEY";
      SpreadsheetInfo.SetLicense(key);
    }

    public static string GetEnvironmentVariable(string name)
    {
      return Environment.GetEnvironmentVariable(name);
    }

    public static CellRange GetUsedCellRange(ExcelWorksheet worksheet)
    {
      return worksheet.GetUsedCellRange(false);
    }

    public static RangeResult GetRange(this ExcelWorksheet worksheet)
    {
      return GetRange(worksheet, 1);
    }

    public static RangeResult GetRange(this ExcelWorksheet worksheet, int top)
    {
      return GetRange(worksheet, top, 1);
    }

    public static RangeResult GetRange(this ExcelWorksheet worksheet, int top, int left)
    {
      return GetRange(worksheet, top, left, int.MaxValue);
    }

    public static RangeResult GetRange(this ExcelWorksheet worksheet, int top, int left, int rows)
    {
      return GetRange(worksheet, top, left, rows, int.MaxValue);
    }

    public static RangeResult GetRange(this ExcelWorksheet worksheet, int top = 1, int left = 1, int rows = int.MaxValue, int cols = int.MaxValue)
    {
      var u = worksheet.GetUsedCellRange(false);

      int Clamp(int x, int w, int max)
      {
        return w == int.MaxValue ? max : (x - 1) + (w - 1);
      }

      int bottom = Clamp(top, rows, u.LastRowIndex);
      int right = Clamp(left, cols, u.LastColumnIndex);

      // Compute intersection of used range with requested range
      int firstRow = Math.Max(u.FirstRowIndex, top - 1);
      int lastRow = Math.Min(u.LastRowIndex, bottom);
      int firstCol = Math.Max(u.FirstColumnIndex, left - 1);
      int lastCol = Math.Min(u.LastColumnIndex, right);

      if (firstRow > lastRow || firstCol > lastCol)
        return new RangeResult();

      // Output array anchored at requested top-left
      int reqFirstRow = top - 1;
      int reqFirstCol = left - 1;
      int rrows = lastRow - reqFirstRow + 1;
      int rcols = lastCol - reqFirstCol + 1;

      var r = new object[rrows, rcols];
      var e = new bool[rrows, rcols];

      for (int n = firstRow; n <= lastRow; n++) {
        int rr = n - reqFirstRow;
        for (int c = firstCol; c <= lastCol; c++) {
          int rc = c - reqFirstCol;
          var cell = worksheet.Cells[n, c];
          switch (cell.ValueType) {
            case CellValueType.Double:
            case CellValueType.Int:
              r[rr, rc] = cell.DoubleValue;
              break;
            case CellValueType.String:
              r[rr, rc] = cell.StringValue;
              break;
            case CellValueType.Bool:
              r[rr, rc] = cell.BoolValue;
              break;
            case CellValueType.Error:
              e[rr, rc] = true;
              break;
            // CellValueType.Null: leave as null (default)
          }
        }
      }

      return new RangeResult { Values = r, Error = e };
    }

    public static string[] AllSheets(ExcelFile ef)
    {
      var names = new string[ef.Worksheets.Count];
      for (int i = 0; i < ef.Worksheets.Count; i++)
        names[i] = ef.Worksheets[i].Name;
      return names;
    }

    public static ExcelWorksheet AddSheet(ExcelFile ef, string name)
    {
      return ef.Worksheets.Add(name);
    }

    public static ExcelWorksheet GetSheet(ExcelFile ef, string name)
    {
      return ef.Worksheets[name];
    }

    public static void PutRange(this ExcelWorksheet worksheet, object[,] values, int top = 1, int left = 1)
    {
      int rows = values.GetLength(0);
      int cols = values.GetLength(1);

      for (int r = 0; r < rows; r++) {
        int rowIndex = (top - 1) + r;
        for (int c = 0; c < cols; c++) {
          int colIndex = (left - 1) + c;
          var value = values[r, c];
          switch (value) {
            case null:
              break;
            case bool b:
              worksheet.Cells[rowIndex, colIndex].Value = b;
              break;
            case DateTime dt:
              worksheet.Cells[rowIndex, colIndex].Value = dt;
              break;
            case double d:
              worksheet.Cells[rowIndex, colIndex].Value = d;
              break;
            case string s:
              worksheet.Cells[rowIndex, colIndex].Value = s;
              break;
            default:
              worksheet.Cells[rowIndex, colIndex].Value = Convert.ToDouble(value);
              break;
          }
        }
      }
    }
  }
}
