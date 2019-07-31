using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCIta.NPOIHelpers
{
  public class RangeResult
  {
    public object[,] Values;
    public bool[,] Error;
  }

  public static class WorksheetExtension
  {
    public static RangeResult GetRange(this NPOI.SS.UserModel.ISheet sheet)
    {
      int maxcols = 0;
      for(int n = sheet.FirstRowNum; n<=sheet.LastRowNum; n++) {
        var row = sheet.GetRow(n);
        maxcols = Math.Max(maxcols, row.LastCellNum);
      }
      var r = new object[1 + (sheet.LastRowNum - sheet.FirstRowNum), maxcols];
      var e = new bool[r.GetLength(0), r.GetLength(1)];

      for (int n = sheet.FirstRowNum; n <= sheet.LastRowNum; n++) {
        var row = sheet.GetRow(n);
        foreach(var cell in row.Cells) {
          switch(cell.CellType) {
            case NPOI.SS.UserModel.CellType.Numeric:
              r[n, cell.ColumnIndex] = cell.NumericCellValue;
              break;
            case NPOI.SS.UserModel.CellType.Boolean:
              r[n, cell.ColumnIndex] = cell.BooleanCellValue;
              break;
            case NPOI.SS.UserModel.CellType.String:
              r[n, cell.ColumnIndex] = cell.StringCellValue;
              break;
            case NPOI.SS.UserModel.CellType.Error:
              e[n, cell.ColumnIndex] = true;
              break;
            case NPOI.SS.UserModel.CellType.Formula:
              switch (cell.CachedFormulaResultType) {
                case NPOI.SS.UserModel.CellType.String:
                  r[n, cell.ColumnIndex] = cell.StringCellValue;
                  break;
                case NPOI.SS.UserModel.CellType.Boolean:
                  r[n, cell.ColumnIndex] = cell.BooleanCellValue;
                  break;
                case NPOI.SS.UserModel.CellType.Numeric:
                  r[n, cell.ColumnIndex] = cell.NumericCellValue;
                  break;
                case NPOI.SS.UserModel.CellType.Error:
                  e[n, cell.ColumnIndex] = true;
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
