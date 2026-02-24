using System;
using GemBox.Spreadsheet;

namespace SCIta.GemBoxHelpers
{
  public static class CellRangeExtension
  {
    public static CellRange Intersect(this CellRange x, CellRange y)
    {
      int firstRow = Math.Max(x.FirstRowIndex, y.FirstRowIndex);
      int lastRow = Math.Min(x.LastRowIndex, y.LastRowIndex);
      int firstCol = Math.Max(x.FirstColumnIndex, y.FirstColumnIndex);
      int lastCol = Math.Min(x.LastColumnIndex, y.LastColumnIndex);

      if (firstRow <= lastRow && firstCol <= lastCol)
        return x.GetSubrangeAbsolute(firstRow, firstCol, lastRow, lastCol);

      return null;
    }

    public static (int, int) GetSize(this CellRange x)
    {
      return (1 + x.LastRowIndex - x.FirstRowIndex, 1 + x.LastColumnIndex - x.FirstColumnIndex);
    }
  }
}
