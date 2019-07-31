using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SCIta.Helpers
{
  public static class CellRangeAddressExtension
  {
    public static SS.Util.CellRangeAddress Intersect(this SS.Util.CellRangeAddress x, SS.Util.CellRangeAddress y)
    {
      (int,int) FirstLast(int xfirst,int xlast, int yfirst, int ylast)
      {
        int first = Math.Max(xfirst, yfirst);
        int last = Math.Min(xlast, ylast);
        
        return (first > last) ? (0,-1) : (first, last);
      }

      var (firstrow, lastrow) = FirstLast(x.FirstRow, x.LastRow, y.FirstRow, y.LastRow);
      var (firstcol, lastcol) = FirstLast(x.FirstColumn, x.LastColumn, y.FirstColumn, y.LastColumn);
      if(firstrow<=lastrow && firstcol<=lastcol)
        return new SS.Util.CellRangeAddress(firstrow, lastrow, firstcol, lastcol);

      return null;
    }

    public static (int, int) GetSize(this SS.Util.CellRangeAddress x)
    {
      return (1+x.LastRow - x.FirstRow, 1+x.LastColumn - x.FirstColumn);
    }

  }
}



