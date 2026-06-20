using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

    private static readonly string[] AllRangeProperties = { "Values", "Error", "NumberFormat", "HasFormula", "Formula" };

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, int top = 1, int left = 1, int rows = int.MaxValue, int cols = int.MaxValue)
    {
      return GetRange(sheet, AllRangeProperties, top, left, rows, cols);
    }

    public static RangeResult GetRange(this SS.UserModel.ISheet sheet, string[] names, int top = 1, int left = 1, int rows = int.MaxValue, int cols = int.MaxValue)
    {
      var valid = new HashSet<string> { "values", "error", "numberformat", "hasformula", "formula" };
      foreach (var name in names)
        if (!valid.Contains(name.ToLowerInvariant()))
          throw new ArgumentException($"Unknown property: '{name}'");

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
      if (rng == null)
        return new RangeResult();

      var rrng = u.Intersect(wanted);
      rrng.FirstRow = top - 1;
      rrng.FirstColumn = left - 1;

      var (rrows, rcols) = rrng.GetSize();

      var nameSet = new HashSet<string>(names.Select(n => n.ToLowerInvariant()));
      var r  = nameSet.Contains("values")       ? new object[rrows, rcols] : null;
      var e  = nameSet.Contains("error")        ? new bool[rrows, rcols]   : null;
      var nf = nameSet.Contains("numberformat") ? new string[rrows, rcols] : null;
      var hf = nameSet.Contains("hasformula")   ? new bool[rrows, rcols]   : null;
      var fm = nameSet.Contains("formula")      ? new string[rrows, rcols] : null;

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
          if (x < row.FirstCellNum || x >= row.LastCellNum)
            continue;

          var cell = row.GetCell(x);
          if (cell == null)
            continue;

          var rc = c + oc;

          switch (cell.CellType) {
            case SS.UserModel.CellType.Numeric:
              if (r != null) r[rr, rc] = cell.NumericCellValue;
              break;
            case SS.UserModel.CellType.Boolean:
              if (r != null) r[rr, rc] = cell.BooleanCellValue;
              break;
            case SS.UserModel.CellType.String:
              if (r != null) r[rr, rc] = cell.StringCellValue;
              break;
            case SS.UserModel.CellType.Error:
              if (e != null) e[rr, rc] = true;
              break;
            case SS.UserModel.CellType.Formula:
              if (hf != null) hf[rr, rc] = true;
              if (fm != null) fm[rr, rc] = cell.CellFormula;
              switch (cell.CachedFormulaResultType) {
                case SS.UserModel.CellType.String:
                  if (r != null) r[rr, rc] = cell.StringCellValue;
                  break;
                case SS.UserModel.CellType.Boolean:
                  if (r != null) r[rr, rc] = cell.BooleanCellValue;
                  break;
                case SS.UserModel.CellType.Numeric:
                  if (r != null) r[rr, rc] = cell.NumericCellValue;
                  break;
                case SS.UserModel.CellType.Error:
                  if (e != null) e[rr, rc] = true;
                  break;
              }
              break;
          }

          if (nf != null) {
            var fmtStr = cell.CellStyle?.GetDataFormatString();
            nf[rr, rc] = (fmtStr == null || fmtStr == "General") ? null : fmtStr;
          }
        }
      }

      var resultNames = new List<string>();
      var resultValues = new List<object>();
      foreach (var name in names) {
        object data;
        switch (name.ToLowerInvariant()) {
          case "values":       data = r;  break;
          case "error":        data = e;  break;
          case "numberformat": data = nf; break;
          case "hasformula":   data = hf; break;
          case "formula":      data = fm; break;
          default:             data = null; break;
        }
        if (data != null) { resultNames.Add(name); resultValues.Add(data); }
      }

      return new RangeResult { Names = resultNames.ToArray(), Values = resultValues.ToArray() };
    }

    public static RangeResult GetUsedRange(this SS.UserModel.ISheet sheet)
    {
      if (sheet.PhysicalNumberOfRows == 0)
        return new RangeResult();
      var u = sheet.GetUsedRangeAddress();
      var (rows, cols) = u.GetSize();
      return sheet.GetRange(u.FirstRow + 1, u.FirstColumn + 1, rows, cols);
    }

    public static string[] AllSheets(SS.UserModel.IWorkbook wb)
    {
      var names = new string[wb.NumberOfSheets];
      for (int i = 0; i < wb.NumberOfSheets; i++)
        names[i] = wb.GetSheetName(i);
      return names;
    }

    public static SS.UserModel.ISheet AddSheet(SS.UserModel.IWorkbook wb, string name, bool first = false)
    {
      var sheet = wb.CreateSheet(name);
      if (first)
        wb.SetSheetOrder(name, 0);
      return sheet;
    }

    public static void DeleteSheet(SS.UserModel.IWorkbook wb, string name)
    {
      int idx = wb.GetSheetIndex(name);
      if (idx < 0)
        throw new ArgumentException($"Sheet '{name}' not found.", nameof(name));
      wb.RemoveSheetAt(idx);
    }

    public static SS.UserModel.ISheet GetSheet(SS.UserModel.IWorkbook wb, string name)
    {
      return wb.GetSheet(name);
    }

    public static SS.UserModel.IWorkbook Open(string fileName, string password = null)
    {
      if (string.IsNullOrEmpty(password))
        return SS.UserModel.WorkbookFactory.Create(fileName);

      var nfs = new NPOI.POIFS.FileSystem.NPOIFSFileSystem(new FileInfo(fileName));
      try
      {
        // Agile/Standard OLE encryption (.xlsx, and .xls saved by modern Excel
        // with "strong" encryption): EncryptionInfo lives as an OLE entry.
        var plain = new MemoryStream();
        using (var dec = NPOI.POIFS.FileSystem.DocumentFactoryHelper.GetDecryptedStream(nfs, password))
          dec.CopyTo(plain);
        plain.Position = 0;
        return SS.UserModel.WorkbookFactory.Create(plain);
      }
      catch (IOException)
      {
        // No EncryptionInfo OLE entry: record-level Biff8/CryptoAPI encryption
        // inside the HSSF stream. Supply the password via the thread-local key.
        NPOI.HSSF.Record.Crypto.Biff8EncryptionKey.CurrentUserPassword = password;
        try   { return SS.UserModel.WorkbookFactory.Create(fileName); }
        finally { NPOI.HSSF.Record.Crypto.Biff8EncryptionKey.CurrentUserPassword = null; }
      }
    }

    public static void Protect(SS.UserModel.IWorkbook wb, string password)
    {
      if (wb is NPOI.HSSF.UserModel.HSSFWorkbook hssf)
      {
        if (string.IsNullOrEmpty(password))
          hssf.UnwriteProtectWorkbook();
        else
          hssf.WriteProtectWorkbook(password, "");
      }
      else if (wb is NPOI.XSSF.UserModel.XSSFWorkbook xssf)
      {
        if (string.IsNullOrEmpty(password))
        {
          xssf.UnlockStructure();
        }
        else
        {
          xssf.LockStructure();
          var prot = xssf.GetCTWorkbook().workbookProtection;
          if (prot != null)
            prot.workbookPassword = ExcelPasswordHash(password);
        }
      }
    }

    public static void ProtectSheet(SS.UserModel.ISheet sheet, string password)
    {
      sheet.ProtectSheet(string.IsNullOrEmpty(password) ? null : password);
    }

    // Legacy XOR hash from ECMA-376 §4.3.1 / MS-OFFCRYPTO §2.3.7.4.
    // Produces a 2-byte big-endian result stored in the workbookPassword hexBinary attribute.
    private static byte[] ExcelPasswordHash(string password)
    {
      byte[] chars = System.Text.Encoding.GetEncoding(1252).GetBytes(password);
      int hash = 0;
      for (int i = chars.Length - 1; i >= 0; i--)
      {
        hash = ((hash >> 1) & 0x3FFF) | ((hash & 1) << 14);
        hash ^= chars[i];
        hash &= 0x7FFF;
      }
      hash = ((hash >> 1) & 0x3FFF) | ((hash & 1) << 14);
      hash ^= chars.Length;
      hash ^= 0xCE4B;
      return new byte[] { (byte)((hash >> 8) & 0xFF), (byte)(hash & 0xFF) };
    }

    public static SS.UserModel.IWorkbook New(string fileName)
    {
      var ext = Path.GetExtension(fileName).ToLowerInvariant();
      if (ext == ".xls")
        return new NPOI.HSSF.UserModel.HSSFWorkbook();
      return new NPOI.XSSF.UserModel.XSSFWorkbook();
    }

    public static void Save(SS.UserModel.IWorkbook wb, string fileName, string password = null)
    {
      if (string.IsNullOrEmpty(password))
      {
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
          wb.Write(fs);
        return;
      }

      var ext = Path.GetExtension(fileName).ToLowerInvariant();
      if (ext == ".xls")
      {
        // HSSF uses thread-local Biff8 RC4 encryption; set before Write, clear after.
        NPOI.HSSF.Record.Crypto.Biff8EncryptionKey.CurrentUserPassword = password;
        try
        {
          using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            wb.Write(fs);
        }
        finally { NPOI.HSSF.Record.Crypto.Biff8EncryptionKey.CurrentUserPassword = null; }
        return;
      }

      // XSSF (.xlsx): write into an Agile-encrypted OLE container.
      var encInfo = new NPOI.POIFS.Crypt.EncryptionInfo(NPOI.POIFS.Crypt.EncryptionMode.Agile);
      var enc = encInfo.Encryptor;
      enc.ConfirmPassword(password);
      var poiFs = new NPOI.POIFS.FileSystem.NPOIFSFileSystem();
      using (var dataStream = enc.GetDataStream(poiFs.Root))
        wb.Write(dataStream);
      using (var outFs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
        poiFs.WriteFileSystem(outFs);
    }

    public static void ClearSheet(this SS.UserModel.ISheet sheet)
    {
      for (int i = sheet.LastRowNum; i >= sheet.FirstRowNum; i--)
      {
        var row = sheet.GetRow(i);
        if (row != null)
          sheet.RemoveRow(row);
      }
    }

    public static void ClearHidden(this SS.UserModel.ISheet sheet)
    {
      for (int i = sheet.LastRowNum; i >= sheet.FirstRowNum; i--)
      {
        var row = sheet.GetRow(i);
        if (row != null && row.ZeroHeight)
          sheet.RemoveRow(row);
      }

      // Column-hidden state persists independently of cell data, so scan
      // even when the sheet is empty.  Use the data range when rows exist;
      // fall back to the full column space when the sheet has no rows.
      int lastCol = sheet.PhysicalNumberOfRows > 0
        ? sheet.GetUsedRangeAddress().LastColumn
        : (sheet is NPOI.XSSF.UserModel.XSSFSheet ? 16383 : 255);

      for (int c = 0; c <= lastCol; c++)
      {
        if (!sheet.IsColumnHidden(c)) continue;
        for (int ri = sheet.FirstRowNum; ri <= sheet.LastRowNum; ri++)
        {
          var row = sheet.GetRow(ri);
          row?.GetCell(c)?.SetBlank();
        }
        sheet.SetColumnHidden(c, false);
        sheet.SetColumnWidth(c, sheet.DefaultColumnWidth * 256);
      }
    }

    public static double[] GetColumnWidths(this SS.UserModel.ISheet sheet, int left = 1, int cols = int.MaxValue)
    {
      if (sheet.PhysicalNumberOfRows == 0 && cols == int.MaxValue)
        return Array.Empty<double>();

      int startCol = left - 1;
      int endCol = cols == int.MaxValue
        ? sheet.GetUsedRangeAddress().LastColumn
        : startCol + cols - 1;

      int count = endCol - startCol + 1;
      var widths = new double[count];
      for (int c = startCol; c <= endCol; c++)
        // GetColumnWidthInPixels returns character pixels (no margin).
        // +5 adds Excel's standard 5px per-column margin; ×0.75 converts 96dpi px → points.
        widths[c - startCol] = (sheet.GetColumnWidthInPixels(c) + 5.0) * 0.75;
      return widths;
    }

    public static void AutoFitColumns(this SS.UserModel.ISheet sheet, int left = 1, int cols = int.MaxValue)
    {
      if (sheet.PhysicalNumberOfRows == 0) return;
      var u = sheet.GetUsedRangeAddress();
      int startCol = left - 1;
      int endCol = cols == int.MaxValue ? u.LastColumn : Math.Min(startCol + cols - 1, u.LastColumn);
      for (int c = startCol; c <= endCol; c++)
        sheet.AutoSizeColumn(c);
    }

    public static void AppendRange(this SS.UserModel.ISheet sheet, object[,] data, int left = 1)
    {
      int top = sheet.PhysicalNumberOfRows == 0
        ? 1
        : sheet.GetUsedRangeAddress().LastRow + 2;
      sheet.PutRange("Values", data, top, left);
    }

    public static void PutRange(this SS.UserModel.ISheet sheet, string name, object[,] value, int top, int left)
    {
      switch (name.ToLowerInvariant())
      {
        case "values":
        {
          var data = value;
          int rows = data.GetLength(0), cols = data.GetLength(1);
          for (int r = 0; r < rows; r++) {
            var row = sheet.GetRow((top - 1) + r) ?? sheet.CreateRow((top - 1) + r);
            for (int c = 0; c < cols; c++) {
              var cell = row.GetCell((left - 1) + c) ?? row.CreateCell((left - 1) + c);
              switch (data[r, c]) {
                case null:       cell.SetBlank(); break;
                case string s:   cell.SetCellValue(s); break;
                case bool b:     cell.SetCellValue(b); break;
                case DateTime dt: cell.SetCellValue(dt); break;
                case double d:   cell.SetCellValue(d); break;
                default:         cell.SetCellValue(Convert.ToDouble(data[r, c])); break;
              }
            }
          }
          break;
        }
        case "numberformat":
        {
          int rows = value.GetLength(0), cols = value.GetLength(1);
          var wb = sheet.Workbook;
          var dataFormat = wb.CreateDataFormat();
          var styleCache = new Dictionary<string, SS.UserModel.ICellStyle>();
          for (int r = 0; r < rows; r++) {
            var row = sheet.GetRow((top - 1) + r) ?? sheet.CreateRow((top - 1) + r);
            for (int c = 0; c < cols; c++) {
              var fmt = value[r, c] as string;
              if (string.IsNullOrEmpty(fmt)) continue;
              var cell = row.GetCell((left - 1) + c) ?? row.CreateCell((left - 1) + c);
              if (!styleCache.TryGetValue(fmt, out var style)) {
                style = wb.CreateCellStyle();
                style.DataFormat = dataFormat.GetFormat(fmt);
                styleCache[fmt] = style;
              }
              cell.CellStyle = style;
            }
          }
          break;
        }
        default:
          throw new ArgumentException($"PutRange does not support property '{name}'.");
      }
    }
  }
}
