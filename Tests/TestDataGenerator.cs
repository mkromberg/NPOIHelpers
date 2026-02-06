using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace Tests
{
  public static class TestDataGenerator
  {
    public static readonly string TestDataPath = Path.Combine(
      Path.GetDirectoryName(typeof(TestDataGenerator).Assembly.Location),
      "TestData");

    public static string BigTestPath => Path.Combine(TestDataPath, "bigtest.xlsx");
    public static string BigTest1Path => Path.Combine(TestDataPath, "bigtest1.xlsx");

    public static void EnsureTestDataExists()
    {
      Directory.CreateDirectory(TestDataPath);

      if (!File.Exists(BigTestPath))
        CreateBigTest();

      if (!File.Exists(BigTest1Path))
        CreateBigTest1();
    }

    private static void CreateBigTest()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Sheet1");

      for (int r = 0; r < 10000; r++)
      {
        var row = sheet.CreateRow(r);
        for (int c = 0; c < 20; c++)
        {
          var cell = row.CreateCell(c);
          cell.SetCellValue($"Cell: Row-{r};CellNo:{c}");
        }
      }

      using (var fs = new FileStream(BigTestPath, FileMode.Create, FileAccess.Write))
      {
        wb.Write(fs);
      }
      wb.Close();
    }

    private static void CreateBigTest1()
    {
      var wb = new XSSFWorkbook();
      var sheet = wb.CreateSheet("Sheet1");

      // Row 0: blank cells at cols 0 and 19 to establish column range
      var row0 = sheet.CreateRow(0);
      row0.CreateCell(0).SetBlank();
      row0.CreateCell(19).SetBlank();

      // Rows 1-8: don't create (sparse)

      // Rows 9-9999: cells at cols 5-19 with data
      for (int r = 9; r < 10000; r++)
      {
        var row = sheet.CreateRow(r);
        for (int c = 5; c < 20; c++)
        {
          var cell = row.CreateCell(c);
          cell.SetCellValue($"Cell: Row-{r};CellNo:{c}");
        }
      }

      using (var fs = new FileStream(BigTest1Path, FileMode.Create, FileAccess.Write))
      {
        wb.Write(fs);
      }
      wb.Close();
    }
  }
}
