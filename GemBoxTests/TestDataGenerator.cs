using GemBox.Spreadsheet;
using ArrayEWE.Helpers;
using System.IO;

namespace GemBoxTests
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
      WorksheetExtension.Initialize();
      Directory.CreateDirectory(TestDataPath);

      if (!File.Exists(BigTestPath))
        CreateBigTest();

      if (!File.Exists(BigTest1Path))
        CreateBigTest1();
    }

    public static (long bigTestMs, long bigTest1Ms) BenchmarkCreation()
    {
      WorksheetExtension.Initialize();
      File.Delete(BigTestPath);
      File.Delete(BigTest1Path);
      Directory.CreateDirectory(TestDataPath);

      var sw = System.Diagnostics.Stopwatch.StartNew();
      CreateBigTest();
      long bigTestMs = sw.ElapsedMilliseconds;

      sw.Restart();
      CreateBigTest1();
      long bigTest1Ms = sw.ElapsedMilliseconds;

      return (bigTestMs, bigTest1Ms);
    }

    private static void CreateBigTest()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Sheet1");

      for (int r = 0; r < 10000; r++) {
        for (int c = 0; c < 20; c++) {
          ws.Cells[r, c].Value = $"Cell: Row-{r};CellNo:{c}";
        }
      }

      ef.Save(BigTestPath);
    }

    private static void CreateBigTest1()
    {
      var ef = new ExcelFile();
      var ws = ef.Worksheets.Add("Sheet1");

      // Rows 9-9999: cells at cols 5-19 with data
      // (No blank boundary row needed: GetUsedCellRange covers the actual data range,
      // and GetRange pads the result array to the requested top-left origin.)
      for (int r = 9; r < 10000; r++) {
        for (int c = 5; c < 20; c++) {
          ws.Cells[r, c].Value = $"Cell: Row-{r};CellNo:{c}";
        }
      }

      ef.Save(BigTest1Path);
    }
  }
}
