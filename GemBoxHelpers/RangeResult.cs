namespace ArrayEWE.Helpers
{
  public class RangeResult
  {
    public object[,] Values { get; internal set; } = new object[0, 0];
    public bool[,] Error { get; internal set; } = new bool[0, 0];
  }
}
