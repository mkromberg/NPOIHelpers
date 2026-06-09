using System;

namespace ArrayEWE.Helpers
{
  public class RangeResult
  {
    public string[] Names { get; internal set; } = Array.Empty<string>();
    public object[] Values { get; internal set; } = Array.Empty<object>();

    public object this[string name]
    {
      get
      {
        int i = Array.FindIndex(Names, n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));
        return i >= 0 ? Values[i] : null;
      }
    }
  }
}
