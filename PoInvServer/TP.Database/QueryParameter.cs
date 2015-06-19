using System;
using System.Data;
using System.Data.SqlClient;

namespace TP.Database
{
  public struct QueryParameter
  {
    public string Name { get; set; }
    public SqlDbType Type { get; set; }
    public object Value { get; set; }

    [System.Diagnostics.DebuggerStepThrough()]
    public QueryParameter(string name, SqlDbType type, object value) : this()
    {
      this.Name = name;
      this.Type = type;
      this.Value = value;
    }
  }
}
