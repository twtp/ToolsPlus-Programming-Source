using System;
using System.Collections.Generic;
using System.Text;

using System.Data;
using System.Data.SqlClient;

namespace TP.Database {
  public struct QueryParameter {
    public string Name;
    public SqlDbType Type;
    public object Value;

    [System.Diagnostics.DebuggerStepThroughAttribute()]
    public QueryParameter(string name, SqlDbType type, object value) {
      this.Name = name;
      this.Type = type;
      this.Value = value;

      /*if (this.Type == SqlDbType.VarChar && this.Value == null) {
        this.Value = System.DBNull.Value;
      }*/
    }
  }
}
