/*using System;
using System.Data;
using System.Data.SqlClient;

namespace TP.Object.BackEnd
{
  internal struct QueryParameter
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

  internal static class Connection
  {
    private delegate object queryRunner();

    private const string CONN_STRING = "Data Source=toolsplus06;Initial Catalog=toolsplus;Integrated Security=SSPI";

    private static object runQuery(queryRunner qr)
    {
      int retryCount = 0;
      while (retryCount < 5)
      {
        try
        {
          object retval = qr();
          return retval;
        }
        catch (SqlException sqlex)
        {
          switch (sqlex.Number) {
            case 10054:
              // A transport-level error has occurred when sending the request to the server. (provider: TCP Provider, error: 0 - An existing connection was forcibly closed by the remote host.)
              retryCount++;
              break;
            // should add other handleable errors here, PK violations, etc
            default:
              throw;
          }
        }
        catch (Exception) {
          throw;
        }
      }
      throw new Exception("SQL connection retry count exceeded");
    }

    public static DataTable Retrieve(string sql, params QueryParameter[] parameters) { return Retrieve(null, sql, parameters); }
    public static DataTable Retrieve(SqlTransaction txn, string sql, params QueryParameter[] parameters)
    {
      object retval = runQuery(delegate()
      {
        SqlConnection conn;
        SqlCommand sth;
        if (txn == null)
        {
          conn = new SqlConnection(CONN_STRING);
          conn.Open();
          sth = new SqlCommand(sql, conn);
        }
        else
        {
          conn = txn.Connection;
          sth = new SqlCommand(sql, txn.Connection, txn);
        }
        foreach (QueryParameter qp in parameters)
        {
          sth.Parameters.Add(qp.Name, qp.Type);
          sth.Parameters[qp.Name].Value = qp.Value;
        }

        if (true)
        {
          System.Diagnostics.Debug.WriteLine("    " + sql);
          foreach (var qp in parameters)
          {
            System.Diagnostics.Debug.WriteLine("      {0} = {1}", qp.Name, qp.Value);
          }
        }

        var da = new SqlDataAdapter(sth);
        var ds = new DataSet("t");
        da.Fill(ds, "t");
        if (txn == null)
        {
          conn.Close();
        }
        return ds.Tables["t"];
      });
      return (DataTable)retval;
    }

    public static int Insert(string sql, params QueryParameter[] parameters) { return Insert(null, sql, parameters); }
    public static int Insert(SqlTransaction txn, string sql, params QueryParameter[] parameters)
    {
      object retval = runQuery(delegate()
      {
        SqlConnection conn;
        SqlCommand sth;
        if (txn == null)
        {
          conn = new SqlConnection(CONN_STRING);
          conn.Open();
          sth = new SqlCommand(sql + "; SELECT CAST(scope_identity() AS int);", conn);
        }
        else
        {
          conn = txn.Connection;
          sth = new SqlCommand(sql + "; SELECT CAST(scope_identity() AS int);", txn.Connection, txn);
        }
        foreach (QueryParameter qp in parameters)
        {
          sth.Parameters.Add(qp.Name, qp.Type);
          sth.Parameters[qp.Name].Value = qp.Value ?? DBNull.Value;
        }

        var temp = sth.ExecuteScalar();
        if (txn == null) {
          conn.Close();
        }
        return (int)temp;
      });
      return (int)retval;
    }

    public static int Execute(string sql, params QueryParameter[] parameters) { return Execute(null, sql, parameters); }
    public static int Execute(SqlTransaction txn, string sql, params QueryParameter[] parameters)
    {
      object retval = runQuery(delegate()
      {
        SqlConnection conn;
        SqlCommand sth;
        if (txn == null)
        {
          conn = new SqlConnection(CONN_STRING);
          conn.Open();
          sth = new SqlCommand(sql, conn);
        }
        else
        {
          conn = txn.Connection;
          sth = new SqlCommand(sql, txn.Connection, txn);
        }
        foreach (QueryParameter qp in parameters)
        {
          sth.Parameters.Add(qp.Name, qp.Type);
          sth.Parameters[qp.Name].Value = qp.Value ?? DBNull.Value;
        }

        int rows = sth.ExecuteNonQuery();
        if (txn == null)
        {
          conn.Close();
        }
        return rows;
      });
      return (int)retval;
    }


    #region Transactions

    public static SqlTransaction BeginTransaction()
    {
      object retval = runQuery(delegate()
      {
        var conn = new SqlConnection(CONN_STRING);
        conn.Open();
        var txn = conn.BeginTransaction();
        return txn;
      });
      return (SqlTransaction)retval;
    }

    public static bool CommitTransaction(SqlTransaction txn)
    {
      object retval = runQuery(delegate()
      {
        var conn = txn.Connection;
        txn.Commit();
        conn.Close();
        return true;
      });
      return (bool)retval;
    }

    public static bool RollbackTransaction(SqlTransaction txn)
    {
      object retval = runQuery(delegate()
      {
        var conn = txn.Connection;
        txn.Rollback();
        conn.Close();
        return true;
      });
      return (bool)retval;
    }

    #endregion
  }
}
*/