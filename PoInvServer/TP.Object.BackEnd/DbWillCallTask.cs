using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "WillCallTask", Namespace = "TP.Object")]
  public class DbWillCallTask : TP.Object.WillCallTask, IWarehouseTask
  {
    #region Cache

    private static AutoCache2Key<DbWillCallTask, int, string> cache = new AutoCache2Key<DbWillCallTask, int, string>
    {
      Constructor = _ => new DbWillCallTask(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "HeaderID"),
      PrimaryKey = _ => _.WillCallTaskID,
      AlternateKey = _ => _.MASTransactionID,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbWillCallTask() { }

    public DbWillCallTask(DataRow r)
    {
      this.WillCallTaskID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "HeaderID");
      this.MASTransactionID = (string)r["TransactionID"];
      this.WarehouseID = (int)r["WarehouseID"];
      this.CreatedByUserID = (int)r["CreatedByUserID"];
      this.TimeInserted = (DateTime)r["TimeInserted"];
    }

    #endregion


    #region Getters

    public static DbWillCallTask GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbWillCallTask GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, TransactionID, WarehouseID, TimeInserted, CreatedByUserID FROM WhseTaskWillCallHeader WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbWillCallTask GetByMASTransactionNo(string masTransactionNo, string[] preloads = null) { return GetByMASTransactionNo(null, masTransactionNo, preloads); }
    internal static DbWillCallTask GetByMASTransactionNo(System.Data.SqlClient.SqlTransaction txn, string masTransactionNo, string[] preloads = null)
    {
      var ret = cache.FetchByAlternateKey(masTransactionNo, () => Connection.Retrieve(txn, "SELECT ID, TransactionID, WarehouseID, TimeInserted, CreatedByUserID FROM WhseTaskWillCallHeader WHERE TransactionID=@id", new QueryParameter("@id", SqlDbType.VarChar, masTransactionNo)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbWillCallTask fillPreloads(DbWillCallTask t, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (t == null)
      {
        return t;
      }
      var retval = (DbWillCallTask)t.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("lines"))
        {
          retval.FillLines(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbWillCallTask> fillPreloads(IEnumerable<DbWillCallTask> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbWillCallTask has-many DbWillCallTaskLine
    public override bool FillLines() { return FillLines(null); }
    internal bool FillLines(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._lines == null)
      {
        this._lines = DbWillCallTaskLine.GetByTaskID(txn, this.WillCallTaskID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // create not allowed
    // update not allowed
    // delete not allowed

    #endregion


    #region Extra Validation

    // none

    #endregion


    // IWarehouseTask
    public int TransactionTypeID { get { return 4; } }
    public int TransactionID { get { return this.WillCallTaskID; } }
    public string TransactionReference { get { return this.MASTransactionID; } }

  }
}
