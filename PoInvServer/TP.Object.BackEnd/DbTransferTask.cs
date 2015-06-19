using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "TransferTask", Namespace = "TP.Object")]
  public class DbTransferTask : TP.Object.TransferTask, IWarehouseTask
  {
    #region Cache

    private static AutoCache2Key<DbTransferTask, int, string> cache = new AutoCache2Key<DbTransferTask, int, string>
    {
      Constructor = _ => new DbTransferTask(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "HeaderID"),
      PrimaryKey = _ => _.TransferTaskID,
      AlternateKey = _ => _.TransferNo,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbTransferTask() { }

    public DbTransferTask(DataRow r)
    {
      this.TransferTaskID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "HeaderID");
      this.InitialDate = (DateTime)r["InitialDate"];
      this.SequenceNo = (int)r["SequenceNumber"];
      this.TransferNo = (string)r["TransferNumber"];
      this.TransferSubtype = (int)r["TransferSubtype"];
      this.WarehouseID = (int)r["WarehouseID"];
      this.DestinationWarehouseID = (int)r["DestinationWarehouseID"];
      this.IsComplete = (bool)r["IsComplete"];
      this.Description = this.nullableString(r["Description"]);
      this.TimeInserted = (DateTime)r["TimeInserted"];
      this.CreatedByUserID = (int)r["CreatedByUserID"];
    }

    #endregion


    #region Getters

    public static DbTransferTask GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbTransferTask GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, InitialDate, SequenceNumber, TransferNumber, TransferSubtype, WarehouseID, DestinationWarehouseID, IsComplete, Description, TimeInserted, CreatedByUserID FROM WhseTaskTransferHeader WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbTransferTask GetByTransferNo(string transferNo, string[] preloads = null) { return GetByTransferNo(null, transferNo, preloads); }
    internal static DbTransferTask GetByTransferNo(System.Data.SqlClient.SqlTransaction txn, string transferNo, string[] preloads = null)
    {
      var ret = cache.FetchByAlternateKey(transferNo, () => Connection.Retrieve(txn, "SELECT ID, InitialDate, SequenceNumber, TransferNumber, TransferSubtype, WarehouseID, DestinationWarehouseID, IsComplete, Description, TimeInserted, CreatedByUserID FROM WhseTaskTransferHeader WHERE TransferNumber=@id", new QueryParameter("@id", SqlDbType.VarChar, transferNo)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbTransferTask fillPreloads(DbTransferTask t, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (t == null)
      {
        return t;
      }
      var retval = (DbTransferTask)t.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("lines"))
        {
          retval.FillLines(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbTransferTask> fillPreloads(IEnumerable<DbTransferTask> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbTransferTask has-many DbTransferTaskLine
    public override bool FillLines() { return FillLines(null); }
    internal bool FillLines(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._lines == null)
      {
        this._lines = DbTransferTaskLine.GetByTaskID(txn, this.TransferTaskID);
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
    public int TransactionTypeID { get { return 7; } }
    public int TransactionID { get { return this.TransferTaskID; } }
    public string TransactionReference { get { return this.TransferNo; } }

  }
}
