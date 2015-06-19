using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "SalesOrderTask", Namespace = "TP.Object")]
  public class DbSalesOrderTask : TP.Object.SalesOrderTask, IWarehouseTask
  {
    #region Cache

    private static AutoCache2Key<DbSalesOrderTask, int, string> cache = new AutoCache2Key<DbSalesOrderTask, int, string>
    {
      Constructor = _ => new DbSalesOrderTask(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "HeaderID"),
      PrimaryKey = _ => _.SalesOrderTaskID,
      AlternateKey = _ => _.SalesOrderNo,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbSalesOrderTask() { }

    public DbSalesOrderTask(DataRow r)
    {
      this.SalesOrderTaskID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "HeaderID");
      this.SalesOrderNo = (string)r["SalesOrderNo"];
      this.WarehouseID = (int)r["WarehouseID"];
      this.CurrentStatus = (int)r["CurrentStatus"];
      this.CurrentHoldReason = (int)r["CurrentHoldReason"];
      this.ShipVia = (string)r["ShipVia"];
      this.ShipState = this.nullableString(r["ShipState"]);
      this.ShipCountry = this.nullableString(r["ShipCountry"]);
      this.MasCreatedByUser = this.nullableString(r["MasCreatedByUser"]);
      this.TimeInserted = (DateTime)r["TimeInserted"];
    }

    #endregion


    #region Getters

    public static DbSalesOrderTask GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbSalesOrderTask GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, SalesOrderNo, WarehouseID, CurrentStatus, CurrentHoldReason, ShipVia, ShipState, ShipCountry, MasCreatedByUser, TimeInserted FROM WhseTaskSalesOrderHeader WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbSalesOrderTask GetBySalesOrderNo(string salesOrderNo, string[] preloads = null) { return GetBySalesOrderNo(null, salesOrderNo, preloads); }
    internal static DbSalesOrderTask GetBySalesOrderNo(System.Data.SqlClient.SqlTransaction txn, string salesOrderNo, string[] preloads = null)
    {
      var ret = cache.FetchByAlternateKey(salesOrderNo, () => Connection.Retrieve(txn, "SELECT ID, SalesOrderNo, WarehouseID, CurrentStatus, CurrentHoldReason, ShipVia, ShipState, ShipCountry, MasCreatedByUser, TimeInserted FROM WhseTaskSalesOrderHeader WHERE SalesOrderNo=@id", new QueryParameter("@id", SqlDbType.VarChar, salesOrderNo)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbSalesOrderTask fillPreloads(DbSalesOrderTask t, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (t == null)
      {
        return t;
      }
      var retval = (DbSalesOrderTask)t.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("lines"))
        {
          retval.FillLines(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbSalesOrderTask> fillPreloads(IEnumerable<DbSalesOrderTask> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbSalesOrderTask has-many DbSalesOrderTaskLine
    public override bool FillLines() { return FillLines(null); }
    internal bool FillLines(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._lines == null)
      {
        this._lines = DbSalesOrderTaskLine.GetByTaskID(txn, this.SalesOrderTaskID);
      }
      return true;
    }

    // DbSalesOrderTask has-one DbSalesChannel
    public override bool FillSalesChannel() { return FillSalesChannel(null); }
    internal bool FillSalesChannel(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._saleschannel == null)
      {
        throw new NotImplementedException("TODO: DbSalesChannel");
        //this._saleschannel = DbSalesChannel.GetBySalesOrderNo(txn, this.SalesOrderNo);
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
    public int TransactionTypeID { get { return 3; } }
    public int TransactionID { get { return this.SalesOrderTaskID; } }
    public string TransactionReference { get { return this.SalesOrderNo; } }

  }
}
