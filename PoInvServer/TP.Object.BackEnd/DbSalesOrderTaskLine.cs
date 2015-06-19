using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "SalesOrderTaskLine", Namespace = "TP.Object")]
  public class DbSalesOrderTaskLine : TP.Object.SalesOrderTaskLine, IWarehouseTaskLine
  {
    #region Cache

    private static AutoCache1Key1Index<DbSalesOrderTaskLine, int, int> cache = new AutoCache1Key1Index<DbSalesOrderTaskLine, int, int>
    {
      Constructor = _ => new DbSalesOrderTaskLine(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "LineID"),
      PrimaryKey = _ => _.TupleID,
      Index1 = _ => _.HeaderID,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbSalesOrderTaskLine() { }

    public DbSalesOrderTaskLine(DataRow r)
    {
      this.TupleID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "LineID");
      this.SalesOrderTaskID = (int)r["HeaderID"];
      this.ComponentID = (int)r["ComponentID"];
      this.QuantityOrdered = (int)r["QuantityOrdered"];
      this.QuantityBackordered = (int)r["QuantityBackordered"];
      this.QuantityPicked = (int)r["QuantityPicked"];
      this.QuantityPacked = (int)r["QuantityStaged"];
      this.QuantityShipped = (int)r["QuantityShipped"];
    }

    #endregion


    #region Getters

    public static DbSalesOrderTaskLine GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbSalesOrderTaskLine GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityOrdered, QuantityBackordered, QuantityPicked, QuantityStaged, QuantityShipped FROM WhseTaskSalesOrderLines WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbSalesOrderTaskLine> GetByTaskID(int id, string[] preloads = null) { return GetByTaskID(null, id, preloads); }
    internal static IEnumerable<DbSalesOrderTaskLine> GetByTaskID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityOrdered, QuantityBackordered, QuantityPicked, QuantityStaged, QuantityShipped FROM WhseTaskSalesOrderLines WHERE HeaderID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbSalesOrderTaskLine> GetPickingRequired(string[] preloads = null) { return GetPickingRequired(null, 5, preloads); }
    public static IEnumerable<DbSalesOrderTaskLine> GetPickingRequired(int warehouseID, string[] preloads = null) { return GetPickingRequired(null, warehouseID, preloads); }
    internal static IEnumerable<DbSalesOrderTaskLine> GetPickingRequired(System.Data.SqlClient.SqlTransaction txn, int warehouseID, string[] preloads = null)
    {
      // completely bypasses cache?
      DataTable dt = Connection.Retrieve(txn,
          "SELECT H.ID AS HeaderID, H.SalesOrderNo, H.WarehouseID, H.CurrentStatus, H.CurrentHoldReason, H.ShipVia, H.ShipState, H.ShipCountry, H.MasCreatedByUser, H.TimeInserted, " +
          "       L.ID AS LineID, L.ComponentID, L.QuantityOrdered, L.QuantityBackordered, L.QuantityPicked, L.QuantityStaged, L.QuantityShipped " +
          "FROM WhseTaskSalesOrderHeader AS H INNER JOIN WhseTaskSalesOrderLines AS L ON H.ID=L.HeaderID " +
          "WHERE H.WarehouseID=@whse " +
          "  AND H.CurrentStatus BETWEEN 3 AND 5 " + // 3 = complete, 4 = partial ship, 5 = partial hold
          "  AND H.ShipVia<>'DROP SHIP' " +
          "  AND H.ShipVia<>'PICK UP' " +
          "  AND L.QuantityOrdered-L.QuantityBackordered-L.QuantityPicked>0 " +
          "ORDER BY H.HeaderID, L.LineID", new QueryParameter("@whse", SqlDbType.Int, warehouseID));
      var temp = new Dictionary<int, DbSalesOrderTask>();
      var retval = new List<DbSalesOrderTaskLine>();
      foreach (DataRow r in dt.Rows)
      {
        int hid = r.Field<int>("HeaderID");
        if (!temp.ContainsKey(hid))
        {
          temp.Add(hid, new DbSalesOrderTask(r));
        }
        var l = new DbSalesOrderTaskLine(r);
        l._salesorder = temp[hid];
        retval.Add(l);
      }

      return fillPreloads(retval, preloads, txn);
    }

    private static DbSalesOrderTaskLine fillPreloads(DbSalesOrderTaskLine l, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (l == null)
      {
        return l;
      }
      var retval = (DbSalesOrderTaskLine)l.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("component"))
        {
          retval.FillComponent(txn, fillLocations: preloads.Contains("locations"));
        }
      }
      return retval;
    }

    private static IEnumerable<DbSalesOrderTaskLine> fillPreloads(IEnumerable<DbSalesOrderTaskLine> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbSalesOrderTaskLine has-one DbSalesOrderTask
    public override bool FillSalesOrderTask() { return FillSalesOrderTask(null); }
    internal bool FillSalesOrderTask(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._salesorder == null)
      {
        this._salesorder = DbSalesOrderTask.GetByID(txn, this.SalesOrderTaskID);
      }
      return true;
    }

    public override bool FillComponent() { return FillComponent(null); }
    internal bool FillComponent(System.Data.SqlClient.SqlTransaction txn, bool fillLocations = false)
    {
      if (this._component == null)
      {
        var c = DbComponent.GetByID(txn, this.ComponentID);
        if (fillLocations) {
          c.FillLocations(txn, fillLocationObjects: true);
        }
        this._component = c;
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // create not allowed
    // update like normal not allowed, only from internal location movement
    // delete not allowed

    // should be internal, but needs public for interface
    public IWarehouseTaskLine UpdateQuantity(System.Data.SqlClient.SqlTransaction txn, int fromLocationID, int? toLocationID, int deltaQuantity) {
      DbInventoryLocation fromLoc = DbInventoryLocation.GetByID(txn, fromLocationID);
      DbInventoryLocation toLoc = toLocationID.HasValue ? DbInventoryLocation.GetByID(txn, toLocationID.Value) : null;
      this.FillSalesOrderTask(txn);
      int whse = this.SalesOrderTask.WarehouseID;

      List<string> updateFields = new List<string>();
      List<QueryParameter> updateParams = new List<QueryParameter>();

      if (fromLoc.WarehouseID == whse && fromLoc.IsTypeStandard &&                       // from a standard stock location
          toLoc != null && toLoc.WarehouseID == whse && toLoc.IsTypePrepack)             // to a tote/backorder/holding
      {                                                                                  // == pick
        updateFields.Add("QuantityPicked=QuantityPicked+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.IsTypePrepack &&                   // from a tote/backorder/holding
               toLoc != null && toLoc.WarehouseID == whse && toLoc.LocationTypeID == 8)  // to a shipping staging type
      {                                                                                  // == pack
        updateFields.Add("QuantityStaged=QuantityStaged+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.LocationTypeID == 8 &&             // from a shipping staging type
               toLoc == null)                                                            // to nowhere
      {                                                                                  // == ship
        updateFields.Add("QuantityShipped=QuantityShipped+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.IsTypePrepack &&                   // from a tote/backorder/holding
               toLoc != null && toLoc.WarehouseID == whse && toLoc.LocationTypeID == 14) // to a restocking staging type
      {                                                                                  // == reversing a pick
        updateFields.Add("QuantityPicked=QuantityPicked-@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.LocationTypeID == 8 &&             // from a shipping staging type
               toLoc != null && toLoc.WarehouseID == whse && toLoc.LocationTypeID == 14) // to a restocking staging type
      {                                                                                  // == reversing a pack && reversing a pick
        updateFields.Add("QuantityStaged=QuantityStaged-@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
        updateFields.Add("QuantityPicked=QuantityPicked-@q2");
        updateParams.Add(new QueryParameter("@q2", SqlDbType.Int, deltaQuantity));
      }
      else
      {
        // ERROR: what should this be updating?
        System.Diagnostics.Debug.WriteLine("ERROR");
        throw new ArgumentException("can't update task line between these locations");
      }

      updateParams.Add(new QueryParameter("@id", SqlDbType.Int, this.TupleID));
      Connection.Execute(txn,
          "UPDATE WhseTaskSalesOrderLines SET " + string.Join(", ", updateFields.ToArray()) + " WHERE ID=@id",
          updateParams.ToArray()
        );

      cache.Clear(this.TupleID, this.HeaderID);
      var newObj = GetByID(txn, this.TupleID);
      if (newObj.QuantityPicked < 0) { throw new ArgumentException("cannot un-pick this many"); }
      if (newObj.QuantityPacked < 0) { throw new ArgumentException("cannot un-pack this many"); }
      if (newObj.QuantityShipped < 0) { throw new ArgumentException("cannot un-ship this many"); } // unship doesn't even exist, but here for completeness
      if (newObj.QuantityPicked > newObj.QuantityOrdered) { throw new ArgumentException("cannot pick this many"); }
      if (newObj.QuantityPacked > newObj.QuantityPicked) { throw new ArgumentException("cannot pack this many"); }
      if (newObj.QuantityShipped > newObj.QuantityPacked) { throw new ArgumentException("cannot ship this many"); }

      return newObj;
    }

    #endregion


    #region Extra Validation

    // none

    #endregion


    // IWarehouseTaskLine
    public int HeaderID { get { return this.SalesOrderTaskID; } }
    public int LineID { get { return this.TupleID; } }
    public IWarehouseTask Header(System.Data.SqlClient.SqlTransaction txn) {
      this.FillSalesOrderTask(txn);
      return (DbSalesOrderTask)this.SalesOrderTask;
    }
    public bool HasChangedFromDbCurrent(System.Data.SqlClient.SqlTransaction txn)
    {
      var orig = GetByID(txn, this.LineID);
      return this.PropertiesModified(orig).Count > 0;
    }



    public int QuantityLeftToPick { get { return this.QuantityOrdered - this.QuantityBackordered - this.QuantityPicked; } }
  }
}
