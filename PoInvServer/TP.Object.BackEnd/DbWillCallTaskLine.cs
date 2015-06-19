using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "WillCallTaskLine", Namespace = "TP.Object")]
  public class DbWillCallTaskLine : TP.Object.WillCallTaskLine, IWarehouseTaskLine
  {
    #region Cache

    private static AutoCache1Key1Index<DbWillCallTaskLine, int, int> cache = new AutoCache1Key1Index<DbWillCallTaskLine, int, int>
    {
      Constructor = _ => new DbWillCallTaskLine(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "LineID"),
      PrimaryKey = _ => _.TupleID,
      Index1 = _ => _.HeaderID,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbWillCallTaskLine() { }

    public DbWillCallTaskLine(DataRow r)
    {
      this.TupleID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "LineID");
      this.WillCallTaskID = (int)r["HeaderID"];
      this.ComponentID = (int)r["ComponentID"];
      this.QuantityRequired = (int)r["QuantityRequired"];
      this.QuantityPicked = (int)r["QuantityPicked"];
      this.QuantityPacked = (int)r["QuantityStaged"];
      this.QuantityFilled = (int)r["QuantityFilled"];
    }

    #endregion


    #region Getters

    public static DbWillCallTaskLine GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbWillCallTaskLine GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityRequired, QuantityPicked, QuantityStaged, QuantityFilled FROM WhseTaskWillCallLines WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbWillCallTaskLine> GetByTaskID(int id, string[] preloads = null) { return GetByTaskID(null, id, preloads); }
    internal static IEnumerable<DbWillCallTaskLine> GetByTaskID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityRequired, QuantityPicked, QuantityStaged, QuantityFilled FROM WhseTaskWillCallLines WHERE HeaderID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbWillCallTaskLine> GetPickingRequired(string[] preloads = null) { return GetPickingRequired(null, 5, preloads); }
    public static IEnumerable<DbWillCallTaskLine> GetPickingRequired(int warehouseID, string[] preloads = null) { return GetPickingRequired(null, warehouseID, preloads); }
    internal static IEnumerable<DbWillCallTaskLine> GetPickingRequired(System.Data.SqlClient.SqlTransaction txn, int warehouseID, string[] preloads = null)
    {
      // completely bypasses cache?
      DataTable dt = Connection.Retrieve(txn,
          "SELECT H.ID AS HeaderID, H.TransactionID, H.WarehouseID,  H.TimeInserted, H.CreatedByUserID, " +
          "       L.ID AS LineID, L.ComponentID, L.QuantityRequired, L.QuantityPicked, L.QuantityStaged, L.QuantityFilled " +
          "FROM WhseTaskWillCallHeader AS H INNER JOIN WhseTaskWillCallLines AS L ON H.ID=L.HeaderID " +
          "WHERE H.WarehouseID=@whse " +
          "  AND L.QuantityRequired>L.QuantityPicked " +
          "ORDER BY H.HeaderID, L.LineID", new QueryParameter("@whse", SqlDbType.Int, warehouseID));
      var temp = new Dictionary<int, DbWillCallTask>();
      var retval = new List<DbWillCallTaskLine>();
      foreach (DataRow r in dt.Rows)
      {
        int hid = r.Field<int>("HeaderID");
        if (!temp.ContainsKey(hid))
        {
          temp.Add(hid, new DbWillCallTask(r));
        }
        var l = new DbWillCallTaskLine(r);
        l._willcall = temp[hid];
        retval.Add(l);
      }

      return fillPreloads(retval, preloads, txn);
    }

    private static DbWillCallTaskLine fillPreloads(DbWillCallTaskLine l, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (l == null)
      {
        return l;
      }
      var retval = (DbWillCallTaskLine)l.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("component"))
        {
          retval.FillComponent(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbWillCallTaskLine> fillPreloads(IEnumerable<DbWillCallTaskLine> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbWillCallTaskLine has-one DbWillCallTask
    public override bool FillWillCallTask() { return FillWillCallTask(null); }
    internal bool FillWillCallTask(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._willcall == null)
      {
        this._willcall = DbWillCallTask.GetByID(txn, this.WillCallTaskID);
      }
      return true;
    }

    public override bool FillComponent() { return FillComponent(null); }
    internal bool FillComponent(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._component == null)
      {
        this._component = DbComponent.GetByID(txn, this.ComponentID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // create not allowed
    // update like normal not allowed, only from internal location movement
    // delete not allowed

    // should be internal, but needs public for interface
    public IWarehouseTaskLine UpdateQuantity(System.Data.SqlClient.SqlTransaction txn, int fromLocationID, int? toLocationID, int deltaQuantity)
    {
      DbInventoryLocation fromLoc = DbInventoryLocation.GetByID(txn, fromLocationID);
      DbInventoryLocation toLoc = toLocationID.HasValue ? DbInventoryLocation.GetByID(txn, toLocationID.Value) : null;
      this.FillWillCallTask(txn);
      int whse = this.WillCallTask.WarehouseID;

      List<string> updateFields = new List<string>();
      List<QueryParameter> updateParams = new List<QueryParameter>();

      if (fromLoc.WarehouseID == whse && fromLoc.IsTypeStandard &&                       // from a standard stock location
          toLoc != null && toLoc.WarehouseID == whse && toLoc.IsTypePrepack)             // to a tote/backorder/holding
      {                                                                                  // == pick
        updateFields.Add("QuantityPicked=QuantityPicked+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.IsTypePrepack &&                   // from a tote/backorder/holding
               toLoc != null && toLoc.WarehouseID == whse && toLoc.LocationTypeID == 9)  // to a will call staging type
      {                                                                                  // == pack
        updateFields.Add("QuantityStaged=QuantityStaged+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == whse && fromLoc.LocationTypeID == 9 &&             // from a will call staging type
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
      else if (fromLoc.WarehouseID == whse && fromLoc.LocationTypeID == 9 &&             // from a will call staging type
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
          "UPDATE WhseTaskWillCallLines SET " + string.Join(", ", updateFields.ToArray()) + " WHERE ID=@id",
          updateParams.ToArray()
        );

      cache.Clear(this.TupleID, this.HeaderID);
      var newObj = GetByID(txn, this.TupleID);
      if (newObj.QuantityPicked < 0) { throw new ArgumentException("cannot un-pick this many"); }
      if (newObj.QuantityPacked < 0) { throw new ArgumentException("cannot un-pack this many"); }
      if (newObj.QuantityFilled < 0) { throw new ArgumentException("cannot un-ship this many"); } // unship doesn't even exist, but here for completeness
      if (newObj.QuantityPicked > newObj.QuantityRequired) { throw new ArgumentException("cannot pick this many"); }
      if (newObj.QuantityPacked > newObj.QuantityPicked) { throw new ArgumentException("cannot pack this many"); }
      if (newObj.QuantityFilled > newObj.QuantityPacked) { throw new ArgumentException("cannot fill this many"); }

      return newObj;
    }

    #endregion


    #region Extra Validation

    // none

    #endregion


    // IWarehouseTaskLine
    public int HeaderID { get { return this.WillCallTaskID; } }
    public int LineID { get { return this.TupleID; } }
    public IWarehouseTask Header(System.Data.SqlClient.SqlTransaction txn)
    {
      this.FillWillCallTask(txn);
      return (DbWillCallTask)this.WillCallTask;
    }
    public bool HasChangedFromDbCurrent(System.Data.SqlClient.SqlTransaction txn)
    {
      var orig = GetByID(txn, this.LineID);
      return this.PropertiesModified(orig).Count > 0;
    }
  }
}
