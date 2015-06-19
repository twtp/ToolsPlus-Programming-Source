using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "TransferTaskLine", Namespace = "TP.Object")]
  public class DbTransferTaskLine : TP.Object.TransferTaskLine, IWarehouseTaskLine
  {
    #region Cache

    private static AutoCache1Key1Index<DbTransferTaskLine, int, int> cache = new AutoCache1Key1Index<DbTransferTaskLine, int, int>
    {
      Constructor = _ => new DbTransferTaskLine(_),
      Identifier = _ => _.Field<int>(_.Table.Columns.Contains("ID") ? "ID" : "LineID"),
      PrimaryKey = _ => _.TupleID,
      Index1 = _ => _.HeaderID,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    protected DbTransferTaskLine() { }

    public DbTransferTaskLine(DataRow r)
    {
      this.TupleID = r.Field<int>(r.Table.Columns.Contains("ID") ? "ID" : "LineID");
      this.TransferTaskID = (int)r["HeaderID"];
      this.ComponentID = (int)r["ComponentID"];
      this.QuantityRequested = (int)r["QuantityRequested"];
      this.QuantityPicked = (int)r["QuantityPicked"];
      this.QuantityPacked = (int)r["QuantityStaged"];
      this.QuantityTransported = (int)r["QuantityTransported"];
      this.QuantityFilled = (int)r["QuantityFilled"];
    }

    #endregion


    #region Getters

    public static DbTransferTaskLine GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbTransferTaskLine GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityRequested, QuantityPicked, QuantityStaged, QuantityTransported, QuantityFilled FROM WhseTaskTransferLines WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbTransferTaskLine> GetByTaskID(int id, string[] preloads = null) { return GetByTaskID(null, id, preloads); }
    internal static IEnumerable<DbTransferTaskLine> GetByTaskID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ComponentID, QuantityRequested, QuantityPicked, QuantityStaged, QuantityTransported, QuantityFilled FROM WhseTaskTransferLines WHERE HeaderID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbTransferTaskLine> GetPickingRequired(string[] preloads = null) { return GetPickingRequired(null, 5, preloads); }
    public static IEnumerable<DbTransferTaskLine> GetPickingRequired(int warehouseID, string[] preloads = null) { return GetPickingRequired(null, warehouseID, preloads); }
    internal static IEnumerable<DbTransferTaskLine> GetPickingRequired(System.Data.SqlClient.SqlTransaction txn, int warehouseID, string[] preloads = null)
    {
      // completely bypasses cache?
      DataTable dt = Connection.Retrieve(txn,
          "SELECT H.ID AS HeaderID, H.InitialDate, H.SequenceNumber, H.TransferNumber, H.TransferSubtype, H.WarehouseID, H.DestinationWarehouseID, H.IsComplete, H.Description, H.TimeInserted, H.CreatedByUserID, " +
          "       L.ID AS LineID, L.ComponentID, L.QuantityRequested, L.QuantityPicked, L.QuantityStaged, L.QuantityTransported, L.QuantityFilled " +
          "FROM WhseTaskTransferHeader AS H INNER JOIN WhseTaskTransferLines AS L ON H.ID=L.HeaderID " +
          "WHERE H.WarehouseID=@whse " +
          "  AND H.IsComplete=0 " +
          "  AND L.QuantityRequested>L.QuantityPicked " +
          "ORDER BY H.HeaderID, L.LineID", new QueryParameter("@whse", SqlDbType.Int, warehouseID));
      var temp = new Dictionary<int, DbTransferTask>();
      var retval = new List<DbTransferTaskLine>();
      foreach (DataRow r in dt.Rows)
      {
        int hid = r.Field<int>("HeaderID");
        if (!temp.ContainsKey(hid))
        {
          temp.Add(hid, new DbTransferTask(r));
        }
        var l = new DbTransferTaskLine(r);
        l._transfer = temp[hid];
        retval.Add(l);
      }

      return fillPreloads(retval, preloads, txn);
    }

    private static DbTransferTaskLine fillPreloads(DbTransferTaskLine l, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (l == null)
      {
        return l;
      }
      var retval = (DbTransferTaskLine)l.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("component"))
        {
          retval.FillComponent(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbTransferTaskLine> fillPreloads(IEnumerable<DbTransferTaskLine> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbTransferTaskLine has-one DbTransferTask
    public override bool FillTransferTask() { return FillTransferTask(null); }
    internal bool FillTransferTask(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._transfer == null)
      {
        this._transfer = DbTransferTask.GetByID(txn, this.TransferTaskID);
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
      DbInventoryLocation fromLoc = DbInventoryLocation.GetByID(txn, fromLocationID, new string[] { "task" });
      DbInventoryLocation toLoc = toLocationID.HasValue ? DbInventoryLocation.GetByID(txn, toLocationID.Value, new string[] { "task" }) : null;
      this.FillTransferTask(txn);
      int fromWhse = this.TransferTask.WarehouseID;
      int toWhse = this.TransferTask.DestinationWarehouseID;

      List<string> updateFields = new List<string>();
      List<QueryParameter> updateParams = new List<QueryParameter>();

      if (fromLoc.WarehouseID == fromWhse && fromLoc.IsTypeStandard &&                        // from a standard stock location
          toLoc != null && toLoc.WarehouseID == fromWhse && toLoc.IsTypePrepack)              // to a tote/backorder/holding
      {                                                                                       // == pick
        updateFields.Add("QuantityPicked=QuantityPicked+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == fromWhse && fromLoc.IsTypePrepack &&                    // from a tote/backorder/holding
               toLoc != null && toLoc.WarehouseID == fromWhse && toLoc.LocationTypeID == 6)   // to a transfer staging type
      {                                                                                       // == pack
        updateFields.Add("QuantityStaged=QuantityStaged+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == fromWhse && fromLoc.LocationTypeID == 6 &&              // from a transfer staging type
               toLoc != null && toLoc.WarehouseID == -1 && toLoc.LocationTypeID == 10)        // to a vehicle type XXX: note hardcoded -1 warehouse id!
      {                                                                                       // == loadtruck
        updateFields.Add("QuantityTransported=QuantityTransported+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == -1 && fromLoc.LocationTypeID == 10 &&                   // from a vehicle type XXX: note hardcoded -1 warehouse id!
               toLoc != null && toLoc.WarehouseID == toWhse)                                  // to any (?) location in the destination warehouse
      {                                                                                       // == unloadtruck
        updateFields.Add("QuantityFilled=QuantityFilled+@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == -1 && fromLoc.LocationTypeID == 10 &&                   // from a vehicle type XXX: note hardcoded -1 warehouse id!
               toLoc != null && toLoc.WarehouseID == fromWhse && toLoc.LocationTypeID == 6)   // to a transfer staging type
      {                                                                                       // == reversing a loadtruck
        updateFields.Add("QuantityTransported=QuantityTransported-@q1");
        updateParams.Add(new QueryParameter("@q1", SqlDbType.Int, deltaQuantity));
      }
      else if (fromLoc.WarehouseID == fromWhse && fromLoc.LocationTypeID == 6 &&              // from a transfer staging type
               toLoc != null && toLoc.WarehouseID == fromWhse && toLoc.LocationTypeID == 14)  // to a restocking staging type
      {                                                                                       // == reversing a pack && reversing a pick
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
          "UPDATE WhseTaskTransferLines SET " + string.Join(", ", updateFields.ToArray()) + " WHERE ID=@id",
          updateParams.ToArray()
        );

      cache.Clear(this.TupleID, this.HeaderID);
      var newObj = GetByID(txn, this.TupleID);
      if (newObj.QuantityPicked < 0) { throw new ArgumentException("cannot un-pick this many"); }
      if (newObj.QuantityPacked < 0) { throw new ArgumentException("cannot un-pack this many"); }
      if (newObj.QuantityTransported < 0) { throw new ArgumentException("cannot un-load this many"); }
      if (newObj.QuantityFilled < 0) { throw new ArgumentException("cannot un-unload this many"); } // doesn't exist, but for completeness
      //if (newObj.QuantityPicked > newObj.QuantityRequested) { throw new ArgumentException("cannot pick this many"); } // allow overpicking?
      if (newObj.QuantityPacked > newObj.QuantityPicked) { throw new ArgumentException("cannot pack this many"); }
      if (newObj.QuantityTransported > newObj.QuantityPacked) { throw new ArgumentException("cannot load this many"); }
      if (newObj.QuantityFilled > newObj.QuantityTransported) { throw new ArgumentException("cannot unload this many"); }

      return newObj;
    }

    #endregion


    #region Extra Validation

    // none

    #endregion


    // IWarehouseTaskLine
    public int HeaderID { get { return this.TransferTaskID; } }
    public int LineID { get { return this.TupleID; } }
    public IWarehouseTask Header(System.Data.SqlClient.SqlTransaction txn)
    {
      this.FillTransferTask(txn);
      return (DbTransferTask)this.TransferTask;
    }
    public bool HasChangedFromDbCurrent(System.Data.SqlClient.SqlTransaction txn)
    {
      var orig = GetByID(txn, this.LineID);
      return this.PropertiesModified(orig).Count > 0;
    }
  }
}
