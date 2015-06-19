using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "InventoryLocation", Namespace = "TP.Object")]
  public class DbInventoryLocation : TP.Object.InventoryLocation
  {
    #region Cache

    private static AutoCache1Key<DbInventoryLocation, int> cache = new AutoCache1Key<DbInventoryLocation, int>
    {
      Constructor = _ => new DbInventoryLocation(_),
      Identifier = _ => _.Field<int>("ID"),
      PrimaryKey = _ => _.LocationID,
      SecondsToLive = 86400 // 1 day
    };

    #endregion


    #region Constructors

    public DbInventoryLocation(DataRow r) : base()
    {
      this.LocationID = (int)r["ID"];
      this.LocationTypeID = (int)r["LocationTypeID"];
      this.WarehouseID = (int)r["WarehouseID"];
      this.Division1 = (int)r["Division1"];
      this.Division2 = this.nullable<int>(r["Division2"]);
      this.Division3 = this.nullable<int>(r["Division3"]);
      this.Division4 = this.nullable<int>(r["Division4"]);
      this.Division5 = this.nullable<int>(r["Division5"]);
      this.LeftID = this.nullable<int>(r["LeftID"]);
      this.RightID = this.nullable<int>(r["RightID"]);
      this.FrontID = this.nullable<int>(r["FrontID"]);
      this.RearID = this.nullable<int>(r["RearID"]);
    }

    #endregion


    #region Getters

    // /invlocation/{id:int}
    public static DbInventoryLocation GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbInventoryLocation GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, LocationTypeID, WarehouseID, Division1, Division2, Division3, Division4, Division5, LeftID, RightID, FrontID, RearID FROM LocationMaster WHERE ID=@id AND Deleted=0", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /invlocation/{bc}
    public static DbInventoryLocation GetByBarcode(string bc, string[] preloads = null) { return GetByBarcode(null, bc, preloads); }
    internal static DbInventoryLocation GetByBarcode(System.Data.SqlClient.SqlTransaction txn, string bc, string[] preloads = null)
    {
      if (bc.Length != 12 || bc.StartsWith("LOC") == false) { return null; }
      int id;
      bool parsed = Int32.TryParse(bc.Substring(3, 9), out id);
      if (parsed == false)
      {
        return null;
      }
      return GetByID(txn, id, preloads);
    }

    public static DbInventoryLocation GetSpanEndpoint(int masterID, string[] preloads = null) { return GetSpanEndpoint(null, masterID, preloads); }
    internal static DbInventoryLocation GetSpanEndpoint(System.Data.SqlClient.SqlTransaction txn, int masterID, string[] preloads = null) {
      var loc = GetByID(txn, masterID);
      var end = loc.GetSpanEndpointLocationID(
          _ => _.FillContents(txn),
          _ => DbInventoryLocationContent.GetByMasterID(txn, _)
        );
      return GetByID(txn, end, preloads);
    }

    public static bool IsSpanValid(int startID, int endID) { return IsSpanValid(null, startID, endID); }
    public static bool IsSpanValid(System.Data.SqlClient.SqlTransaction txn, int startID, int endID)
    {
      var l1 = GetByID(txn, startID);
      var l2 = GetByID(txn, endID);
      return l1.CanSpanTo(l2, _ => _.FillContents(txn), _ => _.FillNeighbors(txn));
    }

    private static DbInventoryLocation fillPreloads(DbInventoryLocation loc, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (loc == null)
      {
        return loc;
      }
      var retval = (DbInventoryLocation)loc.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("contents"))
        {
          retval.FillContents(txn);
        }
        if (preloads.Contains("neighbors"))
        {
          retval.FillNeighbors(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbInventoryLocation> fillPreloads(IEnumerable<DbInventoryLocation> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbInventoryLocation has-many DbInventoryLocationContent
    public override bool FillContents() { return FillContents(null); }
    public override bool FillContents(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._contents == null)
      {
        this._contents = DbInventoryLocationContent.GetByLocationID(txn, this.LocationID);
      }
      return true;
    }

    // (4) InventoryLocation has-one InventoryLocation
    public override bool FillNeighbors() { return FillNeighbors(null); }
    public override bool FillNeighbors(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._neighbor_left == null && this.LeftID.HasValue)
      {
        this._neighbor_left = DbInventoryLocation.GetByID(txn, this.LeftID.Value);
      }
      if (this._neighbor_right == null && this.RightID.HasValue)
      {
        this._neighbor_right = DbInventoryLocation.GetByID(txn, this.RightID.Value);
      }
      if (this._neighbor_front == null && this.FrontID.HasValue)
      {
        this._neighbor_front = DbInventoryLocation.GetByID(txn, this.FrontID.Value);
      }
      if (this._neighbor_rear == null && this.RearID.HasValue)
      {
        this._neighbor_rear = DbInventoryLocation.GetByID(txn, this.RearID.Value);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // no create
    // no update
    // no delete

    #endregion


    #region Extra Validation

    internal static void AssertLocationExists(int locationID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      var l = DbInventoryLocation.GetByID(txn, locationID);
      if (l == null)
      {
        throw new ArgumentException("location id " + locationID + " not in database");
      }
      return;
    }

    #endregion

  }
}
