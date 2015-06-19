using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "InventoryLocation", Namespace = "TP.Object")]
  public class FeInventoryLocation : TP.Object.InventoryLocation
  {

    #region Constructors

    public static FeInventoryLocation FromBase(TP.Object.InventoryLocation baseObj)
    {
      return new FeInventoryLocation(baseObj);
    }
    public static IEnumerable<FeInventoryLocation> FromBase(IEnumerable<TP.Object.InventoryLocation> baseObjColl) {
      IEnumerable<FeInventoryLocation> ret = from TP.Object.InventoryLocation i in baseObjColl select new FeInventoryLocation(i);
      return ret.ToArray();
    }
    public FeInventoryLocation(TP.Object.InventoryLocation baseObj)
    {
      // properties
      this.LocationID = baseObj.LocationID;
      this.LocationTypeID = baseObj.LocationTypeID;
      this.WarehouseID = baseObj.WarehouseID;
      this.Division1 = baseObj.Division1;
      this.Division2 = baseObj.Division2;
      this.Division3 = baseObj.Division3;
      this.Division4 = baseObj.Division4;
      this.Division5 = baseObj.Division5;
      this.LeftID = baseObj.LeftID;
      this.RightID = baseObj.RightID;
      this.FrontID = baseObj.FrontID;
      this.RearID = baseObj.RearID;
      // links
      this._contents = baseObj.Contents;
      this._neighbor_left = baseObj.NeighborLeft;
      this._neighbor_right = baseObj.NeighborRight;
      this._neighbor_front = baseObj.NeighborFront;
      this._neighbor_rear = baseObj.NeighborRear;
    }

    #endregion


    #region Getters

    public static FeInventoryLocation GetByLocationID(int locationID, bool preloadContents = false, bool preloadNeighbors = false)
    {
      var pl = preloads(preloadContents, preloadNeighbors);
      return Extensions.Get<FeInventoryLocation>(pl, "invlocation", locationID);
    }

    public static FeInventoryLocation GetByBarcode(string barcode, bool preloadContents = false, bool preloadNeighbors = false)
    {
      var pl = preloads(preloadContents, preloadNeighbors);
      return Extensions.Get<FeInventoryLocation>(pl, "invlocation", barcode);
    }

    #endregion


    #region Links

    public override bool FillContents()
    {
      if (this._contents == null) {
        this._contents = Extensions.Get<IEnumerable<FeInventoryLocationContent>>("invlocation", this.LocationID, "contents");
      }
      return true;
    }
    public override bool FillContents(System.Data.SqlClient.SqlTransaction txn)
    {
      return FillContents();
    }

    public override bool FillNeighbors()
    {
      if (this._neighbor_left == null && this.LeftID.HasValue)
      {
        this._neighbor_left = FeInventoryLocation.GetByLocationID(this.LeftID.Value);
      }
      if (this._neighbor_right == null && this.RightID.HasValue)
      {
        this._neighbor_right = FeInventoryLocation.GetByLocationID(this.RightID.Value);
      }
      if (this._neighbor_front == null && this.FrontID.HasValue)
      {
        this._neighbor_front = FeInventoryLocation.GetByLocationID(this.FrontID.Value);
      }
      if (this._neighbor_rear == null && this.RearID.HasValue)
      {
        this._neighbor_rear = FeInventoryLocation.GetByLocationID(this.RearID.Value);
      }
      return true;
    }
    public override bool FillNeighbors(System.Data.SqlClient.SqlTransaction txn)
    {
      return FillNeighbors();
    }

    #endregion


    public static string[] preloads(bool preloadContents = false, bool preloadNeighbors = false)
    {
      var pl = new List<string>();
      if (preloadContents) { pl.Add("contents"); }
      if (preloadNeighbors) { pl.Add("neighbors"); }
      return pl.ToArray();
    }
  }

}
