using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "InventoryLocationContent", Namespace = "TP.Object")]
  public class FeInventoryLocationContent : TP.Object.InventoryLocationContent
  {

    #region Constructors

    public static FeInventoryLocationContent FromBase(TP.Object.InventoryLocationContent baseObj)
    {
      return new FeInventoryLocationContent(baseObj);
    }
    public static IEnumerable<FeInventoryLocationContent> FromBase(IEnumerable<TP.Object.InventoryLocationContent> baseObjColl)
    {
      IEnumerable<FeInventoryLocationContent> ret = from TP.Object.InventoryLocationContent c in baseObjColl select new FeInventoryLocationContent(c);
      return ret.ToArray();
    }
    public FeInventoryLocationContent(TP.Object.InventoryLocationContent baseObj)
    {
      // properties
      this.LocationID = baseObj.LocationID;
      this.ComponentID = baseObj.ComponentID;
      this.MasterID = baseObj.MasterID;
      this.Quantity = baseObj.Quantity;
      this.AssociatedTaskType = baseObj.AssociatedTaskType;
      this.AssociatedTaskID = baseObj.AssociatedTaskID;
      this.LastInventoriedDate = baseObj.LastInventoriedDate;
      // links
      this._component = baseObj.Component;
      this._location = baseObj.InventoryLocation;
    }

    #endregion


    #region Getters

    public static IEnumerable<FeInventoryLocationContent> GetByLocationID(int locationID, bool preloadLocation = false, bool preloadComponent = false)
    {
      var pl = preloads(preloadLocation, preloadComponent);
      return Extensions.Get<IEnumerable<FeInventoryLocationContent>>(pl, "invlocation", locationID, "contents");
    }

    public static IEnumerable<FeInventoryLocationContent> GetByComponentID(int componentID, bool preloadLocation = false, bool preloadComponent = false)
    {
      var pl = preloads(preloadLocation, preloadComponent);
      return Extensions.Get<IEnumerable<FeInventoryLocationContent>>(pl, "component", componentID, "locations");
    }

    #endregion


    #region Links

    public override bool FillComponent()
    {
      if (this._component == null)
      {
        this._component = FeComponent.GetByID(this.ComponentID);
      }
      return true;
    }

    public override bool FillInventoryLocation()
    {
      if (this._location == null)
      {
        this._location = FeInventoryLocation.GetByLocationID(this.LocationID);
      }
      return true;
    }

    #endregion


    public static string[] preloads(bool preloadLocation = false, bool preloadComponent = false)
    {
      var pl = new List<string>();
      if (preloadLocation) { pl.Add("location"); }
      if (preloadComponent) { pl.Add("component"); }
      return pl.ToArray();
    }
  }
}
