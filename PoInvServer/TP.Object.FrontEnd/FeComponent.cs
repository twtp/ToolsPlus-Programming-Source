using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Component", Namespace = "TP.Object")]
  public class FeComponent : TP.Object.Component
  {

    #region Constructors

    public static FeComponent FromBase(TP.Object.Component baseObj)
    {
      return new FeComponent(baseObj);
    }
    public static IEnumerable<FeComponent> FromBase(IEnumerable<TP.Object.Component> baseObjColl) {
      IEnumerable<FeComponent> ret = from TP.Object.Component i in baseObjColl select new FeComponent(i);
      return ret.ToArray();
    }
    public FeComponent(TP.Object.Component baseObj)
    {
      // properties
      this.ComponentID = baseObj.ComponentID;
      this.ComponentName = baseObj.ComponentName;
      this.Length = baseObj.Length;
      this.Width = baseObj.Width;
      this.Height = baseObj.Height;
      this.Weight = baseObj.Weight;
      this.PackingType = baseObj.PackingType;
      this.IsNotConcealableDuringShipping = baseObj.IsNotConcealableDuringShipping;
      this.IsHazardous = baseObj.IsHazardous;
      this.IsTrackedBySerialNo = baseObj.IsTrackedBySerialNo;
      this.IsSpecificSideUp = baseObj.IsSpecificSideUp;
      // links
      this._barcodes = baseObj.Barcodes;
      this._locations = baseObj.Locations;
    }

    #endregion


    #region Getters

    public static FeComponent GetByID(int componentID, bool preloadBarcodes = false, bool preloadLocations = false)
    {
      var pl = preloads(preloadBarcodes, preloadLocations);
      return Extensions.Get<FeComponent>(pl, "component", componentID);
    }

    #endregion


    #region Links

    public override bool FillBarcodeList()
    {
      if (this._barcodes == null)
      {
        this._barcodes = Extensions.Get<IEnumerable<FeBarcode>>("barcode", "component", this.ComponentID);
      }
      return true;
    }

    public override bool FillLocations()
    {
      if (this._locations == null)
      {
        this._locations = Extensions.Get<IEnumerable<FeInventoryLocationContent>>("component", this.ComponentID, "locations");
      }
      return true;
    }

    #endregion


    public static string[] preloads(bool preloadBarcodes = false, bool preloadLocations = false)
    {
      var pl = new List<string>();
      if (preloadBarcodes) { pl.Add("barcodes"); }
      if (preloadLocations) { pl.Add("locations"); }
      return pl.ToArray();
    }
  }

}
