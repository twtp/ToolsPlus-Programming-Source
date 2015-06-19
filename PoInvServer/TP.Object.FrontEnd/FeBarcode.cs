using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Barcode", Namespace = "TP.Object")]
  public class FeBarcode : TP.Object.Barcode
  {

    #region Constructors

    public static FeBarcode FromBase(TP.Object.Barcode baseObj)
    {
      return new FeBarcode(baseObj);
    }
    public static IEnumerable<FeBarcode> FromBase(IEnumerable<TP.Object.Barcode> baseObjColl) {
      IEnumerable<FeBarcode> ret = from TP.Object.Barcode i in baseObjColl select new FeBarcode(i);
      return ret.ToArray();
    }
    public FeBarcode(TP.Object.Barcode baseObj)
    {
      // properties
      this.Code = baseObj.Code;
      this.ComponentID = baseObj.ComponentID;
      this.IsInternal = baseObj.IsInternal;
      this.PrintByStandardPack = baseObj.PrintByStandardPack;
      this.Quantity = baseObj.Quantity;
      // links
    }

    #endregion


    #region Getters

    public static FeBarcode GetByCode(string code)
    {
      return Extensions.Get<FeBarcode>(null, "barcode", code);
    }

    #endregion


    #region Links

    // none

    #endregion


    // no preloads to handle
  }

}
