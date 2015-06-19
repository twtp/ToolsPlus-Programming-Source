using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "SalesOrderTaskLine", Namespace = "TP.Object")]
  public class FeSalesOrderTaskLine : TP.Object.SalesOrderTaskLine
  {

    #region Constructors

    public static FeSalesOrderTaskLine FromBase(TP.Object.SalesOrderTaskLine baseObj)
    {
      return new FeSalesOrderTaskLine(baseObj);
    }
    public static IEnumerable<FeSalesOrderTaskLine> FromBase(IEnumerable<TP.Object.SalesOrderTaskLine> baseObjColl)
    {
      IEnumerable<FeSalesOrderTaskLine> ret = from TP.Object.SalesOrderTaskLine i in baseObjColl select new FeSalesOrderTaskLine(i);
      return ret.ToArray();
    }
    public FeSalesOrderTaskLine(TP.Object.SalesOrderTaskLine baseObj)
    {
      // properties
      this.TupleID = baseObj.TupleID;
      this.SalesOrderTaskID = baseObj.SalesOrderTaskID;
      this.ComponentID = baseObj.ComponentID;
      this.QuantityOrdered = baseObj.QuantityOrdered;
      this.QuantityBackordered = baseObj.QuantityBackordered;
      this.QuantityPicked = baseObj.QuantityPicked;
      this.QuantityPacked = baseObj.QuantityPacked;
      this.QuantityShipped = baseObj.QuantityShipped;
      // links
      this._salesorder = baseObj.SalesOrderTask;
      this._component = baseObj.Component;
    }

    #endregion


    #region Getters

    public static IEnumerable<FeSalesOrderTaskLine> GetBySalesOrderTaskID(int salesOrderTaskID, bool preloadComponents = false)
    {
      var pl = preloads(preloadComponents);
      return Extensions.Get<IEnumerable<FeSalesOrderTaskLine>>("task", "salesorder", "id", salesOrderTaskID, "lines");
    }

    #endregion


    #region Links

    public override bool FillSalesOrderTask()
    {
      if (this._salesorder == null)
      {
        this._salesorder = Extensions.Get<FeSalesOrderTask>("task", "salesorder", "id", this.SalesOrderTaskID);
      }
      return true;
    }

    public override bool FillComponent()
    {
      if (this._component == null)
      {
        this._component = Extensions.Get<FeComponent>("component", this.ComponentID);
      }
      return true;
    }

    #endregion


    public static string[] preloads(bool preloadComponents = false)
    {
      var pl = new List<string>();
      if (preloadComponents) { pl.Add("component"); }
      return pl.ToArray();
    }
  }

}
