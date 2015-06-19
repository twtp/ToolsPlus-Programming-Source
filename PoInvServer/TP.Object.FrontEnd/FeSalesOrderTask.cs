using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  [System.Runtime.Serialization.DataContract(Name = "SalesOrderTask", Namespace = "TP.Object")]
  public class FeSalesOrderTask : TP.Object.SalesOrderTask
  {

    #region Constructors

    public static FeSalesOrderTask FromBase(TP.Object.SalesOrderTask baseObj)
    {
      return new FeSalesOrderTask(baseObj);
    }
    public static IEnumerable<FeSalesOrderTask> FromBase(IEnumerable<TP.Object.SalesOrderTask> baseObjColl)
    {
      IEnumerable<FeSalesOrderTask> ret = from TP.Object.SalesOrderTask i in baseObjColl select new FeSalesOrderTask(i);
      return ret.ToArray();
    }
    public FeSalesOrderTask(TP.Object.SalesOrderTask baseObj)
    {
      // properties
      this.SalesOrderTaskID = baseObj.SalesOrderTaskID;
      this.SalesOrderNo = baseObj.SalesOrderNo;
      this.WarehouseID = baseObj.WarehouseID;
      this.CurrentStatus = baseObj.CurrentStatus;
      this.CurrentHoldReason = baseObj.CurrentHoldReason;
      this.ShipVia = baseObj.ShipVia;
      this.ShipState = baseObj.ShipState;
      this.ShipCountry = baseObj.ShipCountry;
      this.MasCreatedByUser = baseObj.MasCreatedByUser;
      this.TimeInserted = baseObj.TimeInserted;
      // links
      this._lines = baseObj.Lines;
      this._saleschannel = baseObj.SalesChannel;
    }

    #endregion


    #region Getters

    public static FeSalesOrderTask GetByID(int salesOrderID, bool preloadLines = false)
    {
      var pl = preloads(preloadLines);
      return Extensions.Get<FeSalesOrderTask>("task", "salesorder", "id", salesOrderID);
    }

    public static FeSalesOrderTask GetBySalesOrderNo(string salesOrderNo, bool preloadLines = false)
    {
      var pl = preloads(preloadLines);
      return Extensions.Get<FeSalesOrderTask>("task", "salesorder", "no", salesOrderNo);
    }

    #endregion


    #region Links

    public override bool FillLines()
    {
      if (this._lines == null)
      {
        this._lines = Extensions.Get<IEnumerable<FeSalesOrderTaskLine>>("salesorder", "id", this.SalesOrderTaskID, "lines");
      }
      return true;
    }

    public override bool FillSalesChannel()
    {
      if (this._saleschannel == null)
      {
        throw new NotImplementedException("TODO: FeSalesChannel");
        //this._saleschannel = Extensions.Get<IEnumerable<FeSalesChannel>>("saleschannel", "salesorder", this.SalesOrderNo);
      }
      return true;
    }

    #endregion


    public static string[] preloads(bool preloadLines = false)
    {
      var pl = new List<string>();
      if (preloadLines) { pl.Add("lines"); }
      return pl.ToArray();
    }
  }

}
