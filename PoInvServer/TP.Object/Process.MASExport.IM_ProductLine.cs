using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_IM_ProductLine", Namespace = "TP.Object")]
  public class IM_ProductLine : IMASExportObject
  {
    public static string MASColumns() { return "ProductLine, ProductLineDesc, InventoryAcctKey, CostOfGoodsSoldAcctKey, SalesIncomeAcctKey, ReturnsAcctKey, AdjustmentAcctKey, PurchaseAcctKey, PurchaseOrderVarianceAcctKey, ManufacturingVarianceAcctKey, ScrapAcctKey, RepairsInProcessAcctKey, RepairsClearingAcctKey"; }
    public static string MASTableName() { return "IM_ProductLine"; }
    public string PrimaryKey() { return this.ProductLineCode; }

    [DataMember]
    public string ProductLineCode { get; set; }
    [DataMember]
    public string ProductLineDescription { get; set; }
    [DataMember]
    public string InventoryAccount { get; set; }
    [DataMember]
    public string CostOfGoodsSoldAccount { get; set; }
    [DataMember]
    public string SalesIncomeAccount { get; set; }
    [DataMember]
    public string SalesReturnsAccount { get; set; }
    [DataMember]
    public string InventoryAdjustmentAccount { get; set; }
    [DataMember]
    public string PurchasesClearingAccount { get; set; }
    [DataMember]
    public string PurchaseOrderVarianceAdjustmentAccount { get; set; }
    [DataMember]
    public string ManufacturerVarianceAdjustmentAccount { get; set; }
    [DataMember]
    public string ReturnMerchandiseAuthorizationScrapAccount { get; set; }
    [DataMember]
    public string RepairsInProgressAccount { get; set; }
    [DataMember]
    public string RepairsClearingAccount { get; set; }

    public static IM_ProductLine FromObject(ProductLine pl)
    {
      return new IM_ProductLine
      {
        ProductLineCode = pl.ProductLinePrefix,
        ProductLineDescription = pl.ProductLineName,
        InventoryAccount = getInventoryAccount(pl.ProductLinePrefix),
        CostOfGoodsSoldAccount = getCostOfGoodsSoldAccount(pl.ProductLinePrefix),
        SalesIncomeAccount = getSalesIncomeAccount(pl.ProductLinePrefix),
        SalesReturnsAccount = getSalesReturnsAccount(pl.ProductLinePrefix),
        InventoryAdjustmentAccount = getInventoryAdjustmentAccount(pl.ProductLinePrefix),
        PurchasesClearingAccount = getPurchasesClearingAccount(pl.ProductLinePrefix),
        PurchaseOrderVarianceAdjustmentAccount = getPurchaseOrderVarianceAdjustmentAccount(pl.ProductLinePrefix),
        ManufacturerVarianceAdjustmentAccount = getManufacturerVarianceAdjustmentAccount(pl.ProductLinePrefix),
        ReturnMerchandiseAuthorizationScrapAccount = getReturnMerchandiseAuthorizationScrapAccount(pl.ProductLinePrefix),
        RepairsInProgressAccount = getRepairsInProgressAccount(pl.ProductLinePrefix),
        RepairsClearingAccount = getRepairsClearingAccount(pl.ProductLinePrefix),
      };
    }

    // NOTE: 1152-00 (99990001B) is inactive? this is probably bad?
    private static string getInventoryAccount(string prefix) { return acctTypeStd(prefix) ? "140000000" : "99990001B"; }
    private static string getCostOfGoodsSoldAccount(string prefix) { return acctTypeStd(prefix) ? "451000000" : "99990001B"; }
    private static string getSalesIncomeAccount(string prefix) { return acctTypeStd(prefix) ? "401000000" : "100000000"; }
    private static string getSalesReturnsAccount(string prefix) { return acctTypeStd(prefix) ? "408000000" : "100000000"; }
    private static string getInventoryAdjustmentAccount(string prefix) { return "470400000"; }
    private static string getPurchasesClearingAccount(string prefix) { return acctTypeStd(prefix) ? "201000000" : "99990001B"; }
    private static string getPurchaseOrderVarianceAdjustmentAccount(string prefix) { return "470500000"; }
    private static string getManufacturerVarianceAdjustmentAccount(string prefix) { return "470500000"; }
    private static string getReturnMerchandiseAuthorizationScrapAccount(string prefix) { return "587502000"; }
    private static string getRepairsInProgressAccount(string prefix) { return "587502000"; }
    private static string getRepairsClearingAccount(string prefix) { return "587502000"; }

    private static bool acctTypeStd(string pl) { return pl != "GFT"; }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}",
          this.csvEscape(this.ProductLineCode),
          this.csvEscape(this.ProductLineDescription),
          this.csvEscape(this.InventoryAccount),
          this.csvEscape(this.CostOfGoodsSoldAccount),
          this.csvEscape(this.SalesIncomeAccount),
          this.csvEscape(this.SalesReturnsAccount),
          this.csvEscape(this.InventoryAdjustmentAccount),
          this.csvEscape(this.PurchasesClearingAccount),
          this.csvEscape(this.PurchaseOrderVarianceAdjustmentAccount),
          this.csvEscape(this.ManufacturerVarianceAdjustmentAccount),
          this.csvEscape(this.ReturnMerchandiseAuthorizationScrapAccount),
          this.csvEscape(this.RepairsInProgressAccount),
          this.csvEscape(this.RepairsClearingAccount)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("ProductLine", this.ProductLineCode, r.Field<string>("ProductLine"));
      yield return new Tuple<string, object, object>("ProductLineDesc", this.ProductLineDescription, r.Field<string>("ProductLineDesc"));
      yield return new Tuple<string, object, object>("InventoryAcctKey", this.InventoryAccount, r.Field<string>("InventoryAcctKey"));
      yield return new Tuple<string, object, object>("CostOfGoodsSoldAcctKey", this.CostOfGoodsSoldAccount, r.Field<string>("CostOfGoodsSoldAcctKey"));
      yield return new Tuple<string, object, object>("SalesIncomeAcctKey", this.SalesIncomeAccount, r.Field<string>("SalesIncomeAcctKey"));
      yield return new Tuple<string, object, object>("ReturnsAcctKey", this.SalesReturnsAccount, r.Field<string>("ReturnsAcctKey"));
      yield return new Tuple<string, object, object>("AdjustmentAcctKey", this.InventoryAdjustmentAccount, r.Field<string>("AdjustmentAcctKey"));
      yield return new Tuple<string, object, object>("PurchaseAcctKey", this.PurchasesClearingAccount, r.Field<string>("PurchaseAcctKey"));
      yield return new Tuple<string, object, object>("PurchaseOrderVarianceAcctKey", this.PurchaseOrderVarianceAdjustmentAccount, r.Field<string>("PurchaseOrderVarianceAcctKey"));
      yield return new Tuple<string, object, object>("ManufacturingVarianceAcctKey", this.ManufacturerVarianceAdjustmentAccount, r.Field<string>("ManufacturingVarianceAcctKey"));
      yield return new Tuple<string, object, object>("ScrapAcctKey", this.ReturnMerchandiseAuthorizationScrapAccount, r.Field<string>("ScrapAcctKey"));
      yield return new Tuple<string, object, object>("RepairsInProcessAcctKey", this.RepairsInProgressAccount, r.Field<string>("RepairsInProcessAcctKey"));
      yield return new Tuple<string, object, object>("RepairsClearingAcctKey", this.RepairsClearingAccount, r.Field<string>("RepairsClearingAcctKey"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<ProductLine> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(ProductLine current)
    {
      if (this.ProductLineCode != current.ProductLinePrefix) { return false; }
      if (this.ProductLineDescription != current.ProductLineName) { return false; }
      if (this.InventoryAccount != getInventoryAccount(current.ProductLinePrefix)) { return false; }
      if (this.CostOfGoodsSoldAccount != getCostOfGoodsSoldAccount(current.ProductLinePrefix)) { return false; }
      if (this.SalesIncomeAccount != getSalesIncomeAccount(current.ProductLinePrefix)) { return false; }
      if (this.SalesReturnsAccount != getSalesReturnsAccount(current.ProductLinePrefix)) { return false; }
      if (this.InventoryAdjustmentAccount != getInventoryAdjustmentAccount(current.ProductLinePrefix)) { return false; }
      if (this.PurchasesClearingAccount != getPurchasesClearingAccount(current.ProductLinePrefix)) { return false; }
      if (this.PurchaseOrderVarianceAdjustmentAccount != getPurchaseOrderVarianceAdjustmentAccount(current.ProductLinePrefix)) { return false; }
      if (this.ManufacturerVarianceAdjustmentAccount != getManufacturerVarianceAdjustmentAccount(current.ProductLinePrefix)) { return false; }
      if (this.ReturnMerchandiseAuthorizationScrapAccount != getReturnMerchandiseAuthorizationScrapAccount(current.ProductLinePrefix)) { return false; }
      if (this.RepairsInProgressAccount != getRepairsInProgressAccount(current.ProductLinePrefix)) { return false; }
      if (this.RepairsClearingAccount != getRepairsClearingAccount(current.ProductLinePrefix)) { return false; }

      return true;
    }

  }
}
