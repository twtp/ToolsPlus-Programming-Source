using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_CI_Item", Namespace = "TP.Object")]
  public class CI_Item : IMASExportObject
  {
    [DataMember]
    public string ItemCode { get; set; }
    [DataMember]
    public string ItemType { get; set; }
    [DataMember]
    public string ItemCodeDesc { get; set; }
    [DataMember]
    public string ProductLine { get; set; }
    [DataMember]
    public string PrimaryVendorNo { get; set; }
    [DataMember]
    public decimal StandardUnitCost { get; set; }
    [DataMember]
    public decimal StandardUnitPrice { get; set; }
    [DataMember]
    public decimal SuggestedRetailPrice { get; set; }
    [DataMember]
    public string TaxClass { get; set; }
    [DataMember]
    public string RQuantityBreak { get; set; }
    [DataMember]
    public string Discontinued { get; set; }
    [DataMember]
    public string EPN { get; set; }
    [DataMember]
    public int StdPack { get; set; }
    [DataMember]
    public decimal WebPrice { get; set; }
    //[DataMember]
    //public string Dimensions { get; set; }
    //[DataMember]
    //public string Weight_New { get; set; }
    [DataMember]
    public TP.Object.IExportableToMAS OriginalObject { get; set; }



    public static CI_Item FromObject(TP.Object.Item i)
    {
      decimal stdCost = 0;
      decimal stdPrice = 0;
      bool stdPriceBreaks = false;
      decimal webPrice = 0;
      decimal listPrice = 0;
      ItemCost c;
      ItemPrice p;
      if ((c = i.ItemCosts.SingleOrDefault(_ => _.CostType == "Std Cost")) != null) {
        stdCost = c.CostValue;
      }
      if ((p = i.ItemPrices.SingleOrDefault(_ => _.SalesChannelID == 0)) != null) {
        stdPrice = p.PriceValue1;
        stdPriceBreaks = p.PriceValue2 != 0;
      }
      if ((p = i.ItemPrices.SingleOrDefault(_ => _.SalesChannelID == 1)) != null)
      {
        webPrice = p.PriceValue1;
      }
      if ((c = i.ItemCosts.SingleOrDefault(_ => _.CostType == "List Price")) != null)
      {
        listPrice = c.CostValue;
      }

      return new CI_Item
      {
        ItemCode = i.ItemNumber,
        ItemType = i.IsMasKit == false ? "F" : "K",
        ItemCodeDesc = i.ItemDescription,
        ProductLine = i.ProductLinePrefix(),
        PrimaryVendorNo = i.PrimaryVendor,
        StandardUnitCost = stdCost,
        StandardUnitPrice = stdPrice,
        SuggestedRetailPrice = listPrice,
        TaxClass = i.IsTaxExempt == false || stdPrice > 50 ? "TX" : "NT",
        RQuantityBreak = stdPriceBreaks ? "CHECK BREAKS ==>" : string.Empty,
        Discontinued = i.IsDiscontinued ? "*** DISCONTINUED DISCONTINUED ***" : string.Empty, // TODO: being d/c?
        EPN = i.AdditionalInfo,
        StdPack = i.StandardPack,
        WebPrice = webPrice,
        //Dimensions = "",
        //Weight_New = "",
        OriginalObject = i,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0}",
          this.csvEscape(this.ItemCode)
        );
    }

  }
}
