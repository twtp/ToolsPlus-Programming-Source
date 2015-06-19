using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_IM_ItemPrice", Namespace = "TP.Object")]
  public class IM_PriceCode : IMASExportObject
  {
    public static string MASColumns() { return "ItemCode, CustomerPriceLevel, PricingMethod, BreakQuantity1, BreakQuantity2, BreakQuantity3, BreakQuantity4, BreakQuantity5, DiscountMarkup1, DiscountMarkup2, DiscountMarkup3, DiscountMarkup4, DiscountMarkup5"; }
    public static string MASTableName() { return "IM_PriceCode"; }
    public string PrimaryKey() { return string.Format("{0}/{1}", this.ItemCode, this.CustomerPriceLevel); }

    [DataMember]
    public string ItemCode { get; set; }
    [DataMember]
    public string CustomerPriceLevel { get; set; }
    [DataMember]
    public string PricingMethod { get; set; }
    [DataMember]
    public int BreakQuantity1 { get; set; }
    [DataMember]
    public decimal DiscountMarkup1 { get; set; }
    [DataMember]
    public int BreakQuantity2 { get; set; }
    [DataMember]
    public decimal DiscountMarkup2 { get; set; }
    [DataMember]
    public int BreakQuantity3 { get; set; }
    [DataMember]
    public decimal DiscountMarkup3 { get; set; }
    [DataMember]
    public int BreakQuantity4 { get; set; }
    [DataMember]
    public decimal DiscountMarkup4 { get; set; }
    [DataMember]
    public int BreakQuantity5 { get; set; }
    [DataMember]
    public decimal DiscountMarkup5 { get; set; }
    [DataMember]
    public int InternalID { get; set; }


    public static IM_PriceCode FromObject(ItemPrice p)
    {
      // ASSERT: p.SalesChannelID == 0 || p.SalesChannelID == 1
      return new IM_PriceCode
      {
        ItemCode = p.Item.ItemNumber,
        CustomerPriceLevel = getCustomerPriceLevel(p.SalesChannelID),
        PricingMethod = getPricingMethod(),
        BreakQuantity1 = p.MaximumQuantity1,
        DiscountMarkup1 = p.PriceValue1,
        BreakQuantity2 = p.MaximumQuantity2,
        DiscountMarkup2 = p.PriceValue2,
        BreakQuantity3 = p.MaximumQuantity3,
        DiscountMarkup3 = p.PriceValue3,
        BreakQuantity4 = p.MaximumQuantity4,
        DiscountMarkup4 = p.PriceValue4,
        BreakQuantity5 = p.MaximumQuantity5,
        DiscountMarkup5 = p.PriceValue5,
        InternalID = p.TupleID
      };
    }

    public static IEnumerable<IM_PriceCode> FromObject(Item i, Action<Item> fillProductLine = null, Action<Item> fillWholesales = null)
    {
      if (fillProductLine == null) { fillProductLine = _ => _.FillProductLine(); }
      if (fillWholesales == null) { fillWholesales = _ => _.ProductLine.FillWholesalePriceLevels(); }

      var retval = new List<IM_PriceCode>();
      retval.Add(new IM_PriceCode { ItemCode = i.ItemNumber, CustomerPriceLevel = "G", PricingMethod = "M", BreakQuantity1 = 99999999, DiscountMarkup1 = 0 });
      retval.Add(new IM_PriceCode { ItemCode = i.ItemNumber, CustomerPriceLevel = "H", PricingMethod = "M", BreakQuantity1 = 99999999, DiscountMarkup1 = 10 });
      foreach (var l in i.ProductLine.WholesalePriceLevels) {
        retval.Add(new IM_PriceCode { ItemCode = i.ItemNumber, CustomerPriceLevel = l.LevelLetter(), PricingMethod = "M", BreakQuantity1 = 99999999, DiscountMarkup1 = l.Percentage });
      }
      return retval;
    }

    private static string getCustomerPriceLevel(int salesChannelID)
    {
      switch (salesChannelID)
      {
        case 0:
          return string.Empty;
        case 1:
          return "2";
        default:
          throw new ApplicationException("unexpected salesChannelID");
      }
    }
    private static string getPricingMethod()
    {
      return "O";
    }

    #region crap for wholesales, etc
    /*public static IEnumerable<IM_PriceCode> FromObject(TP.Object.Item i, Action<TP.Object.Item> fillCurrentPrices, Action<TP.Object.Item> fillProductLine, Action<TP.Object.Item> fillWholesalePriceLevels)
    {
      var retval = new List<IM_PriceCode>();
      if (i.IsExportRequiredStorePrice || i.IsExportRequiredWebPrice)
      {
        fillCurrentPrices(i);
        if (i.IsExportRequiredStorePrice)
        {
          TP.Object.ItemPrice p = i.ItemPrices.Single(_ => _.SalesChannelID == 0);
          retval.Add(new IM_PriceCode {
              ItemCode = i.ItemNumber,
              CustomerPriceLevel = string.Empty,
              PricingMethod = "O",
              BreakQuantity1 = p.MaximumQuantity1,
              DiscountMarkup1 = p.PriceValue1,
              BreakQuantity2 = p.MaximumQuantity2,
              DiscountMarkup2 = p.PriceValue2,
              BreakQuantity3 = p.MaximumQuantity3,
              DiscountMarkup3 = p.PriceValue3,
              BreakQuantity4 = p.MaximumQuantity4,
              DiscountMarkup4 = p.PriceValue4,
              BreakQuantity5 = p.MaximumQuantity5,
              DiscountMarkup5 = p.PriceValue5,
              OriginalObject = i,
            });
        }
        if (i.IsExportRequiredWebPrice)
        {
          TP.Object.ItemPrice p = i.ItemPrices.Single(_ => _.SalesChannelID == 1);
          retval.Add(new IM_PriceCode
          {
            ItemCode = i.ItemNumber,
            CustomerPriceLevel = "2",
            PricingMethod = "O",
            BreakQuantity1 = p.MaximumQuantity1,
            DiscountMarkup1 = p.PriceValue1,
            BreakQuantity2 = p.MaximumQuantity2,
            DiscountMarkup2 = p.PriceValue2,
            BreakQuantity3 = p.MaximumQuantity3,
            DiscountMarkup3 = p.PriceValue3,
            BreakQuantity4 = p.MaximumQuantity4,
            DiscountMarkup4 = p.PriceValue4,
            BreakQuantity5 = p.MaximumQuantity5,
            DiscountMarkup5 = p.PriceValue5,
            OriginalObject = i,
          });
        }
      }
      if (i.IsExportRequiredNewItem)
      {
        retval.Add(new IM_PriceCode(i.ItemNumber, "G", "M", 99999999, 0, i));
        retval.Add(new IM_PriceCode(i.ItemNumber, "H", "M", 99999999, 10, i));
      }

      if (i.IsExportRequiredWholesalePriceLevels)
      {
        //i.FillProductLine(txn);
        fillProductLine(i);
        //i.ProductLine.FillWholesalePriceLevels(txn);
        fillWholesalePriceLevels(i);
        foreach (var l in i.ProductLine.WholesalePriceLevels)
        {
          if (l.LevelLetter() == "A" || l.LevelLetter() == "B")
          {
            retval.Add(new IM_PriceCode(i.ItemNumber, l.LevelLetter(), "M", 99999999, l.Percentage, i));
          }
        }
      }
      return retval;
    }*/

    /*public IM_PriceCode(string itemCode, string customerPriceLevel, string pricingMethod, int breakQuantity1, decimal discountMarkup1, TP.Object.ItemPrice originalObject)
    {
      this.ItemCode = itemCode;
      this.CustomerPriceLevel = customerPriceLevel;
      this.PricingMethod = pricingMethod;
      this.BreakQuantity1 = breakQuantity1;
      this.DiscountMarkup1 = discountMarkup1;
      this.BreakQuantity2 = 0;
      this.DiscountMarkup2 = 0;
      this.BreakQuantity3 = 0;
      this.DiscountMarkup3 = 0;
      this.BreakQuantity4 = 0;
      this.DiscountMarkup4 = 0;
      this.BreakQuantity5 = 0;
      this.DiscountMarkup5 = 0;
      this.OriginalObject = originalObject;
    }
    public IM_PriceCode() { }*/
    #endregion

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0}",
          this.csvEscape(this.ItemCode),
          this.csvEscape(this.CustomerPriceLevel),
          this.csvEscape(this.PricingMethod),
          this.csvEscape(this.BreakQuantity1),
          this.csvEscape(this.DiscountMarkup1),
          this.csvEscape(this.BreakQuantity2),
          this.csvEscape(this.DiscountMarkup2),
          this.csvEscape(this.BreakQuantity3),
          this.csvEscape(this.DiscountMarkup3),
          this.csvEscape(this.BreakQuantity4),
          this.csvEscape(this.DiscountMarkup4),
          this.csvEscape(this.BreakQuantity5),
          this.csvEscape(this.DiscountMarkup5)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("ItemCode", this.ItemCode, r.Field<string>("ItemCode"));
      yield return new Tuple<string, object, object>("CustomerPriceLevel", this.CustomerPriceLevel, r.Field<string>("CustomerPriceLevel"));
      yield return new Tuple<string, object, object>("PricingMethod", this.PricingMethod, r.Field<string>("PricingMethod"));
      yield return new Tuple<string, object, object>("BreakQuantity1", this.BreakQuantity1, r.Field<string>("BreakQuantity1"));
      yield return new Tuple<string, object, object>("DiscountMarkup1", this.DiscountMarkup1, r.Field<string>("DiscountMarkup1"));
      yield return new Tuple<string, object, object>("BreakQuantity2", this.BreakQuantity2, r.Field<string>("BreakQuantity2"));
      yield return new Tuple<string, object, object>("DiscountMarkup2", this.DiscountMarkup2, r.Field<string>("DiscountMarkup2"));
      yield return new Tuple<string, object, object>("BreakQuantity3", this.BreakQuantity3, r.Field<string>("BreakQuantity3"));
      yield return new Tuple<string, object, object>("DiscountMarkup3", this.DiscountMarkup3, r.Field<string>("DiscountMarkup3"));
      yield return new Tuple<string, object, object>("BreakQuantity4", this.BreakQuantity4, r.Field<string>("BreakQuantity4"));
      yield return new Tuple<string, object, object>("DiscountMarkup4", this.DiscountMarkup4, r.Field<string>("DiscountMarkup4"));
      yield return new Tuple<string, object, object>("BreakQuantity5", this.BreakQuantity5, r.Field<string>("BreakQuantity5"));
      yield return new Tuple<string, object, object>("DiscountMarkup5", this.DiscountMarkup5, r.Field<string>("DiscountMarkup5"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<ItemPrice> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(ItemPrice current)
    {
      if (current.Item == null)
      {
        throw new ApplicationException("Item must be filled for IM_PriceCode.Recompare()");
      }

      if (this.ItemCode != current.Item.ItemNumber) { return false; }
      if (this.CustomerPriceLevel != getCustomerPriceLevel(current.SalesChannelID)) { return false; }
      if (this.PricingMethod != getPricingMethod()) { return false; }
      if (this.BreakQuantity1 != current.MaximumQuantity1) { return false; }
      if (this.DiscountMarkup1 != current.PriceValue1) { return false; }
      if (this.BreakQuantity2 != current.MaximumQuantity2) { return false; }
      if (this.DiscountMarkup2 != current.PriceValue2) { return false; }
      if (this.BreakQuantity3 != current.MaximumQuantity3) { return false; }
      if (this.DiscountMarkup3 != current.PriceValue3) { return false; }
      if (this.BreakQuantity4 != current.MaximumQuantity4) { return false; }
      if (this.DiscountMarkup4 != current.PriceValue4) { return false; }
      if (this.BreakQuantity5 != current.MaximumQuantity5) { return false; }
      if (this.DiscountMarkup5 != current.PriceValue5) { return false; }

      return true;
    }

  }
}
