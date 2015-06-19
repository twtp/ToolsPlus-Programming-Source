using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {

  public enum AddressZoningType {
    Unknown = -1,
    Commercial = 0,
    Residential = 1,
    Agricultural = 2,
  };




  // attrs are currently unused, if we keep adding sites retardedly,
  // switch to using them. if we rewrite the db to multi-site the
  // right way, then this will unnecessary, and need to change anyway.
  public enum SellingChannel {
    [SellingPriceDatabaseField("DiscountMarkupPriceRate")]
    [SellingPriceBreaksSupported(5)]
    MeadowStRetail,
    [SellingPriceDatabaseField("IDiscountMarkupPriceRate")]
    [SellingPriceBreaksSupported(5)]
    Telephone,
    [SellingPriceDatabaseField("IDiscountMarkupPriceRate")]
    [SellingPriceBreaksSupported(5)]
    ToolsPlus,
    [SellingPriceDatabaseField("IDiscountMarkupPriceRate")]
    [SellingPriceBreaksSupported(5)]
    ToolPartsStore,
    [SellingPriceDatabaseField("EBayPrice")]
    [SellingPriceBreaksSupported(0)]
    ToolsPlusOutlet,
  };

  internal class SellingPriceDatabaseFieldAttribute : TP.Base.Attributes.GenericEnumAttribute {
    private string sellingPriceDatabaseField;
    public SellingPriceDatabaseFieldAttribute(string sellingPriceDatabaseField) {
      this.sellingPriceDatabaseField = sellingPriceDatabaseField;
    }
    public string SellingPriceDatabaseField { get { return this.sellingPriceDatabaseField; } }
    public override object TheAttribute { get { return this.sellingPriceDatabaseField; } }
  }

  internal class SellingPriceBreaksSupportedAttribute : TP.Base.Attributes.GenericEnumAttribute {
    private int sellingPriceBreaksSupported;
    public SellingPriceBreaksSupportedAttribute(int sellingPriceBreaksSupported) {
      this.sellingPriceBreaksSupported = sellingPriceBreaksSupported;
    }
    public int SellingPriceBreaksSupported { get { return this.sellingPriceBreaksSupported; } }
    public override object TheAttribute { get { return this.sellingPriceBreaksSupported; } }
  }

}
