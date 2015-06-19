using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object {

  [DataContract(Name = "BaseItem", Namespace = "TP.Object")]
  public class BaseItem : BaseBusinessObject {

    #region Constructors

    protected BaseItem() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ItemID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    [DataMember]
    [FieldName("ItemNumber")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ItemNumber { get; protected set; }

    [DataMember]
    [FieldName("IsDiscontinued")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsDiscontinued { get; protected set; }

    #endregion

    public override string ToString()
    {
      return this.ItemNumber;
    }
  }

  [DataContract(Name = "Item", Namespace = "TP.Object")]
  public class Item : BaseItem, IExportableToMAS
  {

    #region Constructors

    protected Item() { }
    
    #endregion


    #region Fields

    [DataMember]
    [FieldName("ItemDescription")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ItemDescription { get; protected set; }

    [DataMember]
    [FieldName("PrimaryVendor")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string PrimaryVendor { get; protected set; }

    [DataMember]
    [FieldName("IsMASKit")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsMasKit { get; protected set; }

    [DataMember]
    [FieldName("EPN")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(512)]
    public string AdditionalInfo { get; protected set; }

    [DataMember]
    [FieldName("StdPack")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    public int StandardPack { get; protected set; }

    // TODO: Item *Changed flags for Mas?

    [DataMember]
    [FieldName("POComment")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(50)]
    public string PurchaseOrderComment { get; protected set; }

    [DataMember]
    [FieldName("LastReceiptDate")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    [Nullable(true)]
    public DateTime? LastReceiptDate { get; protected set; }

    [DataMember]
    [FieldName("Components")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(255)]
    public string Components { get; protected set; }

    // TODO: Item ReplacedBy/ReplacementFor?

    [DataMember]
    [FieldName("PartIntroducedDate")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.SmallDateTime)]
    public DateTime ItemIntroducedDate { get; protected set; }

    [DataMember]
    [FieldName("IsReconditioned")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsReconditioned { get; protected set; }

    [DataMember]
    [FieldName("TaxExempt")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsTaxExempt { get; protected set; }

    [DataMember]
    [FieldName("VendorQuantityOnHand")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [Nullable(true)]
    public int? VendorQuantityOnHand { get; protected set; }

    [DataMember]
    [FieldName("VendorQuantityTimeChecked")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    [Nullable(true)]
    public DateTime? VendorQuantityTimeChecked { get; protected set; }

    // TODO: other Item fields

    #endregion


    #region Links

    // Item has-one ProductLine
    protected ProductLine _productline;
    [DataMember(EmitDefaultValue = false)]
    public ProductLine ProductLine { get { return this._productline; } private set { this._productline = value; } }
    public virtual bool FillProductLine() { throw new ApplicationException("kaboom"); }

    // Item has-many ItemComponent
    protected IEnumerable<ItemComponent> _itemcomponents;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemComponent> ItemComponents { get { return this._itemcomponents; } private set { this._itemcomponents = value; } }
    public virtual bool FillItemComponents(string[] preloads = null) { throw new ApplicationException("kaboom"); }

    // Item has-many ItemCost
    protected IEnumerable<ItemCost> _itemcosts;
    [DataMember(EmitDefaultValue=false)]
    public IEnumerable<ItemCost> ItemCosts { get { return this._itemcosts; } private set { this._itemcosts = value; } }
    public virtual bool FillItemCosts() { throw new ApplicationException("kaboom"); }

    // Item has-many ItemPrice
    protected IEnumerable<ItemPrice> _itemprices;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemPrice> ItemPrices { get { return this._itemprices; } private set { this._itemprices = value; } }
    public virtual bool FillItemPrices() { throw new ApplicationException("kaboom"); }

    // Item has-many ItemQuantity
    protected IEnumerable<ItemQuantity> _itemquantities;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemQuantity> ItemQuantities { get { return this._itemquantities; } private set { this._itemquantities = value; } }
    public virtual bool FillItemQuantities() { throw new ApplicationException("kaboom"); }

    // Item has-many ItemFreightEstimate
    protected IEnumerable<ItemFreightEstimate> _itemfreightestimates;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemFreightEstimate> ItemFreightEstimates { get { return this._itemfreightestimates; } private set { this._itemfreightestimates = value; } }
    public virtual bool FillItemFreightEstimates() { throw new ApplicationException("kaboom"); }

    // Item has-many ItemFreightActual
    protected IEnumerable<ItemFreightActual> _itemfreightactuals;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemFreightActual> ItemFreightActuals { get { return this._itemfreightactuals; } private set { this._itemfreightactuals = value; } }
    public virtual bool FillItemFreightActuals() { throw new ApplicationException("kaboom"); }

    // TODO: Item has-many ItemListing
    
    // TODO: Item has-many ItemSales

    #endregion


    public bool IsItemNumberValid() { return IsItemNumberValid(this.ItemNumber); } // insane if this returns false
    public static bool IsItemNumberValid(string itemNumber)
    {
      if (itemNumber == null)
      {
        return false;
      }
      if (itemNumber.Length < 3 || itemNumber.Length > 30)
      {
        return false;
      }
      if (!itemNumber.StartsWith("XXX") && !itemNumber.StartsWith("ZZZ") && itemNumber.Length > 27)
      {
        return false;
      }
      if (!System.Text.RegularExpressions.Regex.IsMatch(itemNumber, @"^[-._A-Z0-9]+$"))
      {
        return false;
      }
      // does not check for PL existence in DB, only in CreateItem
      return true;
    }

    public string ProductLinePrefix() { return ProductLinePrefix(this.ItemNumber); }
    public static string ProductLinePrefix(string itemNumber)
    {
      if (IsTypeSpecialOrder(itemNumber) || IsTypeDropship(itemNumber))
      {
        return itemNumber.Substring(3, 3);
      }
      else
      {
        return itemNumber.Substring(0, 3);
      }
    }

    public bool IsTypeSpecialOrder() { return IsTypeSpecialOrder(this.ItemNumber); }
    public static bool IsTypeSpecialOrder(string itemNumber) {
      return itemNumber.StartsWith("XXX");
    }

    public bool IsTypeDropship() { return IsTypeDropship(this.ItemNumber); }
    public static bool IsTypeDropship(string itemNumber)
    {
      return itemNumber.StartsWith("ZZZ");
    }

    public override string ToString()
    {
      return this.ItemNumber;
    }



    // TODO: field-ify these
    public bool IsExportRequiredNewItem { get; set; }
    public bool IsExportRequiredWholesalePriceLevels { get; set; }

    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }
  }
}
