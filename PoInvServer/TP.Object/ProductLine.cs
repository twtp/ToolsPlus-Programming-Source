using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object {

  [DataContract(Name = "BaseProductLine", Namespace = "TP.Object")]
  public class BaseProductLine : BaseBusinessObject
  {

    #region Constructors

    protected BaseProductLine() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ProductLineID { get; protected set; }

    [DataMember]
    [FieldName("ProductLine")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string ProductLinePrefix { get; protected set; }

    [DataMember]
    [FieldName("ManufFullName")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(32)]
    public string ProductLineName { get; protected set; }

    [DataMember]
    [FieldName("Hide")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsDeleted { get; protected set; }

    #endregion
  }

  [DataContract(Name = "ProductLine", Namespace = "TP.Object")]
  public class ProductLine : BaseProductLine, IExportableToMAS {

    #region Constructors

    protected ProductLine() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("PrimaryVendorNumber")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(7)]
    public string DefaultVendorNo { get; protected set; }

    [DataMember]
    [FieldName("Coop")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsCoopVendor { get; protected set; }

    [DataMember]
    [FieldName("RealVendor")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string CoopVendor { get; protected set; }

    [DataMember]
    [FieldName("ManufFullNameCleaned")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(32)]
    public string ProductLineNameCleaned { get; protected set; }

    [DataMember]
    [FieldName("ManufFullNameWeb")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(32)]
    public string ProductLineNameWeb { get; protected set; }

    [DataMember]
    [FieldName("DefaultDropshippable")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool NewItemsDefaultDropshippable { get; protected set; }

    [DataMember]
    [FieldName("BOBelow")]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal DropshippableItemsDefaultToBackorderThreshold { get; protected set; }

    [DataMember]
    [FieldName("AvailLimitDefault")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int NewItemsDefaultAvailabilityLimit { get; protected set; }

    [DataMember]
    [FieldName("StdLeadTime")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int StandardLeadTime { get; protected set; }

    [DataMember]
    [FieldName("NewChk")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    [DataMember]
    [FieldName("Repair")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Text)]
    public string RepairNotes { get; protected set; }

    [DataMember]
    [FieldName("Warranty")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Text)]
    public string WarrantyNotes { get; protected set; }

    [DataMember]
    [FieldName("CompanyInfo")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Text)]
    public string CompanyInfoNotes { get; protected set; }

    [DataMember]
    [FieldName("ServiceCenters")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Text)]
    public string ServiceCenterNotes { get; protected set; }

    [DataMember]
    [FieldName("CompanyAddress1")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string CompanyAddress1 { get; protected set; }

    [DataMember]
    [FieldName("CompanyAddress2")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string CompanyAddress2 { get; protected set; }

    [DataMember]
    [FieldName("CompanyCity")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string CompanyCity { get; protected set; }

    [DataMember]
    [FieldName("CompanyState")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(16)]
    public string CompanyState { get; protected set; }

    [DataMember]
    [FieldName("CompanyZip")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(10)]
    public string CompanyPostalCode { get; protected set; }

    [DataMember]
    [FieldName("CompanyCountry")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string CompanyCountry { get; protected set; }

    [DataMember]
    [FieldName("MinimumOrderAmount")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string MinimumOrderAmount { get; protected set; }

    [DataMember]
    [FieldName("PrepaidAmount")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string PrepaidAmount { get; protected set; }

    [DataMember]
    [FieldName("PrepaidSpecialAmount")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string PrepaidSpecialAmount { get; protected set; }

    [DataMember]
    [FieldName("ShipsFromZip")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(10)]
    public string ShipsFromPostalCode { get; protected set; }

    [DataMember]
    [FieldName("ReconditionedURL")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(256)]
    public string ReconditionedURL { get; protected set; }

    [DataMember]
    [FieldName("WebSectionID")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? WebSectionID { get; protected set; }

    [DataMember]
    [FieldName("MappViolationWarning")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool DisplayMappViolationWarning { get; protected set; }

    #endregion


    #region Links

    // ProductLine has-one Vendor
    protected Vendor _vendor;
    [DataMember(EmitDefaultValue = false)]
    public Vendor DefaultVendor { get { return this._vendor; } private set { this._vendor = value; } }
    public virtual bool FillDefaultVendor() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many WholesalePriceLevel
    protected IEnumerable<ProductLineWholesalePriceLevel> _wholesalepricelevels;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineWholesalePriceLevel> WholesalePriceLevels { get { return this._wholesalepricelevels; } private set { this._wholesalepricelevels = value; } }
    public virtual bool FillWholesalePriceLevels() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many ContactInfo
    protected IEnumerable<ProductLineContactInfo> _contacts;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineContactInfo> ContactInfo { get { return this._contacts; } private set { this._contacts = value; } }
    public virtual bool FillContactInfo() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many DropshipCharge
    protected IEnumerable<ProductLineDropshipCharge> _dscharges;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineDropshipCharge> DropshipCharges { get { return this._dscharges; } private set { this._dscharges = value; } }
    public virtual bool FillDropshipCharges() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many ProductLineSalesRep
    protected IEnumerable<ProductLineSalesRep> _salesreps;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineSalesRep> SalesReps { get { return this._salesreps; } private set { this._salesreps = value; } }
    public virtual bool FillSalesReps() { throw new ApplicationException("kaboom"); }

    // ProductLine has-one ProductLineSubstitution
    protected ProductLineSubstitution _sub;
    [DataMember(EmitDefaultValue = false)]
    public ProductLineSubstitution PurchasingSubstituteProductLine { get { return this._sub; } private set { this._sub = value; } }
    public virtual bool FillPurchasingSubstituteProductLine() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many ProductLineSubstitution
    protected IEnumerable<ProductLineSubstitution> _subincoming;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineSubstitution> IncomingSubstitutedProductLines { get { return this._subincoming; } private set { this._subincoming = value; } }
    public virtual bool FillIncomingSubstitutedProductLines() { throw new ApplicationException("kaboom"); }

    // ProductLine has-many ProductLineCostField
    protected IEnumerable<ProductLineCostField> _costfields;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineCostField> CostFields { get { return this._costfields; } private set { this._costfields = value; } }
    public virtual bool FillCostFields() { throw new ApplicationException("kaboom"); }

    #endregion


    public bool IsValidPrefix() { return IsValidPrefix(this.ProductLineName); }
    public static bool IsValidPrefix(string productLine)
    {
      return AttributeHelper.Validate<ProductLine>("ProductLinePrefix", productLine) && 
             System.Text.RegularExpressions.Regex.IsMatch(productLine, @"^[-A-Z0-9]+$");
    }

    public override string ToString()
    {
      return string.Format("{0} - {1}", this.ProductLinePrefix, this.ProductLineNameWeb ?? "<no name>");
    }


    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }

  }
}
