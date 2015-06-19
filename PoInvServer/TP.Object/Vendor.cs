using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "Vendor", Namespace = "TP.Object")]
  public class Vendor : BaseBusinessObject, IExportableToMAS
  {
    #region Constructors

    protected Vendor() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("APDivisionNo")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string DivisionNo { get; set; }

    [DataMember]
    [FieldName("VendorNo")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMinimumLength(3)]
    [StringMaximumLength(7)]
    public string VendorNo { get; protected set; }

    [DataMember]
    [FieldName("VendorName")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Name { get; protected set; }

    [DataMember]
    [FieldName("AddressLine1")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string AddressLine1 { get; protected set; }

    [DataMember]
    [FieldName("AddressLine2")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string AddressLine2 { get; protected set; }

    [DataMember]
    [FieldName("City")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string City { get; protected set; }

    [DataMember]
    [FieldName("State")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string State { get; protected set; }

    [DataMember]
    [FieldName("ZipCode")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(10)]
    public string PostalCode { get; protected set; }

    [DataMember]
    [FieldName("CountryCode")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string CountryCode { get; protected set; }

    [DataMember]
    [FieldName("TelephoneNo")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(17)]
    public string TelephoneNo { get; protected set; }

    [DataMember]
    [FieldName("TermsCode")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string DefaultPurchaseTermsCode { get; protected set; }

    [DataMember]
    [FieldName("Reference")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(15)]
    public string Reference { get; protected set; }

    [DataMember]
    [FieldName("ExportRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    #endregion


    #region Links

    // Vendor has-one PurchaseTerms
    protected PurchaseTerms _purchaseterms;
    [DataMember(EmitDefaultValue = false)]
    public PurchaseTerms PurchaseTerms { get { return this._purchaseterms; } private set { this._purchaseterms = value; } }
    public virtual bool FillPurchaseTerms() { throw new ApplicationException("kaboom"); }

    #endregion

    public bool IsVendorNoValid() { return IsVendorNoValid(this.VendorNo); } // insane if this returns false
    public static bool IsVendorNoValid(string vendorNo)
    {
      vendorNo = vendorNo.ToUpper();
      if (vendorNo == null)
      {
        return false;
      }
      if (false == System.Text.RegularExpressions.Regex.IsMatch(vendorNo, @"^[-_A-Z0-9]{1,7}$"))
      {
        return false;
      }
      return true;
    }

    public bool IsVendorNameValid() { return IsVendorNameValid(this.Name); }
    public static bool IsVendorNameValid(string vendorName)
    {
      return AttributeHelper.Validate<Vendor>("Name", vendorName);
    }

    public override string ToString()
    {
      return this.VendorNo;
    }


    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }

  }
}
