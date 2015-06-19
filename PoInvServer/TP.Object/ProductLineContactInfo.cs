using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ProductLineContactInfo", Namespace = "TP.Object")]
  public class ProductLineContactInfo : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineContactInfo() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ProductLineID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ProductLineID { get; protected set; }

    [DataMember]
    [FieldName("InfoTypeID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public ProductLineContactInfoType InfoType { get; protected set; }

    [DataMember]
    [FieldName("Phone")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(128)]
    public string TelephoneNo { get; protected set; }

    [DataMember]
    [FieldName("Fax")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(128)]
    public string FaxNo { get; protected set; }

    [DataMember]
    [FieldName("URL")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(128)]
    public string HomepageURL { get; protected set; }

    [DataMember]
    [FieldName("Email")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(128)]
    public string EmailAddress { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion

    public bool IsValidTelephoneNo() { return IsValidTelephoneNo(this.TelephoneNo); }
    public static bool IsValidTelephoneNo(string telephoneNo)
    {
      return AttributeHelper.Validate<ProductLineContactInfo>("TelephoneNo", telephoneNo);
    }

    public bool IsValidFaxNo() { return IsValidFaxNo(this.FaxNo); }
    public static bool IsValidFaxNo(string faxNo)
    {
      return AttributeHelper.Validate<ProductLineContactInfo>("FaxNo", faxNo);
    }

    public bool IsValidUrl() { return IsValidUrl(this.HomepageURL); }
    public static bool IsValidUrl(string url)
    {
      return AttributeHelper.Validate<ProductLineContactInfo>("HomepageURL", url);
    }

    public bool IsValidEmailAddress() { return IsValidEmailAddress(this.EmailAddress); }
    public static bool IsValidEmailAddress(string emailAddress)
    {
      return AttributeHelper.Validate<ProductLineContactInfo>("EmailAddress", emailAddress);
    }
  }
}
