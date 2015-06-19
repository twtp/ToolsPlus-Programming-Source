using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "SalesRep", Namespace = "TP.Object")]
  public class SalesRep : BaseBusinessObject
  {

    #region Constructors

    protected SalesRep() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SalesRepID { get; protected set; }

    [DataMember]
    [FieldName("FirstName")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string FirstName { get; protected set; }

    [DataMember]
    [FieldName("LastName")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string LastName { get; protected set; }

    [DataMember]
    [FieldName("PositionTitle")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string PositionTitle { get; protected set; }

    [DataMember]
    [FieldName("Address")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string Address { get; protected set; }

    [DataMember]
    [FieldName("City")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string City { get; protected set; }

    [DataMember]
    [FieldName("State")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string State { get; protected set; }

    [DataMember]
    [FieldName("ZipCode")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string PostalCode { get; protected set; }

    [DataMember]
    [FieldName("Phone")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(17)]
    public string TelephoneNo { get; protected set; }

    [DataMember]
    [FieldName("Cell")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(17)]
    public string MobileTelephoneNo { get; protected set; }

    [DataMember]
    [FieldName("Fax")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(17)]
    public string FaxNo { get; protected set; }

    [DataMember]
    [FieldName("Email")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    [StringMaximumLength(50)]
    public string EmailAddress { get; protected set; }

    [DataMember]
    [FieldName("Notes")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [Nullable(true)]
    public string Notes { get; protected set; }

    #endregion


    #region Links

    // SalesRep has-many ProductLineSalesRep
    protected IEnumerable<ProductLineSalesRep> _productlines;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ProductLineSalesRep> ProductLinesRepped { get { return this._productlines; } private set { this._productlines = value; } }
    public virtual bool FillProductLinesRepped() { throw new ApplicationException("kaboom"); }

    #endregion

    public bool IsValidFirstName() { return IsValidFirstName(this.FirstName); }
    public static bool IsValidFirstName(string firstName)
    {
      return AttributeHelper.Validate<SalesRep>("FirstName", firstName);
    }

    public bool IsValidLastName() { return IsValidLastName(this.LastName); }
    public static bool IsValidLastName(string lastName)
    {
      return AttributeHelper.Validate<SalesRep>("LastName", lastName);
    }

    public override string ToString()
    {
      return string.Format("{0} {1}", this.FirstName ?? "?", this.LastName ?? "?");
    }
  }
}
