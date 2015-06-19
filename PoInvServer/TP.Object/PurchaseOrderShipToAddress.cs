using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "PurchaseOrderShipToAddress", Namespace = "TP.Object")]
  public class PurchaseOrderShipToAddress : BaseBusinessObject
  {
    #region Constructors

    protected PurchaseOrderShipToAddress() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ShipToCode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(4)]
    public string ShipToCode { get; protected set; }

    [DataMember]
    [FieldName("ShipToCodeDesc")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ShipToCodeDesc { get; protected set; }

    [DataMember]
    [FieldName("ShipToName")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ShipToName { get; protected set; }

    [DataMember]
    [FieldName("ShipToAddress1")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ShipToAddress1 { get; protected set; }

    [DataMember]
    [FieldName("ShipToAddress2")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ShipToAddress2 { get; protected set; }

    [DataMember]
    [FieldName("ShipToAddress3")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ShipToAddress3 { get; protected set; }

    [DataMember]
    [FieldName("ShipToCity")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string ShipToCity { get; protected set; }

    [DataMember]
    [FieldName("ShipToState")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string ShipToState { get; protected set; }

    [DataMember]
    [FieldName("ShipToZipCode")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(10)]
    public string ShipToPostalCode { get; protected set; }

    [DataMember]
    [FieldName("ShipToCountryCode")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string ShipToCountryCode { get; protected set; }

    [DataMember]
    [FieldName("ExportRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion
  }
}
