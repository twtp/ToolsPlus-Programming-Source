using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemFreightEstimate", Namespace = "TP.Object")]
  public class ItemFreightEstimate : BaseBusinessObject
  {
    #region Constructors

    protected ItemFreightEstimate() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TupleID { get; protected set; }

    [DataMember]
    [FieldName("ItemID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    [DataMember]
    [FieldName("ZipCode")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    public string PostalCode { get; protected set; }

    [DataMember]
    [FieldName("Cost")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Cost { get; protected set; }

    [DataMember]
    [FieldName("Service")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    public string Service { get; protected set; }

    [DataMember]
    [FieldName("LastUpdated")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime LastUpdated { get; protected set; }

    #endregion

  }
}
