using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemFreightActual", Namespace = "TP.Object")]
  public class ItemFreightActual : BaseBusinessObject
  {
    #region Constructors

    protected ItemFreightActual() { }

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
    [FieldName("DateOfShipment")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime ShipmentDate { get; protected set; }

    [DataMember]
    [FieldName("DateOfSale")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime SaleDate { get; protected set; }

    [DataMember]
    [FieldName("SalesOrderNumber")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    public string SalesOrderNo { get; protected set; }

    [DataMember]
    [FieldName("ItemUnitPrice")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal ItemUnitPrice { get; protected set; }

    #endregion


  }
}
