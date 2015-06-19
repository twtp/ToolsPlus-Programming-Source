using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemQuantity", Namespace = "TP.Object")]
  public class ItemQuantity : BaseBusinessObject
  {
    #region Constructors

    protected ItemQuantity() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ItemID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    [DataMember]
    [FieldName("WhseCode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string WarehouseCode { get; protected set; }

    [DataMember]
    [FieldName("QuantityOnHand")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int OnHand { get; protected set; }

    [DataMember]
    [FieldName("QuantityOnPurchaseOrder")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int OnPurchaseOrder { get; protected set; }

    [DataMember]
    [FieldName("QuantityOnSalesOrder")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int OnSalesOrder { get; protected set; }

    [DataMember]
    [FieldName("QuantityOnBackOrder")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int OnBackOrder { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion


    public bool IsPhysicalWarehouse { get { return this.WarehouseCode.StartsWith("0"); } }

    public override string ToString()
    {
      return string.Format("{0} => {1}", this.WarehouseCode, this.OnHand);
    }
  }
}
