using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "PurchaseOrderItem", Namespace = "TP.Object")]
  public class PurchaseOrderItem : BaseBusinessObject
  {
    #region Constructors

    protected PurchaseOrderItem() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TupleID { get; protected set; }

    [DataMember]
    [FieldName("HeaderID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int HeaderID { get; protected set; }

    [DataMember]
    [FieldName("ItemID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    [DataMember]
    [FieldName("Quantity")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    public int Quantity { get; protected set; }

    #endregion


    #region Links

    // PurchaseOrderItem has-one Item
    protected Item _item;
    [DataMember(EmitDefaultValue = false)]
    public Item Item { get { return this._item; } private set { this._item = value; } }
    public virtual bool FillItem() { throw new ApplicationException("kaboom"); }

    // PurchaseOrderItem has-one PurchaseOrder
    protected PurchaseOrder _po;
    [DataMember(EmitDefaultValue = false)]
    public PurchaseOrder PurchaseOrder { get { return this._po; } private set { this._po = value; } }
    public virtual bool FillPurchaseOrder() { throw new ApplicationException("kaboom"); }

    #endregion
  }
}
