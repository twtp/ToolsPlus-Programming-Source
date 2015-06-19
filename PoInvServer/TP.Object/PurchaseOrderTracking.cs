using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "PurchaseOrderTracking", Namespace = "TP.Object")]
  public class PurchaseOrderTracking : BaseBusinessObject
  {
    #region Constructors

    protected PurchaseOrderTracking() { }

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
    [FieldName("TrackingID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string TrackingID { get; protected set; }

    [DataMember]
    [FieldName("Carrier")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string Carrier { get; protected set; }

    [DataMember]
    [FieldName("CustomerNotified")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsCustomerNotified { get; protected set; }

    #endregion


    #region Links

    // PurchaseOrderTracking has-one PurchaseOrder
    protected PurchaseOrder _po;
    [DataMember(EmitDefaultValue = false)]
    public PurchaseOrder PurchaseOrder { get { return this._po; } private set { this._po = value; } }
    public virtual bool FillPurchaseOrder() { throw new ApplicationException("kaboom"); }

    #endregion
  }
}
