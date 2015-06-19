using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "PurchaseOrder", Namespace = "TP.Object")]
  public class PurchaseOrder : BaseBusinessObject, IExportableToMAS
  {
    #region Constructors

    protected PurchaseOrder() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int PurchaseOrderID { get; protected set; }

    [DataMember]
    [FieldName("PONumber")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    public string PurchaseOrderNo { get; protected set; }

    [DataMember]
    [FieldName("POTerms")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string PaymentTerms { get; protected set; } // TODO: PaymentTerms object?

    [DataMember]
    [FieldName("VendorNumber")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string VendorNo { get; protected set; } // TODO: Vendor object?

    [DataMember]
    [FieldName("CoopVendor")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string CoopVendorNo { get; protected set; }

    [DataMember]
    [FieldName("ProductLineCode")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string ProductLine { get; protected set; }

    [DataMember]
    [FieldName("DueDate")]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime DueDate { get; protected set; }

    [DataMember]
    [FieldName("FreightFree")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsFreightFree { get; protected set; }

    [DataMember]
    [FieldName("Finalized")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsFinalized { get; protected set; }

    [DataMember]
    [FieldName("Exported")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsExported { get; protected set; }

    [DataMember]
    [FieldName("DateExported")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime? ExportedDate { get; protected set; }

    [DataMember]
    [FieldName("POComment")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(255)]
    public string Comment { get; protected set; }

    [DataMember]
    [FieldName("ShipToCode")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(4)]
    public string ShipToCode { get; protected set; }

    [DataMember]
    [FieldName("ShipToName")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShipToName { get; protected set; }

    [DataMember]
    [FieldName("ShipToAddress1")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShipToAddress1 { get; protected set; }

    [DataMember]
    [FieldName("ShipToAddress2")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShipToAddress2 { get; protected set; }

    [DataMember]
    [FieldName("ShipToCity")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShipToCity { get; protected set; }

    [DataMember]
    [FieldName("ShipToState")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string ShipToState { get; protected set; }

    [DataMember]
    [FieldName("ShipToZip")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShipToPostalCode { get; protected set; }

    [DataMember]
    [FieldName("VendorPOAddr")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string VendorPOAddress { get; protected set; }

    [DataMember]
    [FieldName("NumberOfPayments")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    public int NumberOfPayments { get; protected set; }

    [DataMember]
    [FieldName("MiscNotes")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Text)]
    public string MiscellaneousNotes { get; protected set; }

    [DataMember]
    [FieldName("SalesOrderNo")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string DropshipSalesOrderNo { get; protected set; }

    [DataMember]
    [FieldName("ShipToTelephoneNo")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string DropshipShipToTelephoneNo { get; protected set; }

    [DataMember]
    [FieldName("CustomerEmailAddress")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(35)]
    public string DropshipCustomerEmailAddress { get; protected set; }

    [DataMember]
    [FieldName("LiftGateRequired")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool? DropshipIsLiftGateRequired { get; protected set; }

    [DataMember]
    [FieldName("AddressZoningType")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? DropshipAddressZoningType { get; protected set; }

    [DataMember]
    [FieldName("VendorOrderNo")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(15)]
    public string VendorOrderNo { get; protected set; }

    [DataMember]
    [FieldName("ShippingStatus")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int VendorShippingStatus { get; protected set; }

    #endregion


    #region Links

    // PurchaseOrder has-one Vendor
    protected Vendor _vendor;
    [DataMember(EmitDefaultValue = false)]
    public Vendor Vendor { get { return this._vendor; } private set { this._vendor = value; } }
    public virtual bool FillVendor() { throw new ApplicationException("kaboom"); }

    // PurchaseOrder has-many PurchaseOrderItem
    protected IEnumerable<PurchaseOrderItem> _items;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<PurchaseOrderItem> PurchaseOrderItems { get { return this._items; } private set { this._items = value; } }
    public virtual bool FillPurchaseOrderItems() { throw new ApplicationException("kaboom"); }

    // PurchaseOrder has-many PurchaseOrderTracking
    protected IEnumerable<PurchaseOrderTracking> _tracking;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<PurchaseOrderTracking> PurchaseOrderTracking { get { return this._tracking; } private set { this._tracking = value; } }
    public virtual bool FillPurchaseOrderTracking() { throw new ApplicationException("kaboom"); }

    #endregion




    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }
  }
}
