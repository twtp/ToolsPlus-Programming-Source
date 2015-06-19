using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "Component", Namespace = "TP.Object")]
  public class Component : BaseBusinessObject
  {
    #region Constructors

    protected Component() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("Component")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(32)]
    public string ComponentName { get; protected set; }

    [DataMember]
    [FieldName("Length")]
    //[SqlDbType(System.Data.SqlDbType.VarChar)]
    //[StringMaximumLength(8)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Length { get; protected set; }

    [DataMember]
    [FieldName("Width")]
    //[SqlDbType(System.Data.SqlDbType.VarChar)]
    //[StringMaximumLength(8)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Width { get; protected set; }

    [DataMember]
    [FieldName("Height")]
    //[SqlDbType(System.Data.SqlDbType.VarChar)]
    //[StringMaximumLength(8)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Height { get; protected set; }

    [DataMember]
    [FieldName("Weight")]
    //[SqlDbType(System.Data.SqlDbType.VarChar)]
    //[StringMaximumLength(8)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Weight { get; protected set; }

    [DataMember]
    [FieldName("PackingType")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int PackingType { get; protected set; } // this should probably be an enum

    // skip ShippingBoxID, not really used

    [DataMember]
    [FieldName("GiftNotConcealableWarning")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsNotConcealableDuringShipping { get; protected set; }

    [DataMember]
    [FieldName("HazMat")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsHazardous { get; protected set; }

    [DataMember]
    [FieldName("TrackSerialNo")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsTrackedBySerialNo { get; protected set; }

    [DataMember]
    [FieldName("SpecificUpSide")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsSpecificSideUp { get; protected set; }

    // skip NMFC

    // skip PreferredPickLocation

    // skip WtDimsLastModifiedTime

    #endregion


    #region Links

    // Component has-many Barcode
    protected IEnumerable<Barcode> _barcodes;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<Barcode> Barcodes { get { return this._barcodes; } private set { this._barcodes = value; } }
    public virtual bool FillBarcodeList() { throw new ApplicationException("kaboom"); }

    // Component has-many InventoryLocationContent
    protected IEnumerable<InventoryLocationContent> _locations;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<InventoryLocationContent> Locations { get { return this._locations; } private set { this._locations = value; } }
    public virtual bool FillLocations() { throw new ApplicationException("kaboom"); }

    #endregion


    public void CanonicalizeDimensions()
    {
      if (this.Length < this.Width)
      {
        decimal temp = this.Length;
        this.Length = this.Width;
        this.Width = temp;
      }
      if (this.Width < this.Height)
      {
        decimal temp = this.Width;
        this.Width = this.Height;
        this.Height = temp;
      }
      if (this.Length < this.Width)
      {
        decimal temp = this.Length;
        this.Length = this.Width;
        this.Width = temp;
      }
    }

    public override string ToString()
    {
      return string.Format("{0} - {1} lbs, {2}x{3}x{4}", this.ComponentName, this.Weight, this.Length, this.Width, this.Height);
    }
  }
}
