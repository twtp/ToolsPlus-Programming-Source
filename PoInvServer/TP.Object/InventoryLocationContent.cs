using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "InventoryLocationContent", Namespace = "TP.Object")]
  public class InventoryLocationContent : BaseBusinessObject
  {
    #region Constructors

    protected InventoryLocationContent() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("LocationID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int LocationID { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("MasterID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int MasterID { get; protected set; }

    [DataMember]
    [FieldName("Quantity")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? Quantity { get; protected set; }

    [DataMember]
    [FieldName("AssociatedTaskType")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? AssociatedTaskType { get; protected set; }

    [DataMember]
    [FieldName("AssociatedTaskID")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? AssociatedTaskID { get; protected set; }

    [DataMember]
    [FieldName("LastInventoriedDate")]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime LastInventoriedDate { get; protected set; }

    #endregion


    #region Links

    // InventoryLocationContent has-one Component
    protected Component _component;
    [DataMember(EmitDefaultValue = false)]
    public Component Component { get { return this._component; } private set { this._component = value; } }
    public virtual bool FillComponent() { throw new ApplicationException("kaboom"); }

    // InventoryLocationContent has-one InventoryLocation
    protected InventoryLocation _location;
    [DataMember(EmitDefaultValue = false)]
    public InventoryLocation InventoryLocation { get { return this._location; } private set { this._location = value; } }
    public virtual bool FillInventoryLocation() { throw new ApplicationException("kaboom"); }

    #endregion

    public override string ToString()
    {
      return string.Format("{0}x CID {1}", this.Quantity.HasValue ? this.Quantity.Value.ToString() : "SPAN", this.ComponentID);
    }
  }
}
