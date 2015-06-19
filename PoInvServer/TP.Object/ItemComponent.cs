using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemComponent", Namespace = "TP.Object")]
  public class ItemComponent : BaseBusinessObject
  {
    #region Constructors

    protected ItemComponent() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ItemID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("Quantity")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    public int Quantity { get; protected set; }

    #endregion


    #region Links

    // ItemComponent has-one Component
    protected Component _component;
    [DataMember(EmitDefaultValue = false)]
    public Component Component { get { return this._component; } private set { this._component = value; } }
    public virtual bool FillComponent(string[] preloads = null) { throw new ApplicationException("kaboom"); }

    #endregion
  }

}
