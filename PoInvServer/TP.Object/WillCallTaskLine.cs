using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "WillCallTaskLine", Namespace = "TP.Object")]
  public class WillCallTaskLine : BaseBusinessObject
  {
    #region Constructors

    protected WillCallTaskLine() { }

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
    public int WillCallTaskID { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("QuantityRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityRequired { get; protected set; }

    [DataMember]
    [FieldName("QuantityPicked")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityPicked { get; protected set; }

    [DataMember]
    [FieldName("QuantityStaged")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityPacked { get; protected set; }

    [DataMember]
    [FieldName("QuantityFilled")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityFilled { get; protected set; }

    #endregion


    #region Links

    // WillCallTaskLine has-one WillCallTask
    protected WillCallTask _willcall;
    [DataMember(EmitDefaultValue = false)]
    public WillCallTask WillCallTask { get { return this._willcall; } private set { this._willcall = value; } }
    public virtual bool FillWillCallTask() { throw new ApplicationException("kaboom"); }

    // WillCallTaskLine has-one Component
    protected Component _component;
    [DataMember(EmitDefaultValue = false)]
    public Component Component { get { return this._component; } private set { this._component = value; } }
    public virtual bool FillComponent() { throw new ApplicationException("kaboom"); }

    #endregion


    public override string ToString()
    {
      return string.Format("TaskID: {0} => {1}x {2}", this.WillCallTaskID, this.QuantityRequired, this.ComponentID);
    }

  }
}
