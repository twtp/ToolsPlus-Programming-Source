using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "TransferTaskLine", Namespace = "TP.Object")]
  public class TransferTaskLine : BaseBusinessObject
  {
    #region Constructors

    protected TransferTaskLine() { }

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
    public int TransferTaskID { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("QuantityRequested")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityRequested { get; protected set; }

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
    [FieldName("QuantityTransported")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityTransported { get; protected set; }

    [DataMember]
    [FieldName("QuantityFilled")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityFilled { get; protected set; }

    #endregion


    #region Links

    // TransferTaskLine has-one TransferTask
    protected TransferTask _transfer;
    [DataMember(EmitDefaultValue = false)]
    public TransferTask TransferTask { get { return this._transfer; } private set { this._transfer = value; } }
    public virtual bool FillTransferTask() { throw new ApplicationException("kaboom"); }

    // TransferTaskLine has-one Component
    protected Component _component;
    [DataMember(EmitDefaultValue = false)]
    public Component Component { get { return this._component; } private set { this._component = value; } }
    public virtual bool FillComponent() { throw new ApplicationException("kaboom"); }

    #endregion


    public override string ToString()
    {
      return string.Format("TaskID: {0} => {1}x {2}", this.TransferTaskID, this.QuantityRequested, this.ComponentID);
    }

  }
}
