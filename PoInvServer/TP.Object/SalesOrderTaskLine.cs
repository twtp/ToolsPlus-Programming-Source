using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "SalesOrderTaskLine", Namespace = "TP.Object")]
  public class SalesOrderTaskLine : BaseBusinessObject
  {
    #region Constructors

    protected SalesOrderTaskLine() { }

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
    public int SalesOrderTaskID { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("QuantityOrdered")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityOrdered { get; protected set; }

    [DataMember]
    [FieldName("QuantityBackordered")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityBackordered { get; protected set; }

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
    [FieldName("QuantityShipped")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int QuantityShipped { get; protected set; }

    #endregion


    #region Links

    // SalesOrderTaskLine has-one SalesOrderTask
    protected SalesOrderTask _salesorder;
    [DataMember(EmitDefaultValue = false)]
    public SalesOrderTask SalesOrderTask { get { return this._salesorder; } private set { this._salesorder = value; } }
    public virtual bool FillSalesOrderTask() { throw new ApplicationException("kaboom"); }

    // SalesOrderTaskLine has-one Component
    protected Component _component;
    [DataMember(EmitDefaultValue = false)]
    public Component Component { get { return this._component; } private set { this._component = value; } }
    public virtual bool FillComponent() { throw new ApplicationException("kaboom"); }

    #endregion


    public override string ToString()
    {
      return string.Format("TaskID: {0} => {1}x {2}", this.SalesOrderTaskID, this.QuantityOrdered, this.ComponentID);
    }

  }
}
