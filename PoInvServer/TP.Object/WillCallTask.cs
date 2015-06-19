using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "WillCallTask", Namespace = "TP.Object")]
  public class WillCallTask : BaseBusinessObject
  {

    #region Constructors

    protected WillCallTask() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int WillCallTaskID { get; protected set; }

    [DataMember]
    [FieldName("TransactionID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string MASTransactionID { get; protected set; }

    [DataMember]
    [FieldName("WarehouseID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int WarehouseID { get; protected set; }

    [DataMember]
    [FieldName("TimeInserted")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime TimeInserted { get; protected set; }

    [DataMember]
    [FieldName("CreatedByUserID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int CreatedByUserID { get; protected set; }

    #endregion


    #region Links

    // WillCallTask has-many WillCallTaskLine
    protected IEnumerable<WillCallTaskLine> _lines;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<WillCallTaskLine> Lines { get { return this._lines; } private set { this._lines = value; } }
    public virtual bool FillLines() { throw new ApplicationException("kaboom"); }

    #endregion

    public override string ToString()
    {
      return this.MASTransactionID;
    }

  }
}
