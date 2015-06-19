using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "TransferTask", Namespace = "TP.Object")]
  public class TransferTask : BaseBusinessObject
  {

    #region Constructors

    protected TransferTask() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TransferTaskID { get; protected set; }

    [DataMember]
    [FieldName("InitialDate")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime InitialDate { get; protected set; }

    [DataMember]
    [FieldName("SequenceNumber")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SequenceNo { get; protected set; }

    [DataMember]
    [FieldName("TransferNumber")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(9)]
    public string TransferNo { get; protected set; }

    [DataMember]
    [FieldName("TransferSubtype")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TransferSubtype { get; protected set; }

    [DataMember]
    [FieldName("WarehouseID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int WarehouseID { get; protected set; }

    [DataMember]
    [FieldName("DestinationWarehouseID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int DestinationWarehouseID { get; protected set; }

    [DataMember]
    [FieldName("IsComplete")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsComplete { get; protected set; }

    [DataMember]
    [FieldName("Description")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    public string Description { get; protected set; }

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

    // TransferTask has-many TransferTaskLine
    protected IEnumerable<TransferTaskLine> _lines;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<TransferTaskLine> Lines { get { return this._lines; } private set { this._lines = value; } }
    public virtual bool FillLines() { throw new ApplicationException("kaboom"); }

    #endregion

    public override string ToString()
    {
      return this.TransferNo;
    }

  }
}
