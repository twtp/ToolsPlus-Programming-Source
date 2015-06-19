using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "SalesOrderTask", Namespace = "TP.Object")]
  public class SalesOrderTask : BaseBusinessObject
  {

    #region Constructors

    protected SalesOrderTask() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SalesOrderTaskID { get; protected set; }

    [DataMember]
    [FieldName("SalesOrderNo")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(7)]
    public string SalesOrderNo { get; protected set; }

    [DataMember]
    [FieldName("WarehouseID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int WarehouseID { get; protected set; }

    [DataMember]
    [FieldName("CurrentStatus")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int CurrentStatus { get; protected set; }

    [DataMember]
    [FieldName("CurrentHoldReason")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int CurrentHoldReason { get; protected set; }

    [DataMember]
    [FieldName("ShipVia")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(15)]
    public string ShipVia { get; protected set; }

    [DataMember]
    [FieldName("ShipState")]
    [Nullable(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string ShipState { get; protected set; }

    [DataMember]
    [FieldName("ShipCountry")]
    [Nullable(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string ShipCountry { get; protected set; }

    [DataMember]
    [FieldName("MasCreatedByUser")]
    [Nullable(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string MasCreatedByUser { get; protected set; }

    [DataMember]
    [FieldName("TimeInserted")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime TimeInserted { get; protected set; }

    #endregion


    #region Links

    // SalesOrderTask has-many SalesOrderTaskLine
    protected IEnumerable<SalesOrderTaskLine> _lines;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<SalesOrderTaskLine> Lines { get { return this._lines; } private set { this._lines = value; } }
    public virtual bool FillLines() { throw new ApplicationException("kaboom"); }

    // SalesOrderTask has-one SalesChannel
    protected SalesChannel _saleschannel;
    [DataMember(EmitDefaultValue = false)]
    public SalesChannel SalesChannel { get { return this._saleschannel; } private set { this._saleschannel = value; } }
    public virtual bool FillSalesChannel() { throw new ApplicationException("kaboom"); }

    #endregion

    //public SalesChannel SalesChannel
    //{
    //  get
    //  {
    //    if (System.Text.RegularExpressions.Regex.IsMatch(this.SalesOrderNo, @"^\d{7}$")) {
    //      
    //  }
    //}

    public override string ToString()
    {
      return this.SalesOrderNo;
    }

  }
}
