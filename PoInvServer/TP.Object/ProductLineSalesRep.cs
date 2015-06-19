using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "ProductLineSalesRep", Namespace = "TP.Object")]
  public class ProductLineSalesRep : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineSalesRep() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TupleID { get; protected set; }

    [DataMember]
    [FieldName("ProductLineID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ProductLineID { get; protected set; }

    [DataMember]
    [FieldName("SalesRepID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SalesRepID { get; protected set; }

    [DataMember]
    [FieldName("IsPrimary")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsPrimary { get; protected set; }

    #endregion


    #region Links

    // ProductLineSalesRep has-one ProductLine
    protected ProductLine _productline;
    [DataMember(EmitDefaultValue = false)]
    public ProductLine ProductLine { get { return this._productline; } private set { this._productline = value; } }
    public virtual bool FillProductLine() { throw new ApplicationException("kaboom"); }

    // ProductLineSalesRep has-one SalesRep
    protected SalesRep _salesrep;
    [DataMember(EmitDefaultValue = false)]
    public SalesRep SalesRep { get { return this._salesrep; } private set { this._salesrep = value; } }
    public virtual bool FillSalesRep() { throw new ApplicationException("kaboom"); }

    #endregion


  }
}
