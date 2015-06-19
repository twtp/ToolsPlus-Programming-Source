using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "ProductLineSubstitution", Namespace = "TP.Object")]
  public class ProductLineSubstitution : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineSubstitution() { }

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
    [FieldName("PurchasingProductLineID")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int PurchasingProductLineID { get; protected set; }

    #endregion


    #region Links

    // ProductLineSubstitution has-one ProductLine
    protected ProductLine _pl;
    [DataMember(EmitDefaultValue = false)]
    public ProductLine ProductLine { get { return this._pl; } private set { this._pl = value; } }
    public virtual bool FillProductLine() { throw new ApplicationException("kaboom"); }

    // ProductLineSubstitution has-one ProductLine
    protected ProductLine _ppl;
    [DataMember(EmitDefaultValue = false)]
    public ProductLine PurchasingProductLine { get { return this._ppl; } private set { this._ppl = value; } }
    public virtual bool FillPurchasingProductLine() { throw new ApplicationException("kaboom"); }

    #endregion

    public override string ToString()
    {
      if (this._pl == null || this._ppl == null)
      {
        return string.Format("{0} => {1}", this.ProductLineID, this.PurchasingProductLineID);
      }
      else
      {
        return string.Format("{0} => {1}", this.ProductLine.ProductLinePrefix, this.PurchasingProductLine.ProductLinePrefix);
      }
    }
  }
}
