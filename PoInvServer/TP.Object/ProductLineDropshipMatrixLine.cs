using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "ProductLineDropshipMatrixLine", Namespace = "TP.Object")]
  public class ProductLineDropshipMatrixLine : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineDropshipMatrixLine() { }

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
    [FieldName("ThresholdCost")]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal ThresholdCost { get; protected set; }

    [DataMember]
    [FieldName("AmountCalculationType")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public DropshipChargeCalculationType CalculationType { get; protected set; }

    [DataMember]
    [FieldName("Amount")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal Amount { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion

    public decimal CalculatedAmount(decimal itemCost)
    {
      switch (this.CalculationType)
      {
        case DropshipChargeCalculationType.Flat:
          return this.Amount;
        case DropshipChargeCalculationType.Percentage:
          return decimal.Round(this.Amount / 100 * itemCost, 2, MidpointRounding.AwayFromZero);
        default:
          throw new ApplicationException("DropshipChargeCalculationType missing enum in switch()");
      }
    }
  }

  [DataContract(Name = "ProductLineDropshipMatrix", Namespace = "TP.Object")]
  public class ProductLineDropshipMatrix : BaseBusinessObject /*, IEnumerable<ProductLineDropshipMatrixLine>*/
  {
    [DataMember(Name="ProductLineDropshipMatrix")]
    protected List<ProductLineDropshipMatrixLine> _list;

    #region Constructors

    protected ProductLineDropshipMatrix() { }

    protected ProductLineDropshipMatrix(IEnumerable<ProductLineDropshipMatrixLine> coll)
    {
      this._list = new List<ProductLineDropshipMatrixLine>(coll);
    }

    #endregion


    #region Fields

    // none

    #endregion


    #region Links

    // none

    #endregion


    public decimal CalculatedAmount(decimal itemCost)
    {
      foreach (var l in this._list)
      {
        if (itemCost < l.ThresholdCost)
        {
          return l.CalculatedAmount(itemCost);
        }
      }

      return 0m;
    }


    /*#region IEnumerable<ProductLineDropshipMatrixLine> Members

    public IEnumerator<ProductLineDropshipMatrixLine> GetEnumerator()
    {
      return this._list.GetEnumerator();
    }

    #endregion*/

    public int Count()
    {
      return this._list.Count;
    }

  }
}
