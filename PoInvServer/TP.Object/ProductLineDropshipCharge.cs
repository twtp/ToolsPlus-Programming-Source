using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "ProductLineDropshipCharge", Namespace = "TP.Object")]
  public class ProductLineDropshipCharge : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineDropshipCharge() { }

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
    [FieldName("AddressType")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public AddressZoningType AddressZoningType { get; protected set; }

    [DataMember]
    [FieldName("LiftGateCharge")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal? LiftGateCharge { get; protected set; }

    [DataMember]
    [FieldName("LiftGateWeightThreshold")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? LiftGateWeightThreshold { get; protected set; }

    [DataMember]
    [FieldName("AdderAmount")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal? AdderAmount { get; protected set; }

    [DataMember]
    [FieldName("AdderType")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public DropshipChargeCalculationType AdderCalculationType { get; protected set; }

    [DataMember]
    [FieldName("AdderThreshold")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal? AdderThreshold { get; protected set; }

    [DataMember]
    [FieldName("AdderMaximum")]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal? AdderMaximum { get; protected set; }

    #endregion


    #region Links

    // ProductLineDropshipCharge has-many ProductLineDropshipMatrixLine (in the form of a ProductLineDropshipMatrix)
    protected ProductLineDropshipMatrix _matrix;
    [DataMember(EmitDefaultValue = false)]
    public ProductLineDropshipMatrix ChargeMatrix { get { return this._matrix; } private set { this._matrix = value; } }
    public virtual bool FillChargeMatrix() { throw new ApplicationException("kaboom"); }

    #endregion


    public decimal CalculatedAmount(decimal cost, decimal totalWeight)
    {
      decimal totalCharge = 0;

      if (this.LiftGateCharge.HasValue) {
        if (this.LiftGateWeightThreshold.HasValue == false || totalWeight < this.LiftGateWeightThreshold.Value) {
          totalCharge += this.LiftGateCharge.Value;
        }
      }

      if (this.AdderAmount.HasValue)
      {
        if (this.AdderThreshold.HasValue == false || cost < this.AdderThreshold.Value)
        {
          decimal tempAmt;
          switch (this.AdderCalculationType)
          {
            case DropshipChargeCalculationType.Flat:
              tempAmt = this.AdderAmount.Value;
              break;
            case DropshipChargeCalculationType.Percentage:
              tempAmt = decimal.Round(this.AdderAmount.Value / 100 * cost, 2, MidpointRounding.AwayFromZero);
              break;
            default:
              throw new ApplicationException("DropshipChargeCalculationType missing enum in switch()");
          }
          if (this.AdderMaximum.HasValue && tempAmt > this.AdderMaximum.Value)
          {
            tempAmt = this.AdderMaximum.Value;
          }
          totalCharge += tempAmt;
        }
      }

      // charge matrix is always filled in, so this is okay
      if (this.ChargeMatrix.Count() > 0)
      {
        totalCharge += this.ChargeMatrix.CalculatedAmount(cost);
      }

      return totalCharge;
    }

  }
}
