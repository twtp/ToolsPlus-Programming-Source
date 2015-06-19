using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ProductLineWholesalePriceLevel", Namespace = "TP.Object")]
  public class ProductLineWholesalePriceLevel : BaseBusinessObject
  {
    #region Constructors

    protected ProductLineWholesalePriceLevel() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ProductLineID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ProductLineID { get; protected set; }

    [DataMember]
    [FieldName("WholesalePriceLevel")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    [IntegerMaximum(6)]
    public int Level { get; protected set; }
    

    [DataMember]
    [FieldName("Percentage")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int Percentage { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion


    public string LevelLetter() {
      return System.Text.Encoding.UTF8.GetString(new byte[] { (byte)(65 + this.Level) });
    }

    public override string ToString()
    {
      return string.Format("{0} := {1}", this.LevelLetter(), this.Percentage);
    }
  }
}
