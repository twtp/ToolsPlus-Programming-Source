using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "Barcode", Namespace = "TP.Object")]
  public class Barcode : BaseBusinessObject
  {
    #region Constructors

    protected Barcode() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("Barcode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMinimumLength(12)]
    [StringMaximumLength(14)]
    public string Code { get; protected set; }

    [DataMember]
    [FieldName("ComponentID")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComponentID { get; protected set; }

    [DataMember]
    [FieldName("Internal")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsInternal { get; protected set; }

    [DataMember]
    [FieldName("Quantity")]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(1)]
    public int Quantity { get; protected set; }

    [DataMember]
    [FieldName("PrintByStdPack")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool PrintByStandardPack { get; protected set; }

    #endregion


    public bool IsValidCode() { return IsValidCode(this.Code); }
    public static bool IsValidCode(string code)
    {
      return AttributeHelper.Validate<Barcode>("Code", code) &&
             System.Text.RegularExpressions.Regex.IsMatch(code, @"^\d+$");
    }

    public override string ToString()
    {
      return string.Format("{0} => ({1}) {2}", this.Code, this.Quantity, this.ComponentID);
    }
  }
}
