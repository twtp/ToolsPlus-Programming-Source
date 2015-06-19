using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "PurchaseTerms", Namespace = "TP.Object")]
  public class PurchaseTerms : BaseBusinessObject, IExportableToMAS
  {
    #region Constructors

    protected PurchaseTerms() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("TermsCode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string Code { get; protected set; }

    [DataMember]
    [FieldName("TermsCodeDesc")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Description { get; protected set; }

    [DataMember]
    [FieldName("DaysBeforeDue")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int DaysBeforeDue { get; protected set; }

    [DataMember]
    [FieldName("DueDateADayOfTheMonth")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(1)]
    public string DueDateADayOfTheMonth { get; protected set; }

    [DataMember]
    [FieldName("MinimumDaysAllowedInv")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int MinimumDaysAllowedInv { get; protected set; }

    [DataMember]
    [FieldName("DaysDiscountAllowed")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int DaysDiscountAllowed { get; protected set; }

    [DataMember]
    [FieldName("DiscountDateADayOfTheMo")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(1)]
    public string DiscountDateADayOfTheMonth { get; protected set; }

    [DataMember]
    [FieldName("MinimumDaysAllowedDisc")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int MinimumDaysAllowedDisc { get; protected set; }

    [DataMember]
    [FieldName("DiscountRate")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    public decimal DiscountRate { get; protected set; }

    [DataMember]
    [FieldName("ExportRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion

    public static string DefaultTermsCode()
    {
      return "00";
    }

    public override string ToString()
    {
      return this.Code;
    }


    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }

  }
}
