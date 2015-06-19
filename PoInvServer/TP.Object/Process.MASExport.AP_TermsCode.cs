using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_AP_TermsCode", Namespace = "TP.Object")]
  public class AP_TermsCode : IMASExportObject
  {
    public static string MASColumns() { return "TermsCode, TermsCodeDesc, DaysBeforeDue, DueDateADayOfTheMonth, MinimumDaysAllowedInv, DaysDiscountAllowed, DiscountDateADayOfTheMo, MinimumDaysAllowedDisc, DiscountRate"; }
    public static string MASTableName() { return "AP_TermsCode"; }
    public string PrimaryKey() { return this.TermsCode; }

    [DataMember]
    public string TermsCode { get; set; }
    [DataMember]
    public string TermsCodeDesc { get; set; }
    [DataMember]
    public int? DaysBeforeDue { get; set; }
    [DataMember]
    public string DueDateADayOfTheMonth { get; set; }
    [DataMember]
    public int? MinimumDaysAllowedInv { get; set; }
    [DataMember]
    public int? DaysDiscountAllowed { get; set; }
    [DataMember]
    public string DiscountDateADayOfTheMo { get; set; }
    [DataMember]
    public int? MinimumDaysAllowedDisc { get; set; }
    [DataMember]
    public decimal? DiscountRate { get; set; }

    public static AP_TermsCode FromObject(PurchaseTerms tc)
    {
      return new AP_TermsCode
      {
        TermsCode = tc.Code,
        TermsCodeDesc = tc.Description,
        DaysBeforeDue = tc.DaysBeforeDue,
        DueDateADayOfTheMonth = tc.DueDateADayOfTheMonth,
        MinimumDaysAllowedInv = tc.MinimumDaysAllowedInv,
        DaysDiscountAllowed = tc.DaysDiscountAllowed,
        DiscountDateADayOfTheMo = tc.DiscountDateADayOfTheMonth,
        MinimumDaysAllowedDisc = tc.MinimumDaysAllowedDisc,
        DiscountRate = tc.DiscountRate,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7},{8}",
          this.csvEscape(this.TermsCode),
          this.csvEscape(this.TermsCodeDesc),
          this.csvEscape(this.DaysBeforeDue),
          this.csvEscape(this.DueDateADayOfTheMonth),
          this.csvEscape(this.MinimumDaysAllowedInv),
          this.csvEscape(this.DaysDiscountAllowed),
          this.csvEscape(this.DiscountDateADayOfTheMo),
          this.csvEscape(this.MinimumDaysAllowedDisc),
          this.csvEscape(this.DiscountRate)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("TermsCode", this.TermsCode, r.Field<string>("TermsCode"));
      yield return new Tuple<string, object, object>("TermsCodeDesc", this.TermsCodeDesc, r.Field<string>("TermsCodeDesc"));
      yield return new Tuple<string, object, object>("DaysBeforeDue", this.DaysBeforeDue, (int)r.Field<decimal>("DaysBeforeDue"));
      yield return new Tuple<string, object, object>("DueDateADayOfTheMonth", this.DueDateADayOfTheMonth, r.Field<string>("DueDateADayOfTheMonth"));
      yield return new Tuple<string, object, object>("MinimumDaysAllowedInv", this.MinimumDaysAllowedInv, (int)r.Field<decimal>("MinimumDaysAllowedInv"));
      yield return new Tuple<string, object, object>("DaysDiscountAllowed", this.DaysDiscountAllowed, (int)r.Field<decimal>("DaysDiscountAllowed"));
      yield return new Tuple<string, object, object>("DiscountDateADayOfTheMo", this.DiscountDateADayOfTheMo, r.Field<string>("DiscountDateADayOfTheMo"));
      yield return new Tuple<string, object, object>("MinimumDaysAllowedDisc", this.MinimumDaysAllowedDisc, (int)r.Field<decimal>("MinimumDaysAllowedDisc"));
      yield return new Tuple<string, object, object>("DiscountRate", this.DiscountRate, r.Field<decimal?>("DiscountRate"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<PurchaseTerms> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(PurchaseTerms current)
    {
      if (this.TermsCode != current.Code) { return false; }
      if (this.TermsCodeDesc != current.Description) { return false; }
      if (this.DaysBeforeDue != current.DaysBeforeDue) { return false; }
      if (this.DueDateADayOfTheMonth != current.DueDateADayOfTheMonth) { return false; }
      if (this.MinimumDaysAllowedInv != current.MinimumDaysAllowedInv) { return false; }
      if (this.DaysDiscountAllowed != current.DaysDiscountAllowed) { return false; }
      if (this.DiscountDateADayOfTheMo != current.DiscountDateADayOfTheMonth) { return false; }
      if (this.MinimumDaysAllowedDisc != current.MinimumDaysAllowedDisc) { return false; }
      if (this.DiscountRate != current.DiscountRate) { return false; }

      return true;
    }
    
  }
}
