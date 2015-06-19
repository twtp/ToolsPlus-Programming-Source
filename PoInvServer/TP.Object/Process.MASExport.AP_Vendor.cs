using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_AP_Vendor", Namespace = "TP.Object")]
  public class AP_Vendor : IMASExportObject
  {
    public static string MASColumns() { return "VendorNo, VendorName, AddressLine1, AddressLine2, City, State, ZipCode, CountryCode, TelephoneNo, TermsCode, Reference"; }
    public static string MASTableName() { return "AP_Vendor"; }
    public string PrimaryKey() { return this.VendorNo; }

    [DataMember]
    public string VendorNo { get; set; }
    [DataMember]
    public string VendorName { get; set; }
    [DataMember]
    public string AddressLine1 { get; set; }
    [DataMember]
    public string AddressLine2 { get; set; }
    [DataMember]
    public string City { get; set; }
    [DataMember]
    public string State { get; set; }
    [DataMember]
    public string ZipCode { get; set; }
    [DataMember]
    public string CountryCode { get; set; }
    [DataMember]
    public string TelephoneNo { get; set; }
    [DataMember]
    public string TermsCode { get; set; }
    [DataMember]
    public string Reference { get; set; }

    public static AP_Vendor FromObject(Vendor v)
    {
      return new AP_Vendor
      {
        VendorNo = v.VendorNo,
        VendorName = v.Name,
        AddressLine1 = v.AddressLine1,
        AddressLine2 = v.AddressLine2,
        City = v.City,
        State = v.State,
        ZipCode = v.PostalCode,
        CountryCode = v.CountryCode,
        TelephoneNo = v.TelephoneNo,
        TermsCode = v.DefaultPurchaseTermsCode,
        Reference = v.Reference,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}",
          this.csvEscape(this.VendorNo),
          this.csvEscape(this.VendorName),
          this.csvEscape(this.AddressLine1),
          this.csvEscape(this.AddressLine2),
          this.csvEscape(this.City),
          this.csvEscape(this.State),
          this.csvEscape(this.ZipCode),
          this.csvEscape(this.CountryCode),
          this.csvEscape(this.TelephoneNo),
          this.csvEscape(this.TermsCode),
          this.csvEscape(this.Reference)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("VendorNo", this.VendorNo, r.Field<string>("VendorNo"));
      yield return new Tuple<string, object, object>("VendorName", this.VendorName, r.Field<string>("VendorName"));
      yield return new Tuple<string, object, object>("AddressLine1", this.AddressLine1, r.Field<string>("AddressLine1"));
      yield return new Tuple<string, object, object>("AddressLine2", this.AddressLine2, r.Field<string>("AddressLine2"));
      yield return new Tuple<string, object, object>("City", this.City, r.Field<string>("City"));
      yield return new Tuple<string, object, object>("State", this.State, r.Field<string>("State"));
      yield return new Tuple<string, object, object>("ZipCode", this.ZipCode, r.Field<string>("ZipCode"));
      yield return new Tuple<string, object, object>("CountryCode", this.CountryCode, r.Field<string>("CountryCode"));
      yield return new Tuple<string, object, object>("TelephoneNo", this.TelephoneNo, r.Field<string>("TelephoneNo"));
      yield return new Tuple<string, object, object>("TermsCode", this.TermsCode, r.Field<string>("TermsCode"));
      yield return new Tuple<string, object, object>("Reference", this.Reference, r.Field<string>("Reference"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<Vendor> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(Vendor current)
    {
      if (this.VendorNo != current.VendorNo) { return false; }
      if (this.VendorName != current.Name) { return false; }
      if (this.AddressLine1 != current.AddressLine1) { return false; }
      if (this.AddressLine2 != current.AddressLine2) { return false; }
      if (this.City != current.City) { return false; }
      if (this.State != current.State) { return false; }
      if (this.ZipCode != current.PostalCode) { return false; }
      if (this.CountryCode != current.CountryCode) { return false; }
      if (this.TelephoneNo != current.TelephoneNo) { return false; }
      if (this.TermsCode != current.DefaultPurchaseTermsCode) { return false; }
      if (this.Reference != current.Reference) { return false; }

      return true;
    }

  }
}
