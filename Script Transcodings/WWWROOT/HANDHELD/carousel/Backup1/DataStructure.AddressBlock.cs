using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class AddressBlock {

    public AddressBlock(string name, string line1, string line2, string city, string state, string zipCode, string country, string telephoneNo)
      : this(null, name, line1, line2, null, city, state, zipCode, country, AddressZoningType.Unknown, telephoneNo) {
      // nothing
    }
    public AddressBlock(string name, string line1, string line2, string city, string state, string zipCode, string country, AddressZoningType zoningType)
      : this(null, name, line1, line2, null, city, state, zipCode, country, zoningType, null) {
      // nothing
    }
    public AddressBlock(string code, string name, string line1, string line2, string city, string state, string zipCode, string country, int zoningType, string telephoneNo)
      : this(code, name, line1, line2, null, city, state, zipCode, country, (TP.Base.AddressZoningType)zoningType, telephoneNo) {
      // nothing
    }
    public AddressBlock(string code, string name, string line1, string line2, string city, string state, string zipCode, string country, AddressZoningType zoningType, string telephoneNo)
      : this(code, name, line1, line2, null, city, state, zipCode, country, zoningType, telephoneNo) {
      // nothing
    }
    public AddressBlock(string code, string name, string line1, string line2, string line3, string city, string state, string zipCode, string country, AddressZoningType zoningType, string telephoneNo) {
      this.code = code;
      this.name = name;
      this.line1 = line1;
      this.line2 = line2;
      this.line3 = line3;
      this.city = city;
      this.state = state;
      this.zipCode = zipCode;
      this.countryCode = country;
      // we really should clean up mas CC maintenance instead of hacking this here
      switch (this.countryCode) {
        case "USA":
        case "":
          this.countryCode = "US";
          break;
        case "SP":
          this.countryCode = "ES";
          break;
      }
      this.zoningType = zoningType;
      this.telephoneNo = telephoneNo;
    }

    private string code;
    private string name;
    private string line1;
    private string line2;
    private string line3;
    private string city;
    private string state;
    private string zipCode;
    private string countryCode;
    private AddressZoningType zoningType;
    private string telephoneNo;

    public string Code { get { return this.code; } }
    public string Name { get { return this.name; } }
    public string Line1 { get { return this.line1; } }
    public string Line2 { get { return this.line2; } }
    public string Line3 { get { return this.line3; } }
    public string City { get { return this.city; } }
    public string State { get { return this.state; } }
    public string ZipCode { get { return this.zipCode; } }
    public string CountryCode { get { return this.countryCode; } }
    public AddressZoningType ZoningType { get { return this.zoningType; } }
    public string TelephoneNo { get { return this.telephoneNo; } }

    public bool? IsResidential { get { return this.zoningType == AddressZoningType.Unknown ? (bool?)null : this.zoningType == AddressZoningType.Residential; } }
    public bool IsResidentialDefaultYes { get { return this.zoningType == AddressZoningType.Unknown || this.zoningType == AddressZoningType.Residential; } }

    public bool DiffersFrom(AddressBlock other) {
      if (this.name != other.name) { return true; }
      if (this.line1 != other.line1) { return true; }
      if (this.line2 != other.line2) { return true; }
      if (this.line3 != other.line3) { return true; }
      if (this.city != other.city) { return true; }
      if (this.state != other.state) { return true; }
      if (this.zipCode != other.zipCode) { return true; }
      if (this.countryCode != other.countryCode) { return true; }
      if (this.zoningType != other.zoningType) { return true; }
      if (this.telephoneNo != other.telephoneNo) { return true; }
      return false;
    }

    public string TextBlock() {
      return this.TextBlock(false);
    }
    public string TextBlock(bool includeTelephoneNo) {
      List<string> linesToUse = new List<string>();
      linesToUse.Add(this.Name);
      linesToUse.Add(this.Line1);
      if (string.IsNullOrEmpty(this.Line2) == false) {
        linesToUse.Add(this.Line2);
      }
      if (string.IsNullOrEmpty(this.Line3) == false) {
        linesToUse.Add(this.Line3);
      }
      linesToUse.Add(this.City + ", " + this.State + " " + this.ZipCode);
      if (this.CountryCode != "US") {
        linesToUse.Add(this.CountryCode);
      }
      if (includeTelephoneNo && string.IsNullOrEmpty(this.telephoneNo) == false) {
        linesToUse.Add(this.TelephoneNo);
      }

      return string.Join(Environment.NewLine, linesToUse.ToArray());
    }

    private static System.Text.RegularExpressions.Regex poBoxRE = new System.Text.RegularExpressions.Regex(@"P\.?\s*?O\.?\s*?BO?X?\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    public bool IsPostOfficeBox {
      get {
        return poBoxRE.IsMatch(this.line1);
      }
    }

    public bool IsDomestic { get { return this.countryCode == "US"; } }
    public bool IsInternational { get { return this.countryCode != "US"; } }
  }
}
