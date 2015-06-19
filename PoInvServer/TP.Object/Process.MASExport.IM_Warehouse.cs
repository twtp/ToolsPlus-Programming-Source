using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_IM_Warehouse", Namespace = "TP.Object")]
  public class IM_Warehouse : IMASExportObject
  {
    public static string MASColumns() { return "WarehouseCode, WarehouseDesc, WarehouseName, WarehouseAddress1, WarehouseAddress2, WarehouseAddress3, WarehouseCity, WarehouseState, WarehouseZipCode, WarehouseCountryCode"; }
    public static string MASTableName() { return "IM_Warehouse"; }
    public string PrimaryKey() { return this.WarehouseCode; }

    [DataMember]
    public string WarehouseCode { get; set; }
    [DataMember]
    public string WarehouseDesc { get; set; }
    [DataMember]
    public string WarehouseName { get; set; }
    [DataMember]
    public string WarehouseAddress1 { get; set; }
    [DataMember]
    public string WarehouseAddress2 { get; set; }
    [DataMember]
    public string WarehouseAddress3 { get; set; }
    [DataMember]
    public string WarehouseCity { get; set; }
    [DataMember]
    public string WarehouseState { get; set; }
    [DataMember]
    public string WarehouseZipCode { get; set; }
    [DataMember]
    public string WarehouseCountryCode { get; set; }

    public static IM_Warehouse FromObject(MasWarehouse w)
    {
      return new IM_Warehouse
      {
        WarehouseCode = w.WarehouseCode,
        WarehouseDesc = w.Description,
        WarehouseName = w.Name,
        WarehouseAddress1 = w.Address1,
        WarehouseAddress2 = w.Address2,
        WarehouseAddress3 = w.Address3,
        WarehouseCity = w.City,
        WarehouseState = w.State,
        WarehouseZipCode = w.PostalCode,
        WarehouseCountryCode = w.CountryCode,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}",
          this.csvEscape(this.WarehouseCode),
          this.csvEscape(this.WarehouseDesc),
          this.csvEscape(this.WarehouseName),
          this.csvEscape(this.WarehouseAddress1),
          this.csvEscape(this.WarehouseAddress2),
          this.csvEscape(this.WarehouseAddress3),
          this.csvEscape(this.WarehouseCity),
          this.csvEscape(this.WarehouseState),
          this.csvEscape(this.WarehouseZipCode),
          this.csvEscape(this.WarehouseCountryCode)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("WarehouseCode", this.WarehouseCode, r.Field<string>("WarehouseCode"));
      yield return new Tuple<string, object, object>("WarehouseDesc", this.WarehouseDesc, r.Field<string>("WarehouseDesc"));
      yield return new Tuple<string, object, object>("WarehouseName", this.WarehouseName, r.Field<string>("WarehouseName"));
      yield return new Tuple<string, object, object>("WarehouseAddress1", this.WarehouseAddress1, r.Field<string>("WarehouseAddress1"));
      yield return new Tuple<string, object, object>("WarehouseAddress2", this.WarehouseAddress2, r.Field<string>("WarehouseAddress2"));
      yield return new Tuple<string, object, object>("WarehouseAddress3", this.WarehouseAddress3, r.Field<string>("WarehouseAddress3"));
      yield return new Tuple<string, object, object>("WarehouseCity", this.WarehouseCity, r.Field<string>("WarehouseCity"));
      yield return new Tuple<string, object, object>("WarehouseState", this.WarehouseState, r.Field<string>("WarehouseState"));
      yield return new Tuple<string, object, object>("WarehouseZipCode", this.WarehouseZipCode, r.Field<string>("WarehouseZipCode"));
      yield return new Tuple<string, object, object>("WarehouseCountryCode", this.WarehouseCountryCode, r.Field<string>("WarehouseCountryCode"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<MasWarehouse> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(MasWarehouse current)
    {
      if (this.WarehouseCode != current.WarehouseCode) { return false; }
      if (this.WarehouseDesc != current.Description) { return false; }
      if (this.WarehouseName != current.Name) { return false; }
      if (this.WarehouseAddress1 != current.Address1) { return false; }
      if (this.WarehouseAddress2 != current.Address2) { return false; }
      if (this.WarehouseAddress3 != current.Address3) { return false; }
      if (this.WarehouseCity != current.City) { return false; }
      if (this.WarehouseState != current.State) { return false; }
      if (this.WarehouseZipCode != current.PostalCode) { return false; }
      if (this.WarehouseCountryCode != current.CountryCode) { return false; }

      return true;
    }

  }
}
