using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;
using System.Data;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJob_IM_SalesKit", Namespace = "TP.Object")]
  public class IM_SalesKit : IMASExportObject
  {
    public static string MASColumns() { return "SalesKitNo, ComponentItemCode, QuantityPerAssembly"; }
    public static string MASTableName() { return "IM_SalesKitDetail"; }
    public string PrimaryKey() { return string.Format("{0}/{1}", this.SalesKitNo, this.ComponentItemCode); }

    [DataMember]
    public string SalesKitNo { get; set; }
    [DataMember]
    public string ComponentItemCode { get; set; }
    [DataMember]
    public int QuantityPerAssembly { get; set; }

    public static IM_SalesKit FromObject(ItemKitContent kc)
    {
      return new IM_SalesKit
      {
        SalesKitNo = kc.ItemNumber,
        ComponentItemCode = kc.SubItemNumber,
        QuantityPerAssembly = kc.Quantity,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("{0},{1},{2}",
          this.csvEscape(this.SalesKitNo),
          this.csvEscape(this.ComponentItemCode),
          this.csvEscape(this.QuantityPerAssembly)
        );
    }

    public IEnumerable<Tuple<string, object, object>> Validate(DataRow r)
    {
      yield return new Tuple<string, object, object>("SalesKitNo", this.SalesKitNo, r.Field<string>("SalesKitNo"));
      yield return new Tuple<string, object, object>("ComponentItemCode", this.ComponentItemCode, r.Field<string>("ComponentItemCode"));
      yield return new Tuple<string, object, object>("QuantityPerAssembly", this.QuantityPerAssembly, (int)r.Field<decimal>("QuantityPerAssembly"));
      yield break;
    }

    public bool FlagExportedIfUnchanged(Func<ItemKitContent> getter)
    {
      var obj = getter();
      return obj.FlagExported(() => this.Recompare(obj));
    }

    public bool Recompare(ItemKitContent current)
    {
      if (this.SalesKitNo != current.ItemNumber) { return false; }
      if (this.ComponentItemCode != current.SubItemNumber) { return false; }
      if (this.QuantityPerAssembly != current.Quantity) { return false; }

      return true;
    }

  }
}
