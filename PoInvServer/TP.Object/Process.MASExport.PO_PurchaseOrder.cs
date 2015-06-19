using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.Process.MASExport
{
  public class PO_PurchaseOrder : IMASExportObject
  {
    public string PurchaseOrderNo { get; set; }
    public DateTime OrderDate { get; set; }
    public DateTime ExpireDate { get; set; }
    public string VendorNo { get; set; }
    public string ShipToCode { get; set; }
    public string ShipToName { get; set; }
    public string ShipToAddress1 { get; set; }
    public string ShipToAddress2 { get; set; }
    public string ShipToCity { get; set; }
    public string ShipToState { get; set; }
    public string ShipToZipCode { get; set; }
    public string TermsCode { get; set; }
    public IEnumerable<PO_PurchaseOrderDetail> DetailLines { get; set; }
    public TP.Object.IExportableToMAS OriginalObject { get; set; }

    // ASSERT: detail lines and detail line items are all filled in
    public static PO_PurchaseOrder FromObject(TP.Object.PurchaseOrder po)
    {
      return new PO_PurchaseOrder
      {
        PurchaseOrderNo = po.PurchaseOrderNo,
        OrderDate = DateTime.Today,
        ExpireDate = po.DueDate,
        VendorNo = po.VendorNo,
        ShipToCode = po.ShipToCode,
        ShipToName = po.ShipToCode == "1234" ? po.ShipToName : string.Empty,
        ShipToAddress1 = po.ShipToCode == "1234" ? po.ShipToAddress1 : string.Empty,
        ShipToAddress2 = po.ShipToCode == "1234" ? po.ShipToAddress2 : string.Empty,
        ShipToCity = po.ShipToCode == "1234" ? po.ShipToCity : string.Empty,
        ShipToState = po.ShipToCode == "1234" ? po.ShipToState : string.Empty,
        ShipToZipCode = po.ShipToCode == "1234" ? po.ShipToPostalCode : string.Empty,
        TermsCode = po.PaymentTerms,
        DetailLines = PO_PurchaseOrderDetail.FromObject(po),
        OriginalObject = po,
      };
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("H,{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}",
          this.csvEscape(this.PurchaseOrderNo),
          this.csvEscape(this.OrderDate.ToString("yyyyMMdd")),
          this.csvEscape(this.ExpireDate.ToString("yyyyMMdd")),
          this.csvEscape(this.VendorNo),
          this.csvEscape(this.ShipToCode),
          this.csvEscape(this.ShipToName),
          this.csvEscape(this.ShipToAddress1),
          this.csvEscape(this.ShipToAddress2),
          this.csvEscape(this.ShipToCity),
          this.csvEscape(this.ShipToState),
          this.csvEscape(this.ShipToZipCode),
          this.csvEscape(this.TermsCode),
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty
        );
      foreach (var l in this.DetailLines)
      {
        sb.AppendLine();
        l.ToCSVLine(sb);
      }
    }
  }


  public class PO_PurchaseOrderDetail
  {
    public string ItemCode { get; set; }
    public int QuantityOrdered { get; set; }
    public decimal UnitCost { get; set; }
    public string CommentText { get; set; }

    public static IEnumerable<PO_PurchaseOrderDetail> FromObject(TP.Object.PurchaseOrder po)
    {
      var retval = new List<PO_PurchaseOrderDetail>(po.PurchaseOrderItems.Count());
      foreach (var l in po.PurchaseOrderItems)
      {
        var cost = l.Item.ItemCosts.SingleOrDefault(i => i.CostType == "Std Cost");
        retval.Add(new PO_PurchaseOrderDetail
        {
          ItemCode = l.Item.ItemNumber,
          QuantityOrdered = l.Quantity,
          UnitCost = cost == null ? 0M : cost.CostValue,
          CommentText = l.Item.PurchaseOrderComment,
        });
      }
      return retval;
    }

    public void ToCSVLine(StringBuilder sb)
    {
      sb.AppendFormat("L,{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}",
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          string.Empty,
          this.csvEscape(this.ItemCode),
          this.csvEscape(this.QuantityOrdered),
          this.csvEscape(this.UnitCost),
          this.csvEscape(this.CommentText)
        );
    }
  }
}
