using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {
  public static class Functions {

    public static string CanonicalizeTelephoneNo(string telephoneNo) {
      string temp = System.Text.RegularExpressions.Regex.Replace(telephoneNo, "[^0-9]", string.Empty);
      if (temp.Length == 7) {
        return string.Format("203-{0}-{1}", temp.Substring(0, 3), temp.Substring(3));
      }
      if (temp.Length == 11 && temp.StartsWith("1")) {
        temp = temp.Substring(1);
      }
      if (temp.Length == 10) {
        return string.Format("{0}-{1}-{2}", temp.Substring(0, 3), temp.Substring(3, 3), temp.Substring(6));
      }
      return telephoneNo;
    }

    public static string CanonicalizeSalesOrderNumber(string salesOrderNumber) {
      if (System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^\d+$")) {
        // pad numeric sales orders to 7 characters
        return salesOrderNumber.PadLeft(7, '0');
      }
      else {
        // Y/Z sales order, just return
        return salesOrderNumber.ToUpper();
      }
    }

    public static string CanonicalizeGeneralLedgerBatchNumber(string batchNumber) {
      // batches are always numeric. mas probably supports otherwise, but we
      // can handle that if it comes up.
      if (System.Text.RegularExpressions.Regex.IsMatch(batchNumber, @"^\d{4,5}$")) {
        return batchNumber.PadLeft(5, '0');
      }
      else {
        // ok, so it does come up when ricky puts his initials in there...
        return batchNumber;
      }
      //throw new ArgumentException("Invalid batch number!");
    }

    public static string CanonicalizeInvoiceNumber(string invoiceNumber) {
      // TODO: what to do about headerseqno?
      // right now we're only looking at current shipping batches, which don't
      // use that field, but if we eventually add historical, then that will
      // need to be handled.
      if (System.Text.RegularExpressions.Regex.IsMatch(invoiceNumber, @"^\d+$")) {
        // pad numeric invoices to 7 characters
        return invoiceNumber.PadLeft(7, '0');
      }
      else {
        // non-numeric is still ok, possibly an A/B invoice. just return.
        return invoiceNumber;
      }
    }

    public static string CanonicalizeP2TransactionID(string transactionNumber) {
      if (System.Text.RegularExpressions.Regex.IsMatch(transactionNumber, @"^\d+$")) {
        return transactionNumber.PadLeft(7, '0');
      }
      throw new ArgumentException("Invalid transaction number!");
    }

    public static string CanonicalizeCashDrawerNumber(string cashDrawerID) {
      // null is ok, means all drawers later on
      if (cashDrawerID == null) {
        return null;
      }
      else if (System.Text.RegularExpressions.Regex.IsMatch(cashDrawerID, @"^\d{4,7}$")) {
        return cashDrawerID.PadLeft(7, '0');
      }
      else {
        throw new ArgumentException("Invalid cash drawer number!");
      }
    }

    public static string CanonicalizePurchaseOrderNumber(string purchaseOrderNumber) {
      int foo;
      if (Int32.TryParse(purchaseOrderNumber, out foo)) {
        // numeric, needs zero-padding in front
        return string.Format("{0:0000000}", foo);
      }
      else {
        // alpha and numeric, just capitalize
        return purchaseOrderNumber.ToUpper();
      }
    }

    public static string CanonicalizeReceiptOfGoodsNumber(string receiptNumber) {
      receiptNumber = receiptNumber.ToUpper();
      if (receiptNumber.StartsWith("G")) {
        receiptNumber = receiptNumber.Remove(0, 1);
      }
      return "G" + receiptNumber.PadLeft(6, '0');
    }

    // TODO: return of goods?

    public static bool LooksLikePhoneOrder(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      return System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^\d{7}$");
    }
    public static bool LooksLikeToolsPlusOrder(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      return System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^Y\d{6}$");
    }
    public static bool LooksLikeToolPartsStoreOrder(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      return System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^P\d{6}$");
    }
    public static bool LooksLikeToolsPlusEBayOutletOrder(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      return salesOrderNumber.Length == 7 && System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^E\d{1,6}[A-Z]{0,5}$");
    }
    public static bool LooksLikeDropshipSplitOrder(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      return System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^Z\d{5,6}$");
    }
    public static string ConvertDropshipSplitOrderToOriginal(string salesOrderNumber) {
      salesOrderNumber = CanonicalizeSalesOrderNumber(salesOrderNumber);
      if (LooksLikeDropshipSplitOrder(salesOrderNumber) == false) {
        return salesOrderNumber;
      }
      if (System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^Z\d{5}$")) {
        return "00" + salesOrderNumber.Substring(1);
      }
      else {
        return "Y" + salesOrderNumber.Substring(1);
      }
    }

    public static SellingChannel? GetSellingChannelFor(string salesOrderNo) {
      if (LooksLikeDropshipSplitOrder(salesOrderNo)) {
        salesOrderNo = ConvertDropshipSplitOrderToOriginal(salesOrderNo);
      }
      if (LooksLikePhoneOrder(salesOrderNo)) {
        return SellingChannel.Telephone;
      }
      if (LooksLikeToolsPlusOrder(salesOrderNo)) {
        return SellingChannel.ToolsPlus;
      }
      if (LooksLikeToolPartsStoreOrder(salesOrderNo)) {
        return SellingChannel.ToolPartsStore;
      }
      if (LooksLikeToolsPlusEBayOutletOrder(salesOrderNo)) {
        return SellingChannel.ToolsPlusOutlet;
      }
      return null;
    }
  }
}
