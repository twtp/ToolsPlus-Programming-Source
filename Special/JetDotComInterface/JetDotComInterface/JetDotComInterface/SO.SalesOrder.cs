using System;
using System.Collections.Generic;
using System.Text;

namespace TP.MAS200.SO {
  public class SalesOrder {

    #region Factory methods

    public static SalesOrder GetCurrentSalesOrder(string salesOrderNumber) {
      salesOrderNumber = TP.Base.Functions.CanonicalizeSalesOrderNumber(salesOrderNumber);
      return DBI.GetSalesOrder(salesOrderNumber);
    }

    public static Dictionary<string, SalesOrder> GetAllCurrentSalesOrdersAsDict() {
      return DBI.GetAllCurrentSalesOrders();
    }
    public static List<SalesOrder> GetAllCurrentSalesOrdersAsList() {
      Dictionary<string, SalesOrder> temp = DBI.GetAllCurrentSalesOrders();
      return new List<SalesOrder>(temp.Values);
    }

    public static SalesOrder GetSalesOrder(string salesOrderNumber) {
      salesOrderNumber = TP.Base.Functions.CanonicalizeSalesOrderNumber(salesOrderNumber);
      return DBI.GetSalesOrder(salesOrderNumber);
    }

    public static List<SalesOrderItemLine> GetBackorderLines() {
      List<SalesOrderItemLine> retval = new List<SalesOrderItemLine>();
      Dictionary<string, SalesOrder> temp = DBI.GetAllCurrentSalesOrders();
      foreach (string so in temp.Keys) {
        if (temp[so].OrderType == OrderType.Backorder) {
          foreach (TP.MAS200.SO.SalesOrderItemLine soil in temp[so].ItemLines) {
            if (soil.QuantityOrdered != soil.QuantityShipped) {
              retval.Add(soil);
            }
          }
        }
        else {
          foreach (TP.MAS200.SO.SalesOrderItemLine soil in temp[so].ItemLines) {
            if (soil.QuantityBackordered > 0) {
              retval.Add(soil);
            }
          }
        }
      }
      return retval;
    }

    public static List<string> GetCurrentSalesOrderNumbersWithItem(string item) {
      return DBI.GetCurrentSalesOrderNumbersWithItem(item);
    }

    #endregion

    private string salesOrderNumber;
    private DateTime orderDate;
    private OrderStatus orderStatus;
    private OrderType orderType;
    private string shipVia;
    private TP.Base.DataStructure.AddressBlock billingAddress;
    private TP.Base.DataStructure.AddressBlock shippingAddress;
    private OrderHoldReason holdReason;
    private string statusUDF;
    private string divisionNumber;
    private string customerNumber;
    private string salesPerson;
    private string commentsUDF;
    private string paymentType;
    private string otherPaymentTypeRefNo;
    private string emailAddress;
    private decimal depositAmt;
    private decimal totalOrderAmount;
    private decimal totalCaptureAmount;
    private string masHoldState;        // these are just used for debbie's
    private string masCancelReasonCode; // held order picked email (so far)
    private string creditCardTransactionID; // temporary, for pcc->sps migration
    private string createdByUserKey;
    private SY.User createdByUser = null;
    private string telephoneNo = null; // auto-loaded from AR.Customer on demand
    private string internalComments = null; // auto-loaded from SO.SalesOrderHeader (history?) on demand

    private DateTime createdDate;
    private DateTime updatedDate;

    private List<SalesOrderItemLine> lines = null;
    private List<SalesOrderAlertLine> alerts = null;
    private SalesOrderCommentLineList comments = null;

    /*internal SalesOrder(string salesOrderNumber, DateTime orderDate, SalesOrderStatus orderStatus, string shipVia, SalesOrderHoldReason holdReason) {
      this.salesOrderNumber = salesOrderNumber;
      this.orderDate = orderDate;
      this.orderStatus = orderStatus;
      this.shipVia = shipVia;
      this.holdReason = holdReason;
      this.lines = new List<SalesOrderItemLine>();
    }*/

    internal SalesOrder(System.Data.DataRow r, bool fromHistory) {
      // ACTIVE  FIELDS ARE SalesOrderNo, OrderDate, OrderType, ShipVia, UDF_CUST_STATUS
      // HISTORY FIELDS ARE SalesOrderNo, OrderDate, OrderStatus, ShipVia
      this.salesOrderNumber = r["SalesOrderNo"].ToString();
      this.orderDate = DateTime.Parse(r["OrderDate"].ToString());
      if (fromHistory) {
        this.orderType = OrderType.Standard;
      }
      else {
        switch (r["OrderType"].ToString()) {
          case "S":
            this.orderType = OrderType.Standard;
            break;
          case "Q":
            this.orderType = OrderType.Quote;
            break;
          case "B":
            this.orderType = OrderType.Backorder;
            break;
          default:
            TP.Base.Logger.WTF("missing OrderType '" + r["OrderType"].ToString() + "' in switch");
            this.orderType = OrderType.Standard;
            break;
        }
      }
      this.shipVia = r["ShipVia"].ToString();
      this.billingAddress = new TP.Base.DataStructure.AddressBlock(
          r["BillToName"].ToString(),
          r["BillToAddress1"].ToString(),
          r["BillToAddress2"].ToString(),
          r["BillToCity"].ToString(),
          r["BillToState"].ToString(),
          r["BillToZipCode"].ToString(),
          r["BillToCountryCode"].ToString(),
          /*r["UDF_BILL_ADDR_RESIDENTIAL"] == System.DBNull.Value ? (bool?)null : ((string)r["UDF_BILL_ADDR_RESIDENTIAL"] == "Y")*/
          r["UDF_BILL_ADDR_RESIDENTIAL"] == System.DBNull.Value ? TP.Base.AddressZoningType.Unknown : (string)r["UDF_BILL_ADDR_RESIDENTIAL"] == "Y" ? TP.Base.AddressZoningType.Residential : TP.Base.AddressZoningType.Commercial
        );
      this.shippingAddress = new TP.Base.DataStructure.AddressBlock(
          r["ShipToName"].ToString(),
          r["ShipToAddress1"].ToString(),
          r["ShipToAddress2"].ToString(),
          r["ShipToCity"].ToString(),
          r["ShipToState"].ToString(),
          r["ShipToZipCode"].ToString(),
          r["ShipToCountryCode"].ToString(),
          r["UDF_SHIP_ADDR_RESIDENTIAL"] == System.DBNull.Value ? TP.Base.AddressZoningType.Unknown : (string)r["UDF_SHIP_ADDR_RESIDENTIAL"] == "Y" ? TP.Base.AddressZoningType.Residential : TP.Base.AddressZoningType.Commercial
        );
      this.lines = new List<SalesOrderItemLine>();
      this.alerts = new List<SalesOrderAlertLine>();
      this.comments = null; // filled on-demand in getter
      this.statusUDF = r["UDF_CUST_STATUS"].ToString();
      this.divisionNumber = r["ARDivisionNo"].ToString();
      this.customerNumber = r["CustomerNo"].ToString();
      this.salesPerson = r["SalespersonNo"].ToString();
      this.commentsUDF = r["UDF_COMMENTS"].ToString();
      this.paymentType = r["PaymentType"] == System.DBNull.Value ? string.Empty : r["PaymentType"].ToString();
      this.otherPaymentTypeRefNo = r["OtherPaymentTypeRefNo"] == System.DBNull.Value ? string.Empty : r["OtherPaymentTypeRefNo"].ToString();
      this.emailAddress = r["EmailAddress"] == System.DBNull.Value ? string.Empty : (string)r["EmailAddress"];
      this.depositAmt = Convert.ToDecimal(r["DepositAmt"]);
      this.totalOrderAmount = Convert.ToDecimal(r["TaxableAmt"])
                            + Convert.ToDecimal(r["NonTaxableAmt"])
                            + Convert.ToDecimal(r["SalesTaxAmt"]);
      this.totalCaptureAmount = Convert.ToDecimal(r["CreditCardPreAuthorizationAmt"])
                              + Convert.ToDecimal(r["DepositAmt"]);
      // orderstatus
      // holdreason
      if (fromHistory) {
        // possible values for SO_SalesOrderHistoryHeader.OrderStatus are:
        //  (A) active
        //  (C) complete
        //  (X) cancelled
        //  (Q) quote
        //  (Z) cancelled quote
        switch (r["OrderStatus"].ToString()) {
          case "A":
          case "Q":
            throw new ApplicationException("sales order should have been pulled from active, not history");
          case "C":
            this.orderStatus = OrderStatus.Complete;
            break;
          case "X":
          case "Z":
            this.orderStatus = OrderStatus.Cancelled;
            break;
          default:
            throw new ApplicationException("undefined SalesOrderHistoryHeader OrderStatus field value");
        }
        this.holdReason = OrderHoldReason.NoHold;
      }
      else {
        // possible values for SO_SalesOrderHeader.UDF_CUST_STATUS are:
        //  OKAY
        //  VALIDATING
        //  PICKING - COMPLETE
        //  PICKING - B/O PART
        //  PICKING - B/O HOLD
        //  BACKORDER
        //  BAD CC INFO
        //  CC DECLINED
        //  FREIGHT QUOTE
        //  INTL VERIFY
        //  BILLTO/SHIPTO
        //  DROPSHIP
        switch (r["UDF_CUST_STATUS"].ToString()) {
          case "OKAY":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.ValidatingPayment;
            break;
          case "VALIDATING":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.ValidatingPayment;
            break;
          case "PICKING - COMPLETE":
            this.orderStatus = OrderStatus.PickingComplete;
            this.holdReason = OrderHoldReason.NoHold;
            break;
          case "PICKING - B/O PART":
            this.orderStatus = OrderStatus.PickingBackorderPartial;
            this.holdReason = OrderHoldReason.NoHold;
            break;
          case "PICKING - B/O HOLD":
            this.orderStatus = OrderStatus.PickingBackorderHold;
            this.holdReason = OrderHoldReason.NoHold;
            break;
          case "BACKORDER":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.Quantity;
            break;
          case "BAD CC INFO":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.BadPaymentInfo;
            break;
          case "CC DECLINED":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.PaymentDeclined;
            break;
          case "FREIGHT QUOTE":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.ResponseRequired;
            break;
          case "INTL VERIFY":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.ResponseRequired;
            break;
          case "BILLTO/SHIPTO":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.AddressDeclined;
            break;
          case "DROPSHIP":
            this.orderStatus = OrderStatus.Hold;
            this.holdReason = OrderHoldReason.Dropship;
            break;
          case "":
            // TODO
            break;
          default:
            throw new ApplicationException("undefined SalesOrderHeader UDF_CUST_STATUS field value");
        }
      }

      this.masHoldState = r["OrderStatus"].ToString();
      this.masCancelReasonCode = r["CancelReasonCode"].ToString();
      this.creditCardTransactionID = r["CreditCardTransactionID"].ToString();
      this.createdByUserKey = r["UserCreatedKey"].ToString();

      this.createdDate = this.convertDate(r["DateCreated"].ToString(), r["TimeCreated"].ToString());
      this.updatedDate = this.convertDate(r["DateUpdated"].ToString(), r["TimeUpdated"].ToString());
    }

    private DateTime convertDate(string dateField, string timeField) {
      // assumes that dateField is a stringified "mm/dd/yyyy 12:00:00 am" like mas usually has, and
      // timefield is the decimal hours and fractions of hours style
      try {
        if (dateField == string.Empty || timeField == string.Empty) {
          return DateTime.Now;
        }

        string date = dateField.Substring(0, dateField.IndexOf(' '));

        decimal timeNum = Convert.ToDecimal(timeField);
        int hours = (int)Math.Truncate(timeNum);
        decimal minutesPart = (timeNum - hours) * 60;
        int minutes = (int)Math.Truncate(minutesPart);
        int seconds = (int)((minutesPart - minutes) * 60);

        return DateTime.Parse(string.Format("{0} {1}:{2:00}:{3:00}", date, hours, minutes, seconds));
      }
      catch {
        return DateTime.Now;
      }
    }

    public string SalesOrderNumber { get { return this.salesOrderNumber; } }
    public DateTime OrderDate { get { return this.orderDate; } }
    public OrderType OrderType { get { return this.orderType; } }
    public OrderStatus OrderStatus { get { return this.orderStatus; } }
    public string ShipVia { get { return this.shipVia; } }
    public TP.Base.DataStructure.AddressBlock Billing { get { return this.billingAddress; } }
    public TP.Base.DataStructure.AddressBlock Shipping { get { return this.shippingAddress; } }
    public OrderHoldReason HoldReason { get { return this.holdReason; } }
    public List<SalesOrderItemLine> ItemLines { get { return this.lines; } }
    public List<SalesOrderAlertLine> AlertLines { get { return this.alerts; } }
    public SalesOrderCommentLineList CommentLines {
      get {
        if (this.comments == null) {
          this.comments = DBI.GetCommentLinesFor(this);
        }
        return this.comments;
      }
    }
    public string StatusUDF { get { return this.statusUDF; } }
    public string DivisionNo { get { return this.divisionNumber; } }
    public string CustomerNo { get { return this.customerNumber; } }
    public string CustomerNoFormatted { get { return string.Format("{0}-{1}", this.divisionNumber, this.customerNumber); } }
    public string SalesPerson { get { return this.salesPerson; } }
    public string CommentsUDF { get { return this.commentsUDF; } }
    public string PaymentType { get { return this.paymentType; } }
    public string OtherPaymentTypeRefNo { get { return this.otherPaymentTypeRefNo; } }
    public string EmailAddress { get { return this.emailAddress; } }
    public decimal DepositAmt { get { return this.depositAmt; } }
    public decimal TotalOrderAmount { get { return this.totalOrderAmount; } } // gives you actual total amount for a history, but only outstanding amount for a current!
    public decimal TotalCaptureAmount { get { return this.totalCaptureAmount; } }
    public string MasHoldState { get { return this.masHoldState; } }
    public string MasCancelReasonCode { get { return this.masCancelReasonCode; } }
    public string CreditCardTransactionID { get { return this.creditCardTransactionID; } }
    public bool LooksLikePCChargeStyleCreditCardTransactionID { get { return this.creditCardTransactionID != null && this.creditCardTransactionID.Length == 6 && this.creditCardTransactionID.StartsWith("6"); } }
    public SY.User CreatedByUser {
      get {
        if (this.createdByUser == null) {
          this.createdByUser = SY.User.GetUserByKey(this.createdByUserKey);
        }
        return this.createdByUser;
      }
    }
    public string TelephoneNo { //prioritizes shipping over billing, if this changes logic in callers needs changing
      get {
        if (this.telephoneNo == null) {
          this.telephoneNo = this.CommentLines.FindShippingTelephoneNo();
          if (this.telephoneNo == null) {
            TP.MAS200.AR.Customer cust = TP.MAS200.AR.Customer.GetCustomer(this.DivisionNo, this.CustomerNo);
            this.telephoneNo = cust == null ? string.Empty : cust.TelephoneNo;
          }
        }
        return this.telephoneNo;
      }
    }
    public string InternalComments {
      get {
        if (this.internalComments == null) {
          this.internalComments = DBI.GetInternalCommentsFor(this);
        }
        return this.internalComments;
      }
    }
    public DateTime CreatedDate { get { return this.createdDate; } }
    public DateTime UpdatedDate { get { return this.updatedDate; } }

    // duplicated in SalesOrderTask
    public bool IsStandardShippingMethod { get { return this.shipVia == "UPS INET" || this.shipVia == "GROUND INET"; } }

    public bool IsActiveForPicking {
      get {
        // standard or backorder, not on hold
        //return (this.OrderType == "S" || this.OrderType == "B") && (this.OrderStatus != "H");
        return this.orderStatus == OrderStatus.PickingComplete ||
               this.orderStatus == OrderStatus.PickingBackorderPartial ||
               this.orderStatus == OrderStatus.PickingBackorderHold;
      }
    }

    public bool IsCancelled {
      get {
        //return (this.OrderStatus == "X" || this.OrderStatus == "Z");
        return this.orderStatus == OrderStatus.Cancelled;
      }
    }

    public bool RequiresBackorderFix {
      get {
        return this.OrderType == OrderType.Backorder;
      }
    }

    public bool IsCustomerPickup {
      get {
        return this.shipVia == "PICK UP";
      }
    }

    public bool IsTruckFreight {
      get {
        return this.shipVia.StartsWith("TF ") || this.shipVia.StartsWith("TR");
      }
    }

    public bool? IsPaymentValid {
      get {
        switch (this.paymentType) {
          case "":
          case "TERMS":
          case "1TERM":
          case "CHECK":
            // don't know how to handle these yet
            return null;
          case "PAYPL":
            // dammit, paypal references this dll, can't do this here
            // maybe move this entire function out to logic? then
            // reference both fairly easily
            return null;
          default:
            // cash (deposit from p2), credit cards
            return this.totalCaptureAmount >= this.totalOrderAmount;
        }
      }
    }

    public bool IsDifferentBillToShipTo() {
      return this.Billing.DiffersFrom(this.Shipping);
    }

    public static List<SalesOrderItemLine> GetHistoryLinesFor(SalesOrder salesOrder) {
      return DBI.GetHistoryLinesFor(salesOrder);
    }

  }
}
