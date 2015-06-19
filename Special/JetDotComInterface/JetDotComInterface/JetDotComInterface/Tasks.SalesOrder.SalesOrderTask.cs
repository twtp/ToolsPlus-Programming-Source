using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.Tasks.SalesOrder {
  public class SalesOrderTask : IWarehousingTask, IEquatable<IWarehousingTask> {

    private int databaseID;
    public int DatabaseID {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.databaseID; }
    }

    private string salesOrderNumber;
    public string SalesOrderNumber {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.salesOrderNumber; }
    }

    private int warehouseID;
    public int WarehouseID {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.warehouseID; }
    }

    private DateTime startingTime;
    public DateTime StartingTime {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.startingTime; }
    }

    private SalesOrderStatus currentStatus;
    public SalesOrderStatus CurrentStatus {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.currentStatus; }
    }
    public string CurrentStatusPrintable { 
      get { return this.getEnumDescription(this.currentStatus); }
    }

    private SalesOrderHoldReason currentHoldReason;
    public SalesOrderHoldReason CurrentHoldReason {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.currentHoldReason; }
    }
    public string CurrentHoldReasonPrintable {
      get { return this.getEnumDescription(this.currentHoldReason); }
    }

    private string shipVia;
    public string ShipVia {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.shipVia; }
    }

    private string shipState;
    public string ShipState {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.shipState; }
    }

    private string shipCountry;
    public string ShipCountry {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.shipCountry; }
    }

    private string createdByUserLogon;
    public string CreatedByUserLogon {
      [System.Diagnostics.DebuggerStepThrough()]
      get { return this.createdByUserLogon; }
    }

    private List<SalesOrderTaskLine> lines;
    public List<SalesOrderTaskLine> Lines { get { return this.lines; } }
    public SalesOrderTaskLine Line(int componentID) {
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.ComponentID == componentID) {
          return sotl;
        }
      }
      return null;
    }

    #region Canonicalization

    public static string Canonicalize(string salesOrderNumber) {
      if (System.Text.RegularExpressions.Regex.IsMatch(salesOrderNumber, @"^\d+$")) {
        // pad numeric sales orders to 7 characters
        return salesOrderNumber.PadLeft(7, '0');
      }
      else {
        // Y/Z sales order, just return
        return salesOrderNumber;
      }
    }

    #endregion

    #region Factory Methods

    public static SalesOrderTask CreateNew(string salesOrderNumber, int warehouseID, int currentStatus, int currentHoldReason, string shipVia, string shipState, string shipCountry, string createdByUserLogon, Dictionary<int, int> componentsOrdered, Dictionary<int, int> componentsBackordered, Dictionary<int, int> componentsShipped) {
      // both components hashes should have identical keys
      salesOrderNumber = Canonicalize(salesOrderNumber);
      System.Data.SqlClient.SqlTransaction txn = TP.Database.SqlServer.BeginTransaction();
      try {

        int headerID = DBI.CreateHeader(txn, salesOrderNumber, warehouseID, currentStatus, currentHoldReason, shipVia, shipState, shipCountry, createdByUserLogon);
        SalesOrderTask retval = new SalesOrderTask(headerID, salesOrderNumber, warehouseID, currentStatus, currentHoldReason, shipVia, shipState, shipCountry, createdByUserLogon, DateTime.Now);

        foreach (KeyValuePair<int, int> kvp in componentsOrdered) {
          int cid = kvp.Key;
          int qtyO = kvp.Value;
          int qtyB = componentsBackordered[cid];
          int qtyS = componentsShipped[cid];

          int lineID = DBI.CreateLine(txn, headerID, cid, qtyO, qtyB, qtyS);
          retval.Lines.Add(new SalesOrderTaskLine(retval, lineID, cid, qtyO, qtyB, qtyS, qtyS, qtyS));
        }

        TP.Database.SqlServer.CommitTransaction(txn);
        return retval;
      }
      catch (Exception) {
        // well fuck, what do we do now? where is this being called from? can they handle it?
        TP.Database.SqlServer.RollbackTransaction(txn);
        throw;
      }
    }

    public static SalesOrderTask LoadExisting(int headerID) {
      return DBI.LoadExisting(headerID);
    }
    public static SalesOrderTask LoadExisting(string identifier) {
      // identifier is the SO sales order number
      return DBI.LoadExisting(Canonicalize(identifier));
    }

    public static List<SalesOrderTaskLine> LoadActiveLines(int warehouseID, int? componentID) {
      List<SalesOrderTask> trans = GetActiveOrdersAsList(warehouseID);
      List<SalesOrderTaskLine> retval = new List<SalesOrderTaskLine>();
      foreach (SalesOrderTask sot in trans) {
        foreach (SalesOrderTaskLine sotl in sot.Lines) {
          if (componentID == null || sotl.ComponentID == componentID) {
            retval.Add(sotl);
          }
        }
      }
      return retval;
    }

    public static List<SalesOrderTask> GetActiveOrdersAsList(int warehouseID) {
      Dictionary<string, SalesOrderTask> temp = DBI.GetActiveOrders(warehouseID);
      return new List<SalesOrderTask>(temp.Values);
    }
    public static Dictionary<string, SalesOrderTask> GetActiveOrdersAsDict(int warehouseID) {
      return DBI.GetActiveOrders(warehouseID);
    }

    public static List<SalesOrderTaskLine> GetActiveLinesWithComponent(int warehouseID, int componentID) {
      return DBI.GetActiveLinesWithComponent(warehouseID, componentID);
    }

    #endregion

    public bool HasAnyActivityStarted() {
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityPicked > 0) {
          return true;
        }
      }
      return false;
    }

    public bool HasPickingStarted() {
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityPicked > 0 && sotl.QuantityPicked - sotl.QuantityShipped > 0) {
          return true;
        }
      }
      return false;
    }

    public bool HasPackageAwaitingPickup() {
      // this means that the package is currently awaiting pickup, previous partial
      // shipments do not count if they're been picked up. for a really-all-done
      // check, see HasPackingCompleted().
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityStaged > 0 && sotl.QuantityStaged - sotl.QuantityShipped > 0) {
          return true;
        }
      }
      return false;
    }

    public bool HasPickingCompleted() {
      // this needs to mean completely completed, not just backorder part completed
      // packing uses this to identify where a b/o hold is going (bin or out)
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityOrdered != sotl.QuantityPicked) {
          return false;
        }
      }
      return true;
    }

    public bool HasPackingCompleted() {
      // completely completed, no more backorders. for a day's packing completed,
      // potentially with remaining backorders, see HasPackagageAwaitingPickup().
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityPicked != sotl.QuantityStaged) {
          return false;
        }
      }
      return true;
    }

    public bool IsReadyForShipping() {
      if (this.currentStatus != SalesOrderStatus.PickingComplete && this.currentStatus != SalesOrderStatus.PickingBackorderPartial) {
        // picking for b/o or non-picking hold
        return false;
      }
      else {
        // picking to ship
        bool hasPriorShipment = this.HasPriorShipment();
        foreach (SalesOrderTaskLine sotl in this.lines) {
          int qtyReqd = sotl.QuantityOrdered - sotl.QuantityShipped;
          if (this.currentStatus == SalesOrderStatus.PickingBackorderPartial && !hasPriorShipment) {
            qtyReqd -= sotl.QuantityBackordered;
          }
          int qtyPick = sotl.QuantityPicked - sotl.QuantityStaged;

          if (qtyPick < qtyReqd) {
            return false;
          }
        }
        return true;
      }
    }

    public bool HasPriorShipment() {
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.QuantityShipped > 0) {
          return true;
        }
      }
      return false;
    }

    public SalesOrderTaskLine AddNewComponent(int componentID, int orderedQty, int backorderQty, int shippedQuantity) {
      foreach (SalesOrderTaskLine sotl in this.lines) {
        if (sotl.ComponentID == componentID) {
          // error, or return with it?
          throw new ArgumentException("ComponentID is already in SalesOrderTask");
        }
      }

      System.Data.SqlClient.SqlTransaction txn = TP.Database.SqlServer.BeginTransaction();
      try {
        int lineID = DBI.CreateLine(txn, this.databaseID, componentID, orderedQty, backorderQty, shippedQuantity);
        SalesOrderTaskLine sotl = new SalesOrderTaskLine(this, lineID, componentID, orderedQty, backorderQty, shippedQuantity, shippedQuantity, shippedQuantity);
        this.lines.Add(sotl);
        TP.Database.SqlServer.CommitTransaction(txn);
        return sotl;
      }
      catch {
        TP.Database.SqlServer.RollbackTransaction(txn);
        throw;
      }
    }

    /*public bool ChangeStatus(SalesOrderStatus newStatus, SalesOrderHoldReason newHoldReason, string newShipVia, string newShipState, string newShipCountry) {
      return this.ChangeStatus((int)newStatus, (int)newHoldReason, newShipVia, newShipState, newShipCountry);
    }
    public bool ChangeStatus(int newStatus, int newHoldReason, string newShipVia, string newShipState, string newShipCountry) {
      if (DBI.ChangeStatus(this.databaseID, newStatus, newHoldReason, newShipVia, newShipState, newShipCountry)) {
        this.currentStatus = (SalesOrderStatus)newStatus;
        this.currentHoldReason = (SalesOrderHoldReason)newHoldReason;
        this.shipVia = newShipVia;
        this.shipState = newShipState;
        this.shipCountry = newShipCountry;
        return true;
      }
      else {
        return false;
      }
    }*/
    public bool ChangeStatus(Dictionary<SalesOrderHeaderField, object> newStates) {
      if (DBI.ChangeStatus(this.databaseID, newStates)) {
        foreach (SalesOrderHeaderField hf in newStates.Keys) {
          switch (hf) {
            case SalesOrderHeaderField.HoldReason:
              this.currentHoldReason = (SalesOrderHoldReason)newStates[hf];
              break;
            case SalesOrderHeaderField.ShipCountry:
              this.shipCountry = (string)newStates[hf];
              break;
            case SalesOrderHeaderField.ShipState:
              this.shipState = (string)newStates[hf];
              break;
            case SalesOrderHeaderField.ShipVia:
              this.shipVia = (string)newStates[hf];
              break;
            case SalesOrderHeaderField.Status:
              this.currentStatus = (SalesOrderStatus)newStates[hf];
              break;
            case SalesOrderHeaderField.CreatedByUserLogon:
              this.createdByUserLogon = (string)newStates[hf];
              break;
            default:
              TP.Base.Logger.WTF("missing SalesOrderHeaderField in ChangeStatus() switch");
              break;
          }
        }
        return true;
      }
      else {
        return false;
      }
    }

    public bool IsPickupSalesOrder { get { return this.shipVia == "PICK UP"; } }

    // duplicated in MAS200.SalesOrder
    public bool IsStandardShippingMethod { get { return this.shipVia == "UPS INET" || this.shipVia == "GROUND INET"; } }

    public bool IsActiveForPicking {
      get {
        return this.currentStatus == SalesOrderStatus.PickingComplete ||
               this.currentStatus == SalesOrderStatus.PickingBackorderPartial ||
               this.currentStatus == SalesOrderStatus.PickingBackorderHold;
      }
    }

    #region Private CTOR

    // should be private, but is internal so DBI can access it
    internal SalesOrderTask(int databaseID, string salesOrderNumber, int warehouseID, int currentStatus, int currentHoldReason, string shipVia, string shipState, string shipCountry, string createdByUserLogon, DateTime startingTime) {
      this.databaseID = databaseID;
      this.salesOrderNumber = salesOrderNumber;
      this.warehouseID = warehouseID;
      this.currentStatus = (SalesOrderStatus)currentStatus;
      this.currentHoldReason = (SalesOrderHoldReason)currentHoldReason;
      this.shipVia = shipVia;
      this.shipState = shipState;
      this.shipCountry = shipCountry;
      this.createdByUserLogon = createdByUserLogon;
      this.startingTime = startingTime;
      this.lines = new List<SalesOrderTaskLine>();
    }

    #endregion

    #region IWarehousingTask Members

    public bool IsFakeTask { get { return false; } }

    public TransactionType TransactionType { get { return TransactionType.SalesOrder; } }
    public string TransactionTypePrintable { get { return "Sales Order"; } }

    public string TransactionReference { get { return this.salesOrderNumber; } }

    public List<IWarehousingTaskLine> TransactionLines {
      get {
        return this.lines.ConvertAll(new Converter<SalesOrderTaskLine, IWarehousingTaskLine>(delegate(SalesOrderTaskLine sotl) { return (IWarehousingTaskLine)sotl; }));
      }
    }

    public bool PickPriorityFlagUnusual { get { return !this.IsStandardShippingMethod; } }

    public bool PackingRequiresAllTotes {
      get {
        if (this.currentStatus != SalesOrderStatus.PickingBackorderHold) {
          // shipments require the entire set of totes to pack
          return true;
        }

        // pickingbackorderhold only requires the initial tote, to send to a b/o bin
        // however, this is only if the order is being b/o'ed, if it is ready to ship, then
        // all are required again
        if (!this.HasPickingCompleted()) {
          return false;
        }
        return true;
      }
    }

    public bool PackingPartialOK {
      get {
        switch (this.currentStatus) {
          case SalesOrderStatus.PickingBackorderHold:
            return true;
          case SalesOrderStatus.PickingBackorderPartial:
            return true;
          default:
            return false;
        }
      }
    }

    public Locations.LocationType PackingDestinationLocationType {
      get {
        switch (this.currentStatus) {
          case SalesOrderStatus.PickingComplete:
            return Locations.LocationType.StagedShippingLocation;
          case SalesOrderStatus.PickingBackorderPartial:
            return Locations.LocationType.StagedShippingLocation;
          case SalesOrderStatus.PickingBackorderHold:
            // if order picking is complete, goto shipping
            // otherwise, b/o holding
            if (this.HasPickingCompleted()) {
              return Locations.LocationType.StagedShippingLocation;
            }
            else {
              return Locations.LocationType.BackorderLocation;
            }
          default:
            throw new ApplicationException("shouldn't be packing an order with status " + this.currentStatus.ToString());
        }
      }
    }

    public bool IsWaitingForPacking {
      get {
        foreach (SalesOrderTaskLine sotl in this.lines) {
          if (sotl.QuantityOrdered > sotl.QuantityPicked) {
            return false;
          }
        }
        return true;
      }
    }

    public bool IsCancelled { get { return this.currentStatus == SalesOrderStatus.Cancelled; } }
    public bool IsHeld { get { return this.currentHoldReason != SalesOrderHoldReason.NoHold; } }

    public bool IsEasilyPickable {
      get {
        // this doesn't work well with backorders, but good enough for now
        bool foundNonzero = false;
        foreach (SalesOrderTaskLine sotl in this.lines) {
          if ((int)sotl.PickQuantityRequired > 0) {
            foundNonzero = true;
            if ((int)sotl.PickQuantityRequired > TP.Warehousing.Locations.Manager.GetTotalAvailableQuantityFor(sotl.ComponentID, this.warehouseID)) {
              return false;
            }
          }
        }
        return foundNonzero;
      }
    }

    public bool IsNotEasilyPickable {
      get {
        // note: despite the name, this is not a negation of the above
        bool foundNonzero = false;
        foreach (SalesOrderTaskLine sotl in this.lines) {
          if ((int)sotl.PickQuantityRequired > 0) {
            foundNonzero = true;
            int qtyTotal = TP.Warehousing.Locations.Manager.GetTotalAvailableQuantityFor(sotl.ComponentID);
            if ((int)sotl.PickQuantityRequired > qtyTotal) {
              return true;
            }
            int qtyHere = TP.Warehousing.Locations.Manager.GetTotalAvailableQuantityFor(sotl.ComponentID, this.warehouseID);
            if ((int)sotl.PickQuantityRequired > qtyHere) {
              return true;
            }
          }
        }
        return foundNonzero;
      }
    }

    public override string ToString() {
      return this.ToString(false);
    }
    public string ToString(bool inclType) {
      return (inclType ? "SO " : string.Empty) + this.salesOrderNumber;
    }

    public string ToTooltipHeaderText() {
      return string.Format(
          "Status: {0}\r\nShip Via: {1}\r\nShip To: {2}, {3}\r\n\r\n",
          this.CurrentStatusPrintable,
          this.ShipVia,
          this.ShipState,
          this.ShipCountry
        );
    }

    #endregion

    #region IEquatable<IWarehousingTask> Members

    public override bool Equals(object obj) {
      return base.Equals(obj);
    }

    public override int GetHashCode() {
      return base.GetHashCode();
    }

    public bool Equals(IWarehousingTask other) {
      if ((object)other == null) {
        return false;
      }
      return this.TransactionType == other.TransactionType && this.DatabaseID == other.DatabaseID;
    }

    public static bool operator ==(SalesOrderTask c1, IWarehousingTask c2) {
      if ((object)c1 == null) {
        return (object)c2 == null;
      }
      return c1.Equals(c2);
    }
    public static bool operator !=(SalesOrderTask c1, IWarehousingTask c2) {
      if ((object)c1 == null) {
        return (object)c2 != null;
      }
      return !c1.Equals(c2);
    }

    #endregion

    internal string getEnumDescription(System.Enum value) {
      System.Reflection.FieldInfo fi = value.GetType().GetField(value.ToString());
      System.ComponentModel.DescriptionAttribute[] attributes = (System.ComponentModel.DescriptionAttribute[])fi.GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), false);
      if (attributes.Length > 0) {
        return attributes[0].Description;
      }
      else {
        return value.ToString();
      }
    }
  }
}
