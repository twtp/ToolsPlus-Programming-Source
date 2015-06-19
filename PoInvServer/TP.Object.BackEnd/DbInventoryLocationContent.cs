using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "InventoryLocationContent", Namespace = "TP.Object")]
  public class DbInventoryLocationContent : TP.Object.InventoryLocationContent
  {
    #region Cache

    private static AutoCache1Key2Index<DbInventoryLocationContent, string, int, int> cache = new AutoCache1Key2Index<DbInventoryLocationContent, string, int, int>
    {
      Constructor = _ => new DbInventoryLocationContent(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<int>("LocationID"), _.Field<int>("ComponentID")), // matches .PrimaryKeyString
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.LocationID,
      Index2 = _ => _.ComponentID,
      //Index3 = _ => _.MasterID,
      SecondsToLive = 5 // exceptionally short term cache, just for a single request, really
    };

    #endregion


    #region Constructors

    public DbInventoryLocationContent(DataRow r) : base()
    {
      this.LocationID = (int)r["LocationID"];
      this.ComponentID = (int)r["ComponentID"];
      this.MasterID = (int)r["MasterID"];
      this.Quantity = this.nullable<int>(r["Quantity"]);
      this.AssociatedTaskType = this.nullable<int>(r["AssociatedTaskType"]);
      this.AssociatedTaskID = this.nullable<int>(r["AssociatedTaskID"]);
      this.LastInventoriedDate = (DateTime)r["LastInventoriedDate"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.LocationID, this.ComponentID); } }


    #region Getters

    // /invlocation/{id:int}/contents
    public static IEnumerable<DbInventoryLocationContent> GetByLocationID(int id, string[] preloads = null) { return GetByLocationID(null, id, preloads); }
    internal static IEnumerable<DbInventoryLocationContent> GetByLocationID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT LocationID, ComponentID, MasterID, Quantity, AssociatedTaskType, AssociatedTaskID, LastInventoriedDate FROM LocationContents WHERE LocationID=@lid", new QueryParameter("@lid", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /component/{id:int}/locations
    public static IEnumerable<DbInventoryLocationContent> GetByComponentID(int id, string[] preloads = null) { return GetByComponentID(null, id, preloads); }
    internal static IEnumerable<DbInventoryLocationContent> GetByComponentID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex2(id, () => Connection.Retrieve(txn, "SELECT LocationID, ComponentID, MasterID, Quantity, AssociatedTaskType, AssociatedTaskID, LastInventoriedDate FROM LocationContents WHERE ComponentID=@cid", new QueryParameter("@cid", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbInventoryLocationContent> GetByMasterID(int id, string[] preloads = null) { return GetByMasterID(null, id, preloads); }
    internal static IEnumerable<DbInventoryLocationContent> GetByMasterID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT LocationID, ComponentID, MasterID, Quantity, AssociatedTaskType, AssociatedTaskID, LastInventoriedDate FROM LocationContents WHERE MasterID=@mid", new QueryParameter("@mid", SqlDbType.Int, id)));
      //var ret = cache.FetchByIndex3(id => Connection.Retrieve(txn, "SELECT LocationID, ComponentID, MasterID, Quantity, AssociatedTaskType, AssociatedTaskID, LastInventoriedDate FROM LocationContents WHERE MasterID=@mid", new QueryParameter("@mid", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbInventoryLocationContent fillPreloads(DbInventoryLocationContent c, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (c == null)
      {
        return c;
      }
      var retval = (DbInventoryLocationContent)c.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("location"))
        {
          retval.FillInventoryLocation(txn);
        }
        if (preloads.Contains("component"))
        {
          retval.FillComponent(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbInventoryLocationContent> fillPreloads(IEnumerable<DbInventoryLocationContent> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillComponent() { return FillComponent(null); }
    internal bool FillComponent(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._component == null)
      {
        this._component = DbComponent.GetByID(txn, this.ComponentID);
      }
      return true;
    }

    public override bool FillInventoryLocation() { return FillInventoryLocation(null); }
    internal bool FillInventoryLocation(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._location == null)
      {
        this._location = DbInventoryLocation.GetByID(txn, this.LocationID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // one-sided update with a task: final step to remove inventory for tasks (will call, s/o), etc
    // /task/{type}/id/{id:int}/line/{lineid:int}/change
    public static bool AddOrRemove(int startLocationID, int endLocationID, IWarehouseTaskLine taskLine, int deltaQuantity, string comment, string personNTUsername, string computerName, string rdpClientName)
    {
      if (deltaQuantity == 0) { throw new ApplicationException("seriously, wtf."); }

      var txn = Connection.BeginTransaction();
      try
      {
        if (taskLine.HasChangedFromDbCurrent(txn))
        {
          throw new ArgumentException("task line has been modified!");
        }
        var l1 = changeContents(txn, startLocationID, endLocationID, taskLine.ComponentID, deltaQuantity, comment, taskLine.Header(txn).TransactionReference, taskLine.Header(txn).TransactionTypeID, taskLine.Header(txn).TransactionID, personNTUsername, computerName, rdpClientName);
        var tl = taskLine.UpdateQuantity(txn, startLocationID, null, deltaQuantity);
        Connection.CommitTransaction(txn);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /invlocation/{id:int}/contents/{cid:int}/change
    // /component/{id:int}/locations/{lid:int}/change
    // one-sided update without a task: receiving, etc
    public static bool AddOrRemove(int startLocationID, int endLocationID, int componentID, int deltaQuantity, string transactionReference, string comment, string personNTUsername, string computerName, string rdpClientName)
    {
      if (deltaQuantity == 0) { throw new ApplicationException("seriously, wtf."); }

      var txn = Connection.BeginTransaction();
      try
      {
        var ret = changeContents(txn, startLocationID, endLocationID, componentID, deltaQuantity, comment, transactionReference, null, null, personNTUsername, computerName, rdpClientName);
        Connection.CommitTransaction(txn);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // two-sided update with a task: toting, packing, etc
    // /task/{type}/id/{id:int}/line/{lineid:int}/move
    public static bool MoveFromTo(int fromStartLocationID, int fromEndLocationID, int toStartLocationID, int toEndLocationID, IWarehouseTaskLine taskLine, int deltaQuantity, string comment, string personNTUsername, string computerName, string rdpClientName)
    {
      if (deltaQuantity <= 0) { throw new ApplicationException("seriously, wtf."); }

      var txn = Connection.BeginTransaction();
      try
      {
        if (taskLine.HasChangedFromDbCurrent(txn))
        {
          throw new ArgumentException("task line has been modified!");
        }
        var lc1 = changeContents(txn, fromStartLocationID, fromEndLocationID, taskLine.ComponentID, -1 * deltaQuantity, comment, taskLine.Header(txn).TransactionReference, taskLine.Header(txn).TransactionTypeID, taskLine.Header(txn).TransactionID, personNTUsername, computerName, rdpClientName);
        var lc2 = changeContents(txn, toStartLocationID, toEndLocationID, taskLine.ComponentID, deltaQuantity, comment, taskLine.Header(txn).TransactionReference, taskLine.Header(txn).TransactionTypeID, taskLine.Header(txn).TransactionID, personNTUsername, computerName, rdpClientName);
        var tl = taskLine.UpdateQuantity(txn, fromStartLocationID, toStartLocationID, deltaQuantity);
        Connection.CommitTransaction(txn);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /invlocation/{id:int}/contents/{cid:int}/move
    // /component/{id:int}/locations/{lid:int}/move
    // two-sided update without a task: moving between locations, etc
    public static bool MoveFromTo(int fromStartLocationID, int fromEndLocationID, int toStartLocationID, int toEndLocationID, int componentID, int deltaQuantity, string transactionReference, string comment, string personNTUsername, string computerName, string rdpClientName)
    {
      if (deltaQuantity <= 0) { throw new ApplicationException("seriously, wtf."); }

      var txn = Connection.BeginTransaction();
      try
      {
        var lc1 = changeContents(txn, fromStartLocationID, fromEndLocationID, componentID, -1 * deltaQuantity, comment, transactionReference, null, null, personNTUsername, computerName, rdpClientName);
        var lc2 = changeContents(txn, toStartLocationID, toEndLocationID, componentID, deltaQuantity, comment, transactionReference, null, null, personNTUsername, computerName, rdpClientName);
        Connection.CommitTransaction(txn);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // internal only, heavy lifting to do contents change
    internal static DbInventoryLocationContent changeContents(System.Data.SqlClient.SqlTransaction txn, int startLocationID, int endLocationID, int componentID, int deltaQuantity, string comment, string transactionReference, int? transactionTypeID, int? transactionID, string personNTUsername, string computerName, string rdpClientName)
    {
      if (transactionTypeID.HasValue == true && transactionID.HasValue == false) { throw new ApplicationException("invalid transaction type id / transaction id"); }
      if (transactionTypeID.HasValue == false && transactionID.HasValue == true) { throw new ApplicationException("invalid transaction type id / transaction id"); }
      if (deltaQuantity == 0) { throw new ApplicationException("invalid quantity"); }

      DbInventoryLocation.AssertLocationExists(startLocationID, txn);
      DbInventoryLocation.AssertLocationExists(endLocationID, txn);
      DbComponent.AssertComponentExists(componentID, txn);
      DbPerson person = DbPerson.FindOrCreatePerson(txn, personNTUsername);
      DbComputer computer = DbComputer.FindOrCreateComputer(txn, computerName);
      DbComputer rdpClient = DbComputer.FindOrCreateComputer(txn, rdpClientName);
      /* ASSERT TASK EXISTS? */

      // check span validity
      var l1 = DbInventoryLocation.GetByID(txn, startLocationID);
      var l2 = DbInventoryLocation.GetByID(txn, endLocationID);
      l1.FillContents(txn);
      InventoryLocation[] spanList;
      if (l1.Contents.Count() == 0)
      {
        // l1 is empty, make sure it can span to l2
        spanList = l1.GetPotentialSpanMembers(l2, _ => _.FillContents(txn), _ => _.FillNeighbors(txn));
        if (spanList == null) { throw new ArgumentException("span is invalid!"); }
      }
      else
      {
        // l1 is full, make sure l2 is the correct endpoint
        if (l1.LocationID != l1.Contents.First().MasterID) {
          throw new ArgumentException(string.Format("startLocationID {0} does not match existing span startpoint {1}", l1.LocationID, l1.Contents.First().MasterID));
        }
        var end = DbInventoryLocation.GetSpanEndpoint(txn, l1.LocationID);
        if (end.LocationID != l2.LocationID)
        {
          throw new ArgumentException(string.Format("endLocationID {0} does not match existing span endpoint {1}", l2.LocationID, end.LocationID));
        }
        spanList = l1.GetCurrentSpanMembers(
            l2,
            _ => _.FillContents(txn),
            _ => _.FillNeighbors(txn),
            _ => DbInventoryLocationContent.GetByMasterID(txn, _)
          );
      }

      // verify that this update can be made with/without this transaction info
      var contentsList = DbInventoryLocationContent.GetByLocationID(txn, startLocationID);
      if (l1.IsTransactionRequired == true)
      {
        if (transactionTypeID.HasValue == false)
        {
          throw new ArgumentException("location must have an associated task");
        }
        if (contentsList.Any(_ => _.AssociatedTaskType.Value != transactionTypeID.Value || _.AssociatedTaskID.Value != transactionID.Value))
        {
          throw new ArgumentException("location is already assigned to a different task!");
        }
      }
      else
      {
        // something like untoting, transaction required for from location, but not to location,
        // so we can wipe this out here
        if (transactionTypeID.HasValue == true)
        {
          transactionTypeID = null;
          transactionID = null;
        }
      }
      if (transactionTypeID == null && l1.IsTypeStandard == false)
      {
        throw new ArgumentException("must have a task to move from/to this location");
      }
      
      // now get the quantity for the component being changed, verify that it's valid
      var existingContents = contentsList.SingleOrDefault(_ => _.ComponentID == componentID);
      var newQ = (existingContents == null ? deltaQuantity : (existingContents.Quantity + deltaQuantity));
      if (newQ < 0 && spanList[0].IsStrictlyManaged)
      {
        throw new ArgumentException("quantity cannot go below zero");
      }

      // finally, do the insert/update/delete to adjust quantity
      if (existingContents == null)
      {
        for (int i = 0; i < spanList.Length; i++)
        {
          Connection.Execute(txn,
              "INSERT INTO LocationContents (LocationID, ComponentID, MasterID, Quantity, AssociatedTaskType, AssociatedTaskID) VALUES (@l, @c, @m, @q, @t1, @t2)",
              new QueryParameter("@l", SqlDbType.Int, spanList[i].LocationID),
              new QueryParameter("@c", SqlDbType.Int, componentID),
              new QueryParameter("@m", SqlDbType.Int, spanList[0].LocationID),
              new QueryParameter("@q", SqlDbType.Int, i == 0 ? newQ : null),
              new QueryParameter("@t1", SqlDbType.Int, transactionTypeID),
              new QueryParameter("@t2", SqlDbType.Int, transactionID)
            );
        }
      }
      else if (newQ == 0)
      {
        Connection.Execute(txn,
            "DELETE FROM LocationContents WHERE MasterID=@m AND ComponentID=@c",
            new QueryParameter("@m", SqlDbType.Int, spanList[0].LocationID),
            new QueryParameter("@c", SqlDbType.Int, componentID)
          );
      }
      else
      {
        Connection.Execute(txn,
            "UPDATE LocationContents SET Quantity=Quantity+@q WHERE LocationID=@l AND ComponentID=@c",
            new QueryParameter("@q", SqlDbType.Int, deltaQuantity),
            new QueryParameter("@l", SqlDbType.Int, spanList[0].LocationID),
            new QueryParameter("@c", SqlDbType.Int, componentID)
          );
      }

      // and add to the log
      Connection.Execute(txn,
          "INSERT INTO WarehouseTransactionLog (LocationID, TransactionType, TransactionReference, ComponentID, QuantityChange, UserID, ComputerID, RDPClientID, Comment) VALUES (@lid, NULL, @ref, @cid, @qty, @uid, @comp, @rdp, @comm)",
          new QueryParameter("@lid", SqlDbType.Int, spanList[0].LocationID),
          new QueryParameter("@ref", SqlDbType.VarChar, transactionReference),
          new QueryParameter("@cid", SqlDbType.Int, componentID),
          new QueryParameter("@qty", SqlDbType.Int, deltaQuantity),
          new QueryParameter("@uid", SqlDbType.Int, person == null ? -1 : person.PersonID),
          new QueryParameter("@comp", SqlDbType.Int, computer == null ? -1 : computer.ComputerID),
          new QueryParameter("@rdp", SqlDbType.Int, rdpClient == null ? -1 : rdpClient.ComputerID),
          new QueryParameter("@comm", SqlDbType.VarChar, comment)
        );

      foreach (var l in spanList)
      {
        cache.Clear(string.Format("{0}/{1}", l.LocationID, componentID), l.LocationID, componentID);
      }
      return GetByLocationID(txn, startLocationID).SingleOrDefault(_ => _.ComponentID == componentID);
    }

    // /{id:int}/location/{lid:int}/inventoried
    public static void UpdateInventoriedDate(int locationID, int componentID, int checkQuantity)
    {
      var txn = Connection.BeginTransaction();
      try
      {
        var loc = GetByComponentID(txn, componentID).SingleOrDefault(_ => _.LocationID == locationID);

        if (loc == null) { throw new ArgumentException("component does not exist in location"); }
        if (loc.Quantity != checkQuantity) { throw new ArgumentException("check quantity does not match current quantity"); }

        Connection.Execute(txn,
            "UPDATE LocationContents SET LastInventoriedDate=GETDATE() WHERE MasterID=@lid AND ComponentID=@cid",
            new QueryParameter("@lid", SqlDbType.Int, loc.MasterID),
            new QueryParameter("@cid", SqlDbType.Int, componentID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(string.Format("{0}/{1}", locationID, componentID), locationID, componentID);
        return;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


    #region Extra Validation

    // none

    #endregion

  }
}
