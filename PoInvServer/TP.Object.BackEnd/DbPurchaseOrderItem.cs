using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "PurchaseOrderItem", Namespace = "TP.Object")]
  public class DbPurchaseOrderItem : TP.Object.PurchaseOrderItem
  {
    #region Cache

    private static AutoCache1Key2Index<DbPurchaseOrderItem, int, int, int> cache = new AutoCache1Key2Index<DbPurchaseOrderItem, int, int, int> {
        Constructor = _ => new DbPurchaseOrderItem(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.HeaderID,
        Index2 = _ => _.ItemID
      };

    #endregion


    #region Constructors

    public DbPurchaseOrderItem(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.HeaderID = (int)r["HeaderID"];
      this.ItemID = (int)r["ItemID"];
      this.Quantity = (int)r["Quantity"];
    }

    #endregion


    #region Getters

    // not externally accessible
    public static DbPurchaseOrderItem GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbPurchaseOrderItem GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ItemID, Quantity FROM PurchaseOrderLines WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /purchaseorder/{id:int}/itemlines
    public static IEnumerable<DbPurchaseOrderItem> GetByPurchaseOrderID(int id, string[] preloads = null) { return GetByPurchaseOrderID(null, id, preloads); }
    internal static IEnumerable<DbPurchaseOrderItem> GetByPurchaseOrderID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ItemID, Quantity FROM PurchaseOrderLines WHERE HeaderID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /purchaseorder/itemlines/item/{id:int}/active
    public static IEnumerable<DbPurchaseOrderItem> GetCurrentByItemID(int id, string[] preloads = null) { return GetCurrentByItemID(null, id, preloads); }
    internal static IEnumerable<DbPurchaseOrderItem> GetCurrentByItemID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      DbItem.AssertItemExists(id, txn);
      var ret = cache.FetchByIndex2(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, ItemID, Quantity FROM PurchaseOrderLines WHERE ItemID=@id AND HeaderID IN (SELECT ID FROM PurchaseOrders WHERE Exported=0)", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbPurchaseOrderItem fillPreloads(DbPurchaseOrderItem ln, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (ln == null)
      {
        return ln;
      }
      var retval = (DbPurchaseOrderItem)ln.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("item"))
        {
          // always preloads costs
          retval.FillItem(txn);
        }
        if (preloads.Contains("purchaseorder"))
        {
          retval.FillPurchaseOrder(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbPurchaseOrderItem> fillPreloads(IEnumerable<DbPurchaseOrderItem> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbPurchaseOrderItem has-one DbItem
    public override bool FillItem() { return FillItem(null); }
    internal bool FillItem(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._item == null)
      {
        this._item = DbItem.GetByID(txn, this.ItemID, new string[] { "costs" });
      }
      return true;
    }

    // DbPurchaseOrderItem has-one DbPurchaseOrder
    public override bool FillPurchaseOrder() { return FillPurchaseOrder(null); }
    internal bool FillPurchaseOrder(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._po == null)
      {
        this._po = DbPurchaseOrder.GetByID(txn, this.HeaderID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    public static DbPurchaseOrderItem AddItemToPurchaseOrder(int poID, int itemID, int quantity)
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrderItem retval = AddItemToPurchaseOrder(txn, poID, itemID, quantity);
        Connection.CommitTransaction(txn);
        // cache handled already
        return retval;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }
    internal static DbPurchaseOrderItem AddItemToPurchaseOrder(System.Data.SqlClient.SqlTransaction txn, int poID, int itemID, int quantity)
    {
      DbItem.AssertItemExists(itemID, txn);
      DbPurchaseOrder.AssertPurchaseOrderExists(poID, txn);

      DbPurchaseOrder po = DbPurchaseOrder.GetForItem(txn, itemID).SingleOrDefault(_ => _.PurchaseOrderID == poID);
      if (po == null)
      {
        throw new ArgumentException("this item cannot go on this purchase order!");
      }

      int id = Connection.Insert(txn,
          "INSERT INTO PurchaseOrderLines (HeaderID, ItemNumber, ItemID, Quantity) VALUES (@id, @item, @iid, @qty)",
          new QueryParameter("@id", SqlDbType.Int, poID),
          new QueryParameter("@item", SqlDbType.VarChar, DbItem.GetByID(itemID).ItemNumber),
          new QueryParameter("@iid", SqlDbType.Int, itemID),
          new QueryParameter("@qty", SqlDbType.Int, quantity)
        );

      cache.Clear(id, poID, itemID);
      return GetByID(txn, id);
    }

    public DbPurchaseOrderItem UpdatePurchaseOrderItem()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrderItem orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE PurchaseOrderLines SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.HeaderID, this.ItemID);
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public bool RemovePurchaseOrderItem()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrderItem orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          throw new ArgumentException("item line has been modified, not deleting");
        }

        Connection.Execute(txn,
            "DELETE FROM PurchaseOrderLines WHERE ID=@id",
            new QueryParameter("@id", SqlDbType.Int, this.TupleID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.HeaderID, this.ItemID);
        return true;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


    #region Extra Validation

    // any?

    #endregion

  }
}
