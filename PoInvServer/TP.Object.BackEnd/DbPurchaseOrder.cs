using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "PurchaseOrder", Namespace = "TP.Object")]
  public class DbPurchaseOrder : TP.Object.PurchaseOrder
  {
    #region Cache

    // no autocache, special exported flag handling
    private static Cache cache = new Cache();
    private class Cache : CacheBase2Key1Index<DbPurchaseOrder, int, string, int>
    {
      public DbPurchaseOrder FetchByID(int id, Func<DataTable> getter)
      {
        if (!base.contains(id))
        {
          DataTable dt = getter();
          if (dt.Rows.Count == 1)
          {
            var obj = new DbPurchaseOrder(dt.Rows[0]);
            // ugh, this is awful. need to conditionally add to indexes, based
            // on whether it has a purchase order number assigned, and if it
            // has been exported or not.
            base.put(
                obj.PurchaseOrderID,
                obj.PurchaseOrderNo, // can be null for non-assigned
                obj.IsExported ? false : true,
                obj.IsExported ? -1 : 0,
                obj
              );
            /*if (obj.PurchaseOrderNo != null)
            {
              if (obj.IsExported == false)
              {
                base.put(obj.PurchaseOrderID, obj.PurchaseOrderNo, true, 0, obj);
              }
              else
              {
                base.put(obj.PurchaseOrderID, obj.PurchaseOrderNo, false, -1, obj);
              }
            }
            else
            {
              if (obj.IsExported == false)
              {
                base.put(obj.PurchaseOrderID, null, true, 0, obj);
              }
              else
              {
                // exported with no po number, this should never occur.
                base.put(obj.PurchaseOrderID, null, false, -1, obj);
              }
            }*/
          }
          else
          {
            base.put(id, null);
          }
        }
        return base.retrieve(id);
      }

      public DbPurchaseOrder FetchByPurchaseOrderNo(string purchaseOrderNo, Func<DataTable> getter)
      {
        if (!base.containsAlternate(purchaseOrderNo))
        {
          DataTable dt = getter();
          if (dt.Rows.Count == 1)
          {
            int tupleID = (int)dt.Rows[0]["ID"];
            var obj = base.contains(tupleID) ? base.retrieve(tupleID) : new DbPurchaseOrder(dt.Rows[0]);
            base.put(
                obj.PurchaseOrderID,
                obj.PurchaseOrderNo,
                obj.IsExported ? false : true,
                obj.IsExported ? -1 : 0,
                obj
              );
            /*if (obj.IsExported == false)
            {
              base.put(obj.PurchaseOrderID, obj.PurchaseOrderNo, 0, obj);
            }
            else
            {
              base.put(obj.PurchaseOrderID, obj.PurchaseOrderNo, obj);
            }*/
          }
          else
          {
            base.putAlternateMiss(purchaseOrderNo);
          }
        }
        return base.retrieveAlternate(purchaseOrderNo);
      }

      public IEnumerable<DbPurchaseOrder> FetchAll(Func<DataTable> getter)
      {
        if (!base.containsIndex1(0))
        {
          base.startIndex1(0);
          DataTable dt = getter();
          foreach (DataRow r in dt.Rows)
          {
            int tupleID = (int)r["ID"];
            DbPurchaseOrder obj = base.contains(tupleID) ? base.retrieve(tupleID) : new DbPurchaseOrder(r);
            /*if (obj.PurchaseOrderNo != null)
            {
              base.put(obj.PurchaseOrderID, obj.PurchaseOrderNo, 0, obj);
            }
            else
            {
              base.put(obj.PurchaseOrderID, null, 0, obj);
            }*/
            base.put(
                obj.PurchaseOrderID,
                obj.PurchaseOrderNo, // can be null
                obj.IsExported ? false : true, // shouldn't ever be exported here
                obj.IsExported ? -1 : 0,
                obj
              );
          }
        }
        return base.retrieveIndex1(0);
      }

      public bool Contains(int purchaseOrderID)
      {
        return base.contains(purchaseOrderID);
      }
      public bool ContainsByUsername(string purchaseOrderNo)
      {
        return base.containsAlternate(purchaseOrderNo);
      }
      public bool ContainsAll()
      {
        return base.containsIndex1(0);
      }

      public void Clear(int purchaseOrderID, string purchaseOrderNo)
      {
        base.clear(purchaseOrderID, purchaseOrderNo, 0);
      }
    }

    #endregion


    #region Constructors

    public DbPurchaseOrder(DataRow r) : base()
    {
      this.PurchaseOrderID = (int)r["ID"];
      this.PurchaseOrderNo = this.nullableString(r["PONumber"]);
      this.PaymentTerms = (string)r["POTerms"];
      this.VendorNo = (string)r["VendorNumber"];
      this.CoopVendorNo = (string)r["CoopVendor"];
      this.ProductLine = (string)r["ProductLineCode"];
      this.DueDate = (DateTime)r["DueDate"];
      this.IsFreightFree = (bool)r["FreightFree"];
      this.IsFinalized = (bool)r["Finalized"];
      this.IsExported = (bool)r["Exported"];
      this.ExportedDate = this.nullable<DateTime>(r["DateExported"]);
      this.Comment = this.nullableString(r["POComment"]);
      this.ShipToCode = (string)r["ShipToCode"];
      this.ShipToName = (string)r["ShipToName"];
      this.ShipToAddress1 = (string)r["ShipToAddress1"];
      this.ShipToAddress2 = (string)r["ShipToAddress2"];
      this.ShipToCity = (string)r["ShipToCity"];
      this.ShipToState = (string)r["ShipToState"];
      this.ShipToPostalCode = (string)r["ShipToZip"];
      this.VendorPOAddress = this.nullableString(r["VendorPOAddr"]);
      this.NumberOfPayments = (int)r["NumberOfPayments"];
      this.MiscellaneousNotes = this.nullableString(r["MiscNotes"]);
      this.DropshipSalesOrderNo = this.nullableString(r["SalesOrderNo"]);
      this.DropshipShipToTelephoneNo = this.nullableString(r["ShipToTelephoneNo"]);
      this.DropshipCustomerEmailAddress = this.nullableString(r["CustomerEmailAddress"]);
      this.DropshipIsLiftGateRequired = this.nullable<bool>(r["LiftGateRequired"]);
      this.DropshipAddressZoningType = this.nullable<int>(r["AddressZoningType"]);
      this.VendorOrderNo = this.nullableString(r["VendorOrderNo"]);
      this.VendorShippingStatus = (int)r["ShippingStatus"];
    }

    #endregion


    #region Getters

    // /purchaseorder/{id:int}
    public static DbPurchaseOrder GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbPurchaseOrder GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByID(id, () => Connection.Retrieve(txn, "SELECT ID, PONumber, POTerms, VendorNumber, CoopVendor, ProductLineCode, DueDate, FreightFree, Finalized, Exported, DateExported, POComment, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, VendorPOAddr, NumberOfPayments, MiscNotes, SalesOrderNo, ShipToTelephoneNo, CustomerEmailAddress, LiftGateRequired, AddressZoningType, VendorOrderNo, ShippingStatus FROM PurchaseOrders WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /purchaseorder/{po}
    public static DbPurchaseOrder GetByPurchaseOrderNo(string purchaseOrderNo, string[] preloads = null) { return GetByPurchaseOrderNo(null, purchaseOrderNo, preloads); }
    internal static DbPurchaseOrder GetByPurchaseOrderNo(System.Data.SqlClient.SqlTransaction txn, string purchaseOrderNo, string[] preloads = null)
    {
      purchaseOrderNo = purchaseOrderNo.ToUpper();

      var ret = cache.FetchByPurchaseOrderNo(purchaseOrderNo, () => Connection.Retrieve(txn, "SELECT ID, PONumber, POTerms, VendorNumber, CoopVendor, ProductLineCode, DueDate, FreightFree, Finalized, Exported, DateExported, POComment, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, VendorPOAddr, NumberOfPayments, MiscNotes, SalesOrderNo, ShipToTelephoneNo, CustomerEmailAddress, LiftGateRequired, AddressZoningType, VendorOrderNo, ShippingStatus FROM PurchaseOrders WHERE PONumber=@po", new QueryParameter("@po", SqlDbType.VarChar, purchaseOrderNo)));
      return fillPreloads(ret, preloads, txn);
    }

    // /purchaseorder/item/{id:int}
    public static IEnumerable<DbPurchaseOrder> GetForItem(int itemID, string[] preloads = null) { return GetForItem(null, itemID, preloads); }
    internal static IEnumerable<DbPurchaseOrder> GetForItem(System.Data.SqlClient.SqlTransaction txn, int itemID, string[] preloads = null)
    {
      Tuple<string, string, string> keys = getPurchaseOrderKeysByItemID(itemID, txn);

      var list = GetAllActive(txn).Where(_ =>
          keys.Item1 == _.VendorNo &&
          keys.Item2 == _.CoopVendorNo &&
          keys.Item3 == _.ProductLine
        ).ToList();

      return fillPreloads(list, preloads, txn);
    }

    // /purchaseorder/active
    public static IEnumerable<DbPurchaseOrder> GetAllActive(string[] preloads = null) { return GetAllActive(null, preloads); }
    internal static IEnumerable<DbPurchaseOrder> GetAllActive(System.Data.SqlClient.SqlTransaction txn, string[] preloads = null)
    {
      var ret = cache.FetchAll(() => Connection.Retrieve(txn, "SELECT ID, PONumber, POTerms, VendorNumber, CoopVendor, ProductLineCode, DueDate, FreightFree, Finalized, Exported, DateExported, POComment, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, VendorPOAddr, NumberOfPayments, MiscNotes, SalesOrderNo, ShipToTelephoneNo, CustomerEmailAddress, LiftGateRequired, AddressZoningType, VendorOrderNo, ShippingStatus FROM PurchaseOrders WHERE Exported=0"));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbPurchaseOrder> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbPurchaseOrder> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = GetAllActive(txn).Where(po => po.IsFinalized);
      //var ret = GetAllActive(txn).Where(po => po.PurchaseOrderNo != null);
      //var ret = new List<DbPurchaseOrder>();
      //ret.Add(GetByPurchaseOrderNo(txn, "A35159"));
      //ret.Add(GetByPurchaseOrderNo(txn, "A35156"));
      return fillPreloads(ret, new string[] { "items", "item", "costs" }, txn);
    }

    private static DbPurchaseOrder fillPreloads(DbPurchaseOrder po, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (po == null)
      {
        return po;
      }
      var retval = (DbPurchaseOrder)po.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("vendor"))
        {
          retval.FillVendor(txn);
        }
        if (preloads.Contains("items"))
        {
          retval.FillPurchaseOrderItems(txn, preloads.Contains("item"));
        }
        if (preloads.Contains("tracking"))
        {
          retval.FillPurchaseOrderTracking(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbPurchaseOrder> fillPreloads(IEnumerable<DbPurchaseOrder> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillVendor() { return FillVendor(null); }
    internal bool FillVendor(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._vendor == null)
      {
        this._vendor = DbVendor.GetByID(txn, this.VendorNo);
      }
      return true;
    }

    public override bool FillPurchaseOrderItems() { return FillPurchaseOrderItems(null); }
    internal bool FillPurchaseOrderItems(System.Data.SqlClient.SqlTransaction txn, bool preloadItemObjects = false)
    {
      if (this._items == null)
      {
        this._items = DbPurchaseOrderItem.GetByPurchaseOrderID(txn, this.PurchaseOrderID, preloadItemObjects ? new string[] { "item" } : null);
      }
      return true;
    }

    public override bool FillPurchaseOrderTracking() { return FillPurchaseOrderTracking(null); }
    internal bool FillPurchaseOrderTracking(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._tracking == null)
      {
        this._tracking = DbPurchaseOrderTracking.GetByPurchaseOrderID(txn, this.PurchaseOrderID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /purchaseorder/create/stock
    public static DbPurchaseOrder CreatePurchaseOrderForStock(int initialItemID, int initialItemQuantity)
    {
      var txn = Connection.BeginTransaction();
      int id = -1;
      try
      {
        DbItem.AssertItemExists(initialItemID, txn);
        var keys = getPurchaseOrderKeysByItemID(initialItemID, txn);

        DbVendor vendor = DbVendor.GetByID(txn, keys.Item1, new string[] { "terms" });

        id = Connection.Insert(txn,
            "INSERT INTO PurchaseOrders (VendorNumber, CoopVendor, ProductLineCode, POTerms) VALUES (@v, @c, @p, @t)",
            new QueryParameter("@v", SqlDbType.VarChar, keys.Item1),
            new QueryParameter("@c", SqlDbType.VarChar, keys.Item2),
            new QueryParameter("@p", SqlDbType.VarChar, keys.Item3),
            new QueryParameter("@t", SqlDbType.VarChar, vendor.PurchaseTerms == null ? DbPurchaseTerms.DefaultTermsCode() : vendor.PurchaseTerms.Code)
          );
        DbPurchaseOrderItem.AddItemToPurchaseOrder(txn, id, initialItemID, initialItemQuantity);
        cache.Clear(id, null);
        Connection.CommitTransaction(txn);

        return GetByID(id, new string[] { "items" });
      }
      catch (Exception)
      {
        // since AddItemToPurchaseOrder could fail, add some cache validation here
        if (id != -1)
        {
          cache.Clear(id, null);
        }
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // TODO: create purchase order by sales order (dropship)

    // /purchaseorder/update/{id:int}
    public DbPurchaseOrder UpdatePurchaseOrder()
    {
      // XXX: exporting a purchase order makes a lot of changes, make sure that
      // the cache for DbPurchaseOrderItem objects is consistent before and after!
      // also, we should do something about the ship to code and addresses. for
      // stock orders, maybe lock out the dropship option and vice-versa? in the
      // gui, the address field should only be enabled for editing when it's a
      // dropship (dummy shiptocode 1234).
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrder orig = GetByID(txn, this.PurchaseOrderID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE PurchaseOrders SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.PurchaseOrderID, this.PurchaseOrderNo);
        return GetByID(this.PurchaseOrderID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /purchaseorder/delete/{id:int}
    public bool DeletePurchaseOrder()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrder orig = GetByID(txn, this.PurchaseOrderID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          throw new ArgumentException("order has been modified, not deleting");
        }

        if (orig.IsExported)
        {
          throw new ArgumentException("can't delete an exported purchase order!");
        }

        orig.FillPurchaseOrderItems(txn);
        if (orig.PurchaseOrderItems.Count() > 0)
        {
          throw new ArgumentException("can't delete an order with lines");
        }
        orig.FillPurchaseOrderTracking(txn);
        if (orig.PurchaseOrderTracking.Count() > 0)
        {
          throw new ArgumentException("can't delete an order with tracking");
        }

        Connection.Execute(txn,
            "DELETE FROM PurchaseOrders WHERE ID=@id",
            new QueryParameter("@id", SqlDbType.Int, orig.PurchaseOrderID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(this.PurchaseOrderID, this.PurchaseOrderNo);
        return true;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public override bool FlagExported(Func<bool> recomparer, params string[] options)
    {
      // options is unused

      // no transaction, no checking. purchase orders do no overwrite themselves
      // well, so if this is successful according to the calling function,
      // assume that they're correct.
      Connection.Execute(
            "UPDATE PurchaseOrders SET Exported=1 WHERE ID=@id",
            new QueryParameter("@id", SqlDbType.Int, this.PurchaseOrderID)
          );

      cache.Clear(this.PurchaseOrderID, this.PurchaseOrderNo);
      return true;
    }

    #endregion


    #region Extra Validation

    internal static void AssertPurchaseOrderExists(int poID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbPurchaseOrder po = GetByID(txn, poID);
      if (po == null)
      {
        throw new ArgumentException("purchase order does not exist");
      }
      return;
    }

    #endregion


    private static Tuple<string, string, string> getPurchaseOrderKeysByItemID(int itemID, System.Data.SqlClient.SqlTransaction txn)
    {
      DbItem.AssertItemExists(itemID, txn);
      var item = DbItem.GetByID(txn, itemID);
      DbProductLine.AssertProductLineExists(item.ProductLinePrefix(), txn);
      var pl = DbProductLine.GetByPrefix(txn, item.ProductLinePrefix(), new string[] { "purchasingsubs" });

      string vendorNoToFind, coopVendorNoToFind, productLineToFind;
      vendorNoToFind = item.PrimaryVendor;
      coopVendorNoToFind = pl.IsCoopVendor ? pl.CoopVendor : item.PrimaryVendor;
      if (item.IsTypeDropship())
      {
        productLineToFind = "ZZZ";
      }
      else if (pl.PurchasingSubstituteProductLine != null)
      {
        pl.PurchasingSubstituteProductLine.FillPurchasingProductLine();
        productLineToFind = pl.PurchasingSubstituteProductLine.PurchasingProductLine.ProductLinePrefix;
      }
      else
      {
        productLineToFind = item.ProductLinePrefix();
      }

      return new Tuple<string, string, string>(vendorNoToFind, coopVendorNoToFind, productLineToFind);
    }


    public static string DUMP { get { return cache.Dump(); } }




    

  }
}
