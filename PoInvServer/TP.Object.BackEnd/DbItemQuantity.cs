using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemQuantity", Namespace = "TP.Object")]
  public class DbItemQuantity : TP.Object.ItemQuantity
  {
    #region Cache

    private static AutoCache1Key1Index<DbItemQuantity, string, int> cache = new AutoCache1Key1Index<DbItemQuantity, string, int>
    {
      Constructor = _ => new DbItemQuantity(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<int>("ItemID"), _.Field<string>("WhseCode")), // matches .PrimaryKeyString and cache.Clear in .Create
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ItemID
    };

    #endregion


    #region Constructors

    public DbItemQuantity(DataRow r) : base()
    {
      this.ItemID = (int)r["ItemID"];
      this.WarehouseCode = (string)r["WhseCode"];
      this.OnHand = (int)r["QuantityOnHand"];
      this.OnPurchaseOrder = (int)r["QuantityOnPurchaseOrder"];
      this.OnSalesOrder = (int)r["QuantityOnSalesOrder"];
      this.OnBackOrder = (int)r["QuantityOnBackOrder"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.ItemID, this.WarehouseCode); } }


    #region Getters

    public static IEnumerable<DbItemQuantity> GetByItemID(int itemID, string[] preloads = null) { return GetByItemID(null, itemID, preloads); }
    internal static IEnumerable<DbItemQuantity> GetByItemID(System.Data.SqlClient.SqlTransaction txn, int itemID, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(itemID, () => Connection.Retrieve(txn, "SELECT ItemID, WhseCode, QuantityOnHand, QuantityOnPurchaseOrder, QuantityOnSalesOrder, QuantityOnBackOrder FROM InventoryQuantities WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, itemID)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbItemQuantity fillPreloads(DbItemQuantity qty, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (qty == null)
      {
        return qty;
      }
      var retval = (DbItemQuantity)qty.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbItemQuantity> fillPreloads(IEnumerable<DbItemQuantity> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // internal only, only called by item creation
    internal static bool CreateDefaultFor(int itemID, string itemNumber, System.Data.SqlClient.SqlTransaction txn)
    {
      DbItem.AssertItemExists(itemID, txn);

      // TODO: table structure: remove InventoryQuantities.ItemNumber
      Connection.Execute(txn,
          "INSERT INTO InventoryQuantities (ItemID, ItemNumber, WhseCode) VALUES (@iid, @item, '000')",
          new QueryParameter("@iid", SqlDbType.Int, itemID),
          new QueryParameter("@item", SqlDbType.VarChar, itemNumber)
        );

      cache.Clear(string.Format("{0}/{1}", itemID, "000"), itemID);
      return true;
    }

    // internal only, only called by nightly data import
    internal static bool CreateMasWarehouseQuantityLine(int itemID, string itemNumber, string warehouseCode, int onHand, int onPurchaseOrder, int onSalesOrder, int onBackOrder, System.Data.SqlClient.SqlTransaction txn)
    {
      DbItem.AssertItemExists(itemID, txn);

      // TODO: table structure: remove InventoryQuantities.ItemNumber
      Connection.Execute(txn,
          "INSERT INTO InventoryQuantities (ItemID, ItemNumber, WhseCode, QuantityOnHand, QuantityOnPurchaseOrder, QuantityOnSalesOrder, QuantityOnBackOrder) VALUES (@iid, @item, @w, @h, @p, @s, @b)",
          new QueryParameter("@iid", SqlDbType.Int, itemID),
          new QueryParameter("@item", SqlDbType.VarChar, itemNumber),
          new QueryParameter("@w", SqlDbType.VarChar, warehouseCode),
          new QueryParameter("@h", SqlDbType.Int, onHand),
          new QueryParameter("@pd", SqlDbType.Int, onPurchaseOrder),
          new QueryParameter("@s", SqlDbType.Int, onSalesOrder),
          new QueryParameter("@b", SqlDbType.Int, onBackOrder)
        );

      cache.Clear(string.Format("{0}/{1}", itemID, warehouseCode),  itemID);
      return true;
    }

    // internal only, updating is handled through Mas synchronization
    internal bool UpdateItemQuantity(System.Data.SqlClient.SqlTransaction txn)
    {
      DbItemQuantity orig = GetByItemID(txn, this.ItemID).Single(_ => _.WarehouseCode == this.WarehouseCode);

      Extensions.ObjectChanges changed = this.PropertiesModified(orig);
      if (changed.Count > 0) {
        int affected = Connection.Execute(txn,
            "UPDATE InventoryQuantities SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
            changed.UpdateValues
          );
        if (affected == 0) {
          throw new ApplicationException("error updating!");
        }
      }
      Connection.CommitTransaction(txn);

      cache.Clear(this.PrimaryKeyString, this.ItemID);
      return true;
    }

    // deleting not allowed

    #endregion

  }
}
