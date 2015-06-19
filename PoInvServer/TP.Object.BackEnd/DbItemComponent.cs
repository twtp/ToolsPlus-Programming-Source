using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemComponent", Namespace = "TP.Object")]
  public class DbItemComponent : TP.Object.ItemComponent
  {
    #region Cache

    private static AutoCache1Key2Index<DbItemComponent, string, int, int> cache = new AutoCache1Key2Index<DbItemComponent, string, int, int>
    {
      Constructor = _ => new DbItemComponent(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<int>("ItemID"), _.Field<int>("ComponentID")), // matches .PrimaryKeyString
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ItemID,
      Index2 = _ => _.ComponentID
    };

    #endregion


    #region Constructors

    public DbItemComponent(DataRow r) : base()
    {
      this.ItemID = (int)r["ItemID"];
      this.ComponentID = (int)r["ComponentID"];
      this.Quantity = (int)r["Quantity"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.ItemID, this.ComponentID); } }


    #region Getters

    // /component/item/{id:int}
    public static IEnumerable<DbItemComponent> GetByItemID(int id, string[] preloads = null) { return GetByItemID(null, id, preloads); }
    internal static IEnumerable<DbItemComponent> GetByItemID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ItemID, ComponentID, Quantity FROM InventoryComponentMap WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /component/item/{id:int}
    public static IEnumerable<DbItemComponent> GetByComponentID(int id, string[] preloads = null) { return GetByComponentID(null, id, preloads); }
    internal static IEnumerable<DbItemComponent> GetByComponentID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex2(id, () => Connection.Retrieve(txn, "SELECT ItemID, ComponentID, Quantity FROM InventoryComponentMap WHERE ComponentID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // not publically available, but could be by something like /component/{id:int}/item/{iid:int}
    public static DbItemComponent GetByPrimaryKey(int itemID, int componentID, string[] preloads = null) { return GetByPrimaryKey(null, itemID, componentID, preloads); }
    internal static DbItemComponent GetByPrimaryKey(System.Data.SqlClient.SqlTransaction txn, int itemID, int componentID, string[] preloads = null)
    {
      return GetByItemID(txn, itemID, preloads).SingleOrDefault(_ => _.ComponentID == componentID);
    }

    private static DbItemComponent fillPreloads(DbItemComponent ic, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (ic == null)
      {
        return ic;
      }
      var retval = (DbItemComponent)ic.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("components"))
        {
          retval.FillComponent(preloads, txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbItemComponent> fillPreloads(IEnumerable<DbItemComponent> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillComponent(string[] preloads = null) { return FillComponent(preloads, null); }
    internal bool FillComponent(string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._component == null)
      {
        this._component = DbComponent.GetByID(txn, this.ComponentID, preloads); // always fill barcodes too
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // internal only, called from Item creation
    internal static bool CreateDefaultFor(int itemID, string itemNumber, System.Data.SqlClient.SqlTransaction txn)
    {
      // TODO: table structure: remove InventoryComponentMap.ItemNumber
      int cid = DbComponent.CreateDefaultFor(itemNumber, txn);
      Connection.Execute(txn,
          "INSERT INTO InventoryComponentMap (ItemID, ItemNumber, ComponentID, Quantity) VALUES (@iid, @item, @cid, 1)",
          new QueryParameter("@iid", SqlDbType.Int, itemID),
          new QueryParameter("@item", SqlDbType.VarChar, itemNumber),
          new QueryParameter("@cid", SqlDbType.Int, cid)
        );
      // doesn't add to cache? this doesn't actually load the object, so probably okay
      return true;
    }

    // /component/{id:int}/link/item/{iid:int}
    public static bool CreateItemComponentLink(int componentID, int itemID, int quantity)
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbItem.AssertItemExists(itemID, txn);
        DbItem item = DbItem.GetByID(txn, itemID);

        bool retval = CreateItemComponentLink(componentID, itemID, item.ItemNumber, quantity, txn);
        Connection.CommitTransaction(txn);
        return retval;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }
    internal static bool CreateItemComponentLink(int cid, int itemID, string itemNumber, int quantity, System.Data.SqlClient.SqlTransaction txn)
    {
      DbItem.AssertItemExists(itemID, txn);
      DbComponent.AssertComponentExists(cid, txn);

      Connection.Execute(txn,
          "INSERT INTO InventoryComponentMap (ItemID, ItemNumber, ComponentID, Quantity) VALUES (@iid, @item, @cid, @qty)",
          new QueryParameter("@iid", SqlDbType.Int, itemID),
          new QueryParameter("@item", SqlDbType.VarChar, itemNumber),
          new QueryParameter("@cid", SqlDbType.Int, cid),
          new QueryParameter("@qty", SqlDbType.Int, quantity)
        );
      // doesn't add to cache? this doesn't actually load the object, so probably okay
      return true;
    }

    // /component/{id:int}/item/{iid:int}/update
    public DbItemComponent UpdateItemComponentLink()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbItemComponent orig = GetByPrimaryKey(txn, this.ItemID, this.ComponentID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE InventoryComponentMap SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.PrimaryKeyString, this.ItemID, this.ComponentID); // item/component are immutable, easy
        return GetByPrimaryKey(this.ItemID, this.ComponentID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /component/{id:int}/unlink/item/{iid:int}
    public bool DeleteItemComponentLink()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbItemComponent orig = GetByPrimaryKey(txn, this.ItemID, this.ComponentID);
        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          throw new ArgumentException("item/component link has been modified, not deleting");
        }

        Connection.Execute(txn,
            "DELETE FROM InventoryComponentMap WHERE ItemID=@iid AND ComponentID=@cid",
            new QueryParameter("@iid", SqlDbType.Int, this.ItemID),
            new QueryParameter("@cid", SqlDbType.Int, this.ComponentID)
          );
        DataTable dt = Connection.Retrieve(txn, "SELECT COUNT(*) FROM InventoryComponentMap WHERE ComponentID=@cid", new QueryParameter("@cid", SqlDbType.Int, this.ComponentID));
        if ((int)dt.Rows[0][0] == 0)
        {
          DbComponent.GetByID(txn, this.ComponentID).DeleteComponent(txn);
        }

        cache.Clear(this.PrimaryKeyString, this.ItemID, this.ComponentID);
        Connection.CommitTransaction(txn);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


  }
}
