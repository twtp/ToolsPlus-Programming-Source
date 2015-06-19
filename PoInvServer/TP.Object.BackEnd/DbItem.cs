using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "BaseItem", Namespace = "TP.Object")]
  public class DbBaseItem : TP.Object.BaseItem
  {
    #region Constructors

    public DbBaseItem(DataRow r) : base()
    {
      this.ItemID = (int)r["ItemID"];
      this.ItemNumber = (string)r["ItemNumber"];
      this.IsDiscontinued = (bool)r["IsDiscontinued"];
    }

    #endregion


    #region Getters

    // see DbItem

    #endregion

  }

  [System.Runtime.Serialization.DataContract(Name = "Item", Namespace = "TP.Object")]
  public class DbItem : TP.Object.Item
  {
    #region Cache

    private static AutoCache2Key<DbItem, int, string> cache = new AutoCache2Key<DbItem, int, string> {
        Constructor = _ => new DbItem(_),
        Identifier = _ => _.Field<int>("ItemID"),
        PrimaryKey = _ => _.ItemID,
        AlternateKey = _ => _.ItemNumber
      };

    private static CacheSimples cacheSimples = new CacheSimples();
    private class CacheSimples : CacheBase1Key<DbBaseItemList, string>
    {
      public ICollection<DbBaseItem> Fetch(string prefix, Func<DataTable> getter)
      {
        if (!base.contains(prefix))
        {
          DataTable dt = getter();
          base.put(prefix, new DbBaseItemList(from DataRow r in dt.Rows select new DbBaseItem(r)));
        }
        return base.retrieve(prefix).List;
      }

      public bool Contains(string prefix)
      {
        return base.contains(prefix);
      }

      public void Clear(string prefix)
      {
        base.clear(prefix);
      }

      public void ClearItemNumber(string itemNumber)
      {
        foreach (string k in base.keys())
        {
          if (itemNumber.StartsWith(k))
          {
            base.clear(k);
          }
        }
      }
    }
    private class DbBaseItemList : TP.Object.BaseBusinessObject
    {
      private List<DbBaseItem> _list;
      public DbBaseItemList(IEnumerable<DbBaseItem> incoming)
      {
        this._list = new List<DbBaseItem>(incoming);
      }
      public List<DbBaseItem> List { get { return this._list; } }
    }

    #endregion


    #region Constructors

    public DbItem(DataRow r) : base()
    {
      this.ItemID = (int)r["ItemID"];
      this.ItemNumber = (string)r["ItemNumber"];
      this.ItemDescription = (string)r["ItemDescription"];
      this.PrimaryVendor = (string)r["PrimaryVendor"];
      this.IsMasKit = (bool)r["IsMasKit"];
      this.IsDiscontinued = (bool)r["IsDiscontinued"];
      this.AdditionalInfo = this.nullableString(r["EPN"]);
      this.StandardPack = (int)r["StdPack"];
      this.PurchaseOrderComment = this.nullableString(r["POComment"]);
      this.LastReceiptDate = this.nullable<DateTime>(r["LastReceiptDate"]);
      this.Components = this.nullableString(r["Components"]);
      this.ItemIntroducedDate = (DateTime)r["PartIntroducedDate"];
      this.IsReconditioned = (bool)r["IsReconditioned"];
      this.IsTaxExempt = (bool)r["TaxExempt"];
      this.VendorQuantityOnHand = this.nullable<int>(r["VendorQuantityOnHand"]);
      this.VendorQuantityTimeChecked = this.nullable<DateTime>(r["VendorQuantityTimeChecked"]);
    }

    #endregion


    #region Getters

    // /item/{id:int}
    public static DbItem GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbItem GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      // TODO: table structure: add InventoryMaster.IsDiscontinued
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ItemID, ItemNumber, ItemDescription, PrimaryVendor, IsMasKit, CAST(0 as bit) AS IsDiscontinued, EPN, StdPack, POComment, LastReceiptDate, Components, PartIntroducedDate, IsReconditioned, TaxExempt, VendorQuantityOnHand, VendorQuantityTimeChecked FROM InventoryMaster WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /item/{item}
    public static DbItem GetByItemNumber(string itemNumber, string[] preloads = null) { return GetByItemNumber(null, itemNumber, preloads); }
    internal static DbItem GetByItemNumber(System.Data.SqlClient.SqlTransaction txn, string itemNumber, string[] preloads = null)
    {
      itemNumber = itemNumber.ToUpper();
      // TODO: table structure: add InventoryMaster.IsDiscontinued
      var ret = cache.FetchByAlternateKey(itemNumber, () => Connection.Retrieve(txn, "SELECT ItemID, ItemNumber, ItemDescription, PrimaryVendor, IsMasKit, CAST(0 as bit) AS IsDiscontinued, EPN, StdPack, POComment, LastReceiptDate, Components, PartIntroducedDate, IsReconditioned, TaxExempt, VendorQuantityOnHand, VendorQuantityTimeChecked FROM InventoryMaster WHERE ItemNumber=@item", new QueryParameter("@item", SqlDbType.VarChar, itemNumber)));
      return fillPreloads(ret, preloads, txn);
    }

    // /item/list/{prefix}
    public static IEnumerable<DbBaseItem> ListByPrefix(string prefix) { return ListByPrefix(null, prefix); }
    internal static IEnumerable<DbBaseItem> ListByPrefix(System.Data.SqlClient.SqlTransaction txn, string prefix)
    {
      prefix = prefix.ToUpper();
      var ret = cacheSimples.Fetch(prefix, () => Connection.Retrieve(txn, "SELECT ItemID, ItemNumber, CAST(0 as bit) AS IsDiscontinued FROM InventoryMaster WHERE ItemNumber LIKE @p", new QueryParameter("@p", SqlDbType.VarChar, prefix + "%")));
      return ret;
    }

    public static IEnumerable<DbItem> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbItem> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT ItemID, ItemNumber, ItemDescription, PrimaryVendor, IsMasKit, CAST(0 as bit) AS IsDiscontinued, EPN, StdPack, POComment, LastReceiptDate, Components, PartIntroducedDate, IsReconditioned, TaxExempt, VendorQuantityOnHand, VendorQuantityTimeChecked FROM InventoryMaster WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    public static IEnumerable<DbItem> GetExportRequiredPriceLevels() { return GetExportRequiredPriceLevels(null); }
    internal static IEnumerable<DbItem> GetExportRequiredPriceLevels(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT ItemID, ItemNumber, ItemDescription, PrimaryVendor, IsMasKit, CAST(0 as bit) AS IsDiscontinued, EPN, StdPack, POComment, LastReceiptDate, Components, PartIntroducedDate, IsReconditioned, TaxExempt, VendorQuantityOnHand, VendorQuantityTimeChecked FROM InventoryMaster WHERE ExportRequiredPriceLevels=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbItem fillPreloads(DbItem item, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (item == null)
      {
        return item;
      }
      var retval = (DbItem)item.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("productline"))
        {
          retval.FillProductLine(txn);
        }
        if (preloads.Contains("components"))
        {
          retval.FillItemComponents(preloads, txn);
        }
        if (preloads.Contains("costs"))
        {
          retval.FillItemCosts(txn);
        }
        if (preloads.Contains("prices"))
        {
          retval.FillItemPrices(txn);
        }
        if (preloads.Contains("quantities"))
        {
          retval.FillItemQuantities(txn);
        }
        if (preloads.Contains("freight"))
        {
          retval.FillItemFreightActuals(txn);
          retval.FillItemFreightEstimates(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbItem> fillPreloads(IEnumerable<DbItem> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillProductLine() { return this.FillProductLine(null); }
    internal bool FillProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._productline == null)
      {
        this._productline = DbProductLine.GetByPrefix(txn, this.ProductLinePrefix());
      }
      return true;
    }

    public override bool FillItemComponents(string[] preloads = null) { return FillItemComponents(preloads, null); }
    internal bool FillItemComponents(string[] preloads = null, System.Data.SqlClient.SqlTransaction txn = null)
    {
      if (this._itemcomponents == null) // send in preloads!
      {
        this._itemcomponents = DbItemComponent.GetByItemID(txn, this.ItemID, preloads);
      }
      return true;
    }

    public override bool FillItemCosts() { return FillItemCosts(null); }
    internal bool FillItemCosts(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemcosts == null)
      {
        this._itemcosts = DbItemCost.GetCurrentByItemID(txn, this.ItemID);
      }
      return true;
    }

    public override bool FillItemPrices() { return FillItemPrices(null); }
    internal bool FillItemPrices(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemprices == null)
      {
        this._itemprices = DbItemPrice.GetCurrentByItemID(txn, this.ItemID);
      }
      return true;
    }

    public override bool FillItemQuantities() { return FillItemQuantities(null); }
    internal bool FillItemQuantities(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemquantities == null)
      {
        this._itemquantities = DbItemQuantity.GetByItemID(txn, this.ItemID);
      }
      return true;
    }

    public override bool FillItemFreightEstimates() { return FillItemFreightEstimates(null); }
    internal bool FillItemFreightEstimates(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemfreightestimates == null)
      {
        this._itemfreightestimates = DbItemFreightEstimate.GetByItemID(txn, this.ItemID);
      }
      return true;
    }

    public override bool FillItemFreightActuals() { return FillItemFreightActuals(null); }
    internal bool FillItemFreightActuals(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemfreightactuals == null)
      {
        this._itemfreightactuals = DbItemFreightActual.GetByItemID(null, this.ItemID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /item/create
    public static DbItem CreateItem(string itemNumber, string itemDescription = "no description")
    {
      if (itemNumber != null)
      {
        itemNumber = itemNumber.ToUpper();
      }
      if (itemDescription == null) { itemDescription = "no description"; }

      // validate item number
      if (!IsItemNumberValid(itemNumber))
      {
        throw new ArgumentException("Invalid item number");
      }

      var txn = Connection.BeginTransaction();
      try {
        AssertItemDoesNotExist(itemNumber, txn);
        DbProductLine.AssertProductLineExists(ProductLinePrefix(itemNumber), txn);

        DbProductLine pl = DbProductLine.GetByPrefix(txn, ProductLinePrefix(itemNumber));
        
        int iid = Connection.Insert(txn,
            "INSERT INTO InventoryMaster (ItemNumber, ItemDescription, ProductLine, PrimaryVendor) VALUES (@item, @desc, @pl, @vendor)",
            new QueryParameter("@item", SqlDbType.VarChar, itemNumber),
            new QueryParameter("@desc", SqlDbType.VarChar, itemDescription),
            new QueryParameter("@pl", SqlDbType.VarChar, ProductLinePrefix(itemNumber)),
            new QueryParameter("@vendor", SqlDbType.VarChar, pl.DefaultVendorNo)
          );
        DbItemComponent.CreateDefaultFor(iid, itemNumber, txn);
        DbItemQuantity.CreateDefaultFor(iid, itemNumber, txn);
        // default site listings (from product line?)
        Connection.CommitTransaction(txn);

        /*cacheIDToObj.Remove(iid);
        cacheNameToObj.Remove(itemNumber);
        foreach (var k in cachePrefixToSimpleObjList.Keys) {
          if (itemNumber.StartsWith(k))
          {
            cachePrefixToSimpleObjList.Remove(k);
          }
        }*/
        cache.Clear(iid, itemNumber);
        cacheSimples.ClearItemNumber(itemNumber);
        return GetByID(iid);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /item/update/{id:int}
    public DbItem UpdateItem()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbItem orig = GetByID(txn, this.ItemID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE InventoryMaster SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        /*cacheIDToObj.Remove(orig.ItemID);
        cacheNameToObj.Remove(orig.ItemNumber);*/
        // don't need to invalid orig.*, since ItemNumber is immutable
        cache.Clear(this.ItemID, this.ItemNumber);
        return GetByID(this.ItemID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed

    public override bool FlagExported(Func<bool> recomparer, params string[] options)
    {
      // options can contain:
      //   "NewItemPriceLevels"

      throw new NotImplementedException();
    }

    #endregion


    #region Extra Validation

    internal static void AssertIsValidItemNumber(string itemNumber)
    {
      if (!IsItemNumberValid(itemNumber))
      {
        throw new ArgumentException("item number is not valid");
      }
      return;
    }

    internal static void AssertItemExists(int itemID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbItem item = GetByID(txn, itemID);
      if (item == null)
      {
        throw new ArgumentException("item does not exist");
      }
      return;
    }
    internal static void AssertItemExists(string itemNumber, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbItem item = GetByItemNumber(txn, itemNumber);
      if (item == null)
      {
        throw new ArgumentException("item does not exist");
      }
      return;
    }

    internal static void AssertItemDoesNotExist(string itemNumber, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbItem item = GetByItemNumber(txn, itemNumber);
      if (item != null)
      {
        throw new ArgumentException("item does not exist");
      }
      return;
    }

    #endregion

  }
}
