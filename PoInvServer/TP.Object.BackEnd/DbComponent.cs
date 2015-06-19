using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Component", Namespace = "TP.Object")]
  public class DbComponent : TP.Object.Component
  {
    #region Cache

    private static AutoCache1Key<DbComponent, int> cache = new AutoCache1Key<DbComponent, int> {
        Constructor = _ => new DbComponent(_),
        Identifier = _ => _.Field<int>("ID"),
        PrimaryKey = _ => _.ComponentID
      };

    #endregion


    #region Constructors

    public DbComponent(DataRow r) : base()
    {
      this.ComponentID = (int)r["ID"];
      this.ComponentName = (string)r["Component"];
      this.Length = (decimal)r["Length"];
      this.Width = (decimal)r["Width"];
      this.Height =(decimal)r["Height"];
      this.Weight = (decimal)r["Weight"];
      this.PackingType = (int)r["PackingType"];
      this.IsNotConcealableDuringShipping = (bool)r["GiftNotConcealableWarning"];
      this.IsHazardous = (bool)r["HazMat"];
      this.IsTrackedBySerialNo = (bool)r["TrackSerialNo"];
      this.IsSpecificSideUp = (bool)r["SpecificUpSide"];
    }

    #endregion


    #region Getters

    // /component/{id:int}
    public static DbComponent GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbComponent GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, Component, CAST(Length as decimal(9,4)) AS Length, CAST(Width as decimal(9,4)) AS Width, CAST(Height as decimal(9,4)) AS Height, CAST(Weight as decimal(9,4)) AS Weight, PackingType, GiftNotConcealableWarning, HazMat, TrackSerialNo, SpecificUpSide FROM InventoryComponents WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /component/{name}
    public static DbComponent GetByName(string name, string[] preloads = null) { return GetByName(null, name, preloads); }
    internal static DbComponent GetByName(System.Data.SqlClient.SqlTransaction txn, string name, string[] preloads = null)
    {
      // name is not guaranteed to be unique, so not sure how this is going to work!
      var list = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT ID, Component, CAST(Length as decimal(9,4)) AS Length, CAST(Width as decimal(9,4)) AS Width, CAST(Height as decimal(9,4)) AS Height, CAST(Weight as decimal(9,4)) AS Weight, PackingType, GiftNotConcealableWarning, HazMat, TrackSerialNo, SpecificUpSide FROM InventoryComponents WHERE Component=@name", new QueryParameter("@name", SqlDbType.VarChar, name)));
      var ret = list.FirstOrDefault();
      return fillPreloads(ret, preloads, txn);
    }

    private static DbComponent fillPreloads(DbComponent c, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (c == null)
      {
        return c;
      }
      var retval = (DbComponent)c.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("barcodes"))
        {
          retval.FillBarcodeList(txn);
        }
        if (preloads.Contains("locations"))
        {
          retval.FillLocations(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbComponent> fillPreloads(ICollection<DbComponent> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillBarcodeList() { return FillBarcodeList(null); }
    internal bool FillBarcodeList(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._barcodes == null)
      {
        this._barcodes = DbBarcode.GetByComponentID(txn, this.ComponentID);
      }
      return true;
    }

    public override bool FillLocations() { return FillLocations(null); }
    internal bool FillLocations(System.Data.SqlClient.SqlTransaction txn, bool fillLocationObjects = false)
    {
      if (this._locations == null)
      {
        var l = DbInventoryLocationContent.GetByComponentID(txn, this.ComponentID);
        if (fillLocationObjects) {
          foreach (var iter in l)
          {
            iter.FillInventoryLocation(txn);
          }
        }
        this._locations = l;
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // internal only, called from item creation
    internal static int CreateDefaultFor(string itemNumber, System.Data.SqlClient.SqlTransaction txn)
    {
      int cid = Connection.Insert(txn,
          "INSERT INTO InventoryComponents (Component) VALUES (@item)",
          new QueryParameter("@item", SqlDbType.VarChar, itemNumber)
        );
      cache.Clear(cid);
      return cid;
    }

    // /component/create
    public static DbComponent CreateComponent(string componentName, int attachedToItemID)
    {
      DbItem.AssertItemExists(attachedToItemID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbItem item = DbItem.GetByID(txn, attachedToItemID);

        int cid = CreateDefaultFor(componentName, txn);
        DbItemComponent.CreateItemComponentLink(cid, attachedToItemID, item.ItemNumber, 1, txn);
        Connection.CommitTransaction(txn);

        cache.Clear(cid);
        return GetByID(cid);
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /component/update/{id:int}
    public DbComponent UpdateComponent()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbComponent orig = GetByID(txn, this.ComponentID);

        this.CanonicalizeDimensions();

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE InventoryComponents SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.ComponentID);
        return GetByID(this.ComponentID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // internal only, called via ItemComponent deletion
    internal bool DeleteComponent(System.Data.SqlClient.SqlTransaction txn)
    {
      DbComponent orig = GetByID(txn, this.ComponentID);

      Extensions.ObjectChanges changed = this.PropertiesModified(orig);
      if (changed.Count > 0)
      {
        throw new ArgumentException("component has been modified, not deleting");
      }

      DbBarcode.GetByComponentID(txn, this.ComponentID).Select(_ => _.DeleteBarcode(txn));
      Connection.Execute(txn, "DELETE FROM InventoryComponents WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, this.ComponentID));

      cache.Clear(this.ComponentID);
      return true;
    }

    #endregion


    #region Extra Validation

    internal static void AssertComponentExists(int componentID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbComponent c = DbComponent.GetByID(txn, componentID);
      if (c == null)
      {
        throw new ArgumentException("component id " + componentID + " not in database");
      }
      return;
    }

    #endregion

  }
}
