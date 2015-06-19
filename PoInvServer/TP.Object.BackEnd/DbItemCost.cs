using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemCost", Namespace = "TP.Object")]
  public class DbItemCost : TP.Object.ItemCost
  {
    #region Cache

    private static AutoCache1Key1Index<DbItemCost, int, int> cache = new AutoCache1Key1Index<DbItemCost, int, int> {
        Constructor = _ => new DbItemCost(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ItemID
      };

    #endregion


    #region Constructors

    public DbItemCost(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.ItemID = (int)r["ItemID"];
      this.CostType = (string)r["ItemCostType"];
      this.CostValue = (decimal)r["CostValue"];
      this.TimeEffective = (DateTime)r["TimeEffective"];
      this.PersonID = (int)r["PersonID"];
      this.ApplicationModuleID = (TP.Object.ApplicationModule)r["ApplicationModuleID"];
    }

    #endregion


    #region Getters

    public static DbItemCost GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbItemCost GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ItemID, ItemCostType, CostValue, TimeEffective, PersonID, ApplicationModuleID FROM ItemCost WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbItemCost> GetCurrentByItemID(int id, string[] preloads = null) { return GetCurrentByItemID(null, id, preloads); }
    internal static IEnumerable<DbItemCost> GetCurrentByItemID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ItemID, ItemCostType, CostValue, TimeEffective, PersonID, ApplicationModuleID FROM vItemCostCurrent WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbItemCost GetSpecifiedCurrentByItemID(string type, int id, string[] preloads = null) { return GetSpecifiedCurrentByItemID(null, type, id, preloads); }
    public static DbItemCost GetSpecifiedCurrentByItemID(System.Data.SqlClient.SqlTransaction txn, string type, int id, string[] preloads = null)
    {
      return GetCurrentByItemID(txn, id, preloads).SingleOrDefault(_ => _.CostType == type);
    }

    public static IEnumerable<DbItemCost> GetHistoryByItemID(int id, DateTime dateLimit, int recordLimit = -1)
    {
      // TODO: DbItemCost.GetHistoryByItemID is not cached?
      var dt = Connection.Retrieve("SELECT " + (recordLimit > 0 ? ("TOP " + recordLimit.ToString() + " ") : "") + " ID, ItemID, ItemCostType, CostValue, TimeEffective, PersonID, ApplicationModuleID FROM ItemCost WHERE ItemID=@id AND TimeEffective>=@date ORDER BY TimeEffective DESC", new QueryParameter("@id", SqlDbType.Int, id), new QueryParameter("@date", SqlDbType.DateTime, dateLimit));
      var retval = new List<DbItemCost>(dt.Rows.Count);
      retval.AddRange(from DataRow r in dt.Rows select new DbItemCost(r));
      return fillPreloads(retval, null, null);
    }

    public static IEnumerable<DbItemCost> GetSpecifiedHistoryByItemID(string type, int id, DateTime dateLimit, int recordLimit = -1)
    {
      // TODO: DbItemCost.GetSpecifiedHistoryByItemID is not cached?
      var dt = Connection.Retrieve(
        "SELECT " + (recordLimit > 0 ? ("TOP " + recordLimit.ToString() + " ") : "") + "ID, ItemID, ItemCostType, CostValue, TimeEffective, PersonID, ApplicationModuleID FROM ItemCost WHERE ItemID=@id AND ItemCostType=@type AND TimeEffective>=@date ORDER BY TimeEffective DESC",
          new QueryParameter("@id", SqlDbType.Int, id),
          new QueryParameter("@type", SqlDbType.VarChar, type),
          new QueryParameter("@date", SqlDbType.DateTime, dateLimit)
        );
      var retval = new List<DbItemCost>(dt.Rows.Count);
      retval.AddRange(from DataRow r in dt.Rows select new DbItemCost(r));
      return fillPreloads(retval, null, null);
    }

    private static DbItemCost fillPreloads(DbItemCost cost, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (cost == null)
      {
        return cost;
      }
      var retval = (DbItemCost)cost.MemberwiseClone();
      if (preloads != null)
      {
        // preloads here
      }
      return retval;
    }

    private static IEnumerable<DbItemCost> fillPreloads(IEnumerable<DbItemCost> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbItemCost has-one DbPerson
    public override bool FillPerson() { return FillPerson(null); }
    internal bool FillPerson(System.Data.SqlClient.SqlTransaction txn)
    {
      this._person = DbPerson.GetByID(txn, this.PersonID);
      return true;
    }

    #endregion


    #region Create/Update/Delete

    public static DbItemCost CreateItemCost(int itemID, string costType, decimal costValue, int personID, TP.Object.ApplicationModule applicationModule)
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbItem.AssertItemExists(itemID, txn);
        DbPerson.AssertPersonExists(personID, txn);

        int id = Connection.Insert(txn,
            "INSERT INTO ItemCost (ItemID, ItemCostType, CostValue, PersonID, ApplicationModuleID) VALUES (@i, @t, @c, @p, @a)",
            new QueryParameter("@i", SqlDbType.Int, itemID),
            new QueryParameter("@t", SqlDbType.VarChar, costType),
            new QueryParameter("@c", SqlDbType.Decimal, costValue),
            new QueryParameter("@p", SqlDbType.Int, personID),
            new QueryParameter("@a", SqlDbType.Int, (int)applicationModule)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, itemID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // no updating or deleting allowed

    #endregion

  }
}
