using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineSubstitution", Namespace = "TP.Object")]
  public class DbProductLineSubstitution : TP.Object.ProductLineSubstitution
  {
    #region Cache

    private static AutoCache2Key1Index<DbProductLineSubstitution, int, int, int> cache = new AutoCache2Key1Index<DbProductLineSubstitution, int, int, int> {
        Constructor = _ => new DbProductLineSubstitution(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        AlternateKey = _ => _.ProductLineID,
        Index1 = _ => _.PurchasingProductLineID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbProductLineSubstitution() { }

    public DbProductLineSubstitution(DataRow r)
    {
      this.TupleID = (int)r["ID"];
      this.ProductLineID = (int)r["ProductLineID"];
      this.PurchasingProductLineID = (int)r["PurchasingProductLineID"];
    }

    #endregion


    #region Getters

    // not externally accessible
    public static DbProductLineSubstitution GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLineSubstitution GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, PurchasingProductLineID FROM ProductLineSubForPOs WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{id:int}/substitutes/to
    public static DbProductLineSubstitution GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static DbProductLineSubstitution GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByAlternateKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, PurchasingProductLineID FROM ProductLineSubForPOs WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{id:int}/substitutes/from
    public static IEnumerable<DbProductLineSubstitution> GetByPurchasingProductLineID(int id, string[] preloads = null) { return GetByPurchasingProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineSubstitution> GetByPurchasingProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, PurchasingProductLineID FROM ProductLineSubForPOs WHERE PurchasingProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineSubstitution fillPreloads(DbProductLineSubstitution sub, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (sub == null)
      {
        return sub;
      }
      var retval = (DbProductLineSubstitution)sub.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("productline"))
        {
          sub.FillProductLine(txn);
          sub.FillPurchasingProductLine(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbProductLineSubstitution> fillPreloads(IEnumerable<DbProductLineSubstitution> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillProductLine() { return FillProductLine(null); }
    internal bool FillProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._pl == null)
      {
        this._pl = DbProductLine.GetByID(txn, this.ProductLineID);
      }
      return true;
    }

    public override bool FillPurchasingProductLine() { return FillPurchasingProductLine(null); }
    internal bool FillPurchasingProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._ppl == null)
      {
        this._ppl = DbProductLine.GetByID(txn, this.PurchasingProductLineID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /productline/{id:int}/substitutes/to/add/{pplid:int}
    public static DbProductLineSubstitution CreateSubstitution(int productLineID, int purchasingProductLineID)
    {
      DbProductLine.AssertProductLineExists(productLineID);
      DbProductLine.AssertProductLineExists(purchasingProductLineID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineSubstitution check = DbProductLineSubstitution.GetByProductLineID(txn, productLineID);
        if (check != null)
        {
          throw new ArgumentException("substitution already exists!");
        }

        int id = Connection.Insert(txn,
            "INSERT INTO ProductLineSubForPOs (ProductLineID, PurchasingProductLineID) VALUES (@pl1, @pl2)",
            new QueryParameter("@pl1", SqlDbType.Int, productLineID),
            new QueryParameter("@pl2", SqlDbType.Int, purchasingProductLineID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, productLineID, purchasingProductLineID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/{id:int}/substitutes/to/update
    public DbProductLineSubstitution UpdateSubstitution()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineSubstitution orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineSubForPOs SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID, this.PurchasingProductLineID);
        cache.Clear(this.TupleID, orig.ProductLineID, orig.PurchasingProductLineID); // PPLID is mutable
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public bool DeleteSubstitution()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        Connection.Execute(txn,
            "DELETE FROM ProductLineSubForPOs WHERE ID=@id",
            new QueryParameter("@id", SqlDbType.Int, this.TupleID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID, this.PurchasingProductLineID);
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
