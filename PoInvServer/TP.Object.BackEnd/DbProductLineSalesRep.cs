using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineSalesRep", Namespace = "TP.Object")]
  public class DbProductLineSalesRep : TP.Object.ProductLineSalesRep
  {
    #region Cache

    private static AutoCache1Key2Index<DbProductLineSalesRep, int, int, int> cache = new AutoCache1Key2Index<DbProductLineSalesRep, int, int, int>() {
        Constructor = _ => new DbProductLineSalesRep(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ProductLineID,
        Index2 = _ => _.SalesRepID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbProductLineSalesRep() { }

    public DbProductLineSalesRep(DataRow r)
    {
      this.TupleID = (int)r["ID"];
      this.ProductLineID = (int)r["ProductLineID"];
      this.SalesRepID = (int)r["SalesRepID"];
      this.IsPrimary = (bool)r["IsPrimary"];
    }

    #endregion


    #region Getters

    // not externally visible
    public static DbProductLineSalesRep GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLineSalesRep GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, SalesRepID, IsPrimary FROM ProductLineSalesReps WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{id:int}/salesreps
    public static IEnumerable<DbProductLineSalesRep> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineSalesRep> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, SalesRepID, IsPrimary FROM ProductLineSalesReps WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /salesrep/{id:int}/lines
    public static IEnumerable<DbProductLineSalesRep> GetBySalesRepID(int id, string[] preloads = null) { return GetBySalesRepID(null, id, preloads); }
    internal static IEnumerable<DbProductLineSalesRep> GetBySalesRepID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex2(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, SalesRepID, IsPrimary FROM ProductLineSalesReps WHERE SalesRepID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineSalesRep fillPreloads(DbProductLineSalesRep mm, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (mm == null)
      {
        return mm;
      }
      var retval = (DbProductLineSalesRep)mm.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("productlines"))
        {
          retval.FillProductLine(txn);
        }
        if (preloads.Contains("reps"))
        {
          retval.FillSalesRep(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbProductLineSalesRep> fillPreloads(IEnumerable<DbProductLineSalesRep> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbProductLineSalesRep has-one DbProductLine
    public override bool FillProductLine() { return FillProductLine(null); }
    internal bool FillProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._productline == null)
      {
        this._productline = DbProductLine.GetByID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLineSalesRep has-one DbSalesRep
    public override bool FillSalesRep() { return FillSalesRep(null); }
    internal bool FillSalesRep(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._salesrep == null)
      {
        this._salesrep = DbSalesRep.GetByID(txn, this.SalesRepID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /productline/{id:int}/salesreps/add/{srid:int}
    // /salesreps/{id:int}/lines/add/{plid:int}
    public static DbProductLineSalesRep CreateLinkage(int productLineID, int salesRepID, bool isPrimary)
    {
      DbProductLine.AssertProductLineExists(productLineID);
      DbSalesRep.AssertSalesRepExists(salesRepID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineSalesRep check = GetByProductLineID(txn, productLineID).FirstOrDefault(_ => _.SalesRepID == salesRepID);
        if (check != null)
        {
          throw new ArgumentException("line/rep linkage already exists!");
        }

        int id = Connection.Insert(txn,
            "INSERT INTO ProductLineSalesReps (ProductLineID, SalesRepID, IsPrimary) VALUES (@pl, @sr, @pr)",
            new QueryParameter("@pl", SqlDbType.Int, productLineID),
            new QueryParameter("@sr", SqlDbType.Int, salesRepID),
            new QueryParameter("@pr", SqlDbType.Bit, isPrimary)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, productLineID, salesRepID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/{id:int}/salesreps/update/{srid:int}
    // /salesreps/{id:int}/lines/update/{plid:int}
    public DbProductLineSalesRep UpdateLinkage()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineSalesRep orig = GetByID(txn, this.TupleID);

        // TODO: if primary is updated to true, it should be un-set for others
        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineSalesReps SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID, this.SalesRepID);
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/{id:int}/salesreps/delete/{srid:int}
    // /salesreps/{id:int}/lines/delete/{plid:int}
    public bool DeleteLinkage()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineSalesRep orig = GetByID(txn, this.TupleID);
        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          throw new ArgumentException("line/rep link has been modified, not deleting");
        }

        Connection.Execute(txn,
            "DELETE FROM ProductLineSalesReps WHERE ID=@id",
            new QueryParameter("@id", SqlDbType.Int, this.TupleID)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID, this.SalesRepID);
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

    // none

    #endregion

  }
}
