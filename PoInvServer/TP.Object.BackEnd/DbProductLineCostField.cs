using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineCostField", Namespace = "TP.Object")]
  public class DbProductLineCostField : TP.Object.ProductLineCostField
  {
    #region Cache

    private static AutoCache1Key1Index<DbProductLineCostField, int, int> cache = new AutoCache1Key1Index<DbProductLineCostField, int, int> {
        Constructor = _ => new DbProductLineCostField(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ProductLineID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbProductLineCostField() { }

    public DbProductLineCostField(DataRow r)
    {
      this.TupleID = (int)r["ID"];
      this.ProductLineID = (int)r["ProductLineID"];
      this.Name = (string)r["Name"];
      this.Description = (string)r["Description"];
      this.IsDeleted = (bool)r["IsDeleted"];
    }

    #endregion


    #region Getters

    // not externally visible
    public static DbProductLineCostField GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLineCostField GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, Name, Description, IsDeleted FROM ProductLineCostField WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{id:int}/costfields
    public static IEnumerable<DbProductLineCostField> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineCostField> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, Name, Description, IsDeleted FROM ProductLineCostField WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbProductLineCostField> GetActiveByProductLineID(int id, string[] preloads = null) { return GetActiveByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineCostField> GetActiveByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, Name, Description, IsDeleted FROM ProductLineCostField WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret.Where(_ => _.IsDeleted == false).ToArray(), preloads, txn);
    }

    private static DbProductLineCostField fillPreloads(DbProductLineCostField cf, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (cf == null)
      {
        return cf;
      }
      var retval = (DbProductLineCostField)cf.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("productlines"))
        {
          retval.FillProductLine(txn);
        }
        /*if (preloads.Contains("itemcosts"))
        {
          retval.FillCurrentItemCosts(txn);
        }*/
      }
      return retval;
    }

    private static IEnumerable<DbProductLineCostField> fillPreloads(IEnumerable<DbProductLineCostField> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbProductLineCostField has-one DbProductLine
    public override bool FillProductLine() { return FillProductLine(null); }
    internal bool FillProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._productline == null)
      {
        this._productline = DbProductLine.GetByID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLineCostField has-many ItemCost
    /*public override bool FillCurrentItemCosts() { return FillCurrentItemCosts(null); }
    internal bool FillCurrentItemCosts(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._itemcostcurrent == null)
      {
        this._itemcostcurrent = DbItemCost.GetBy(txn, this.SalesRepID);
      }
      return true;
    }*/

    #endregion


    #region Create/Update/Delete

    // /productline/{id:int}/costfields/create
    public static DbProductLineCostField CreateNewCostField(int productLineID, string name, string desc)
    {
      DbProductLine.AssertProductLineExists(productLineID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineCostField check = GetByProductLineID(txn, productLineID).SingleOrDefault(_ => _.Name == name);
        if (check != null)
        {
          throw new ArgumentException("cost field with this name already exists!");
        }

        int id = Connection.Insert(txn,
            "INSERT INTO ProductLineCostField (ProductLineID, Name, Description) VALUES (@pl, @name, @desc)",
            new QueryParameter("@pl", SqlDbType.Int, productLineID),
            new QueryParameter("@name", SqlDbType.VarChar, name),
            new QueryParameter("@desc", SqlDbType.VarChar, desc)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, productLineID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/{id:int}/costfields/update/{name}
    public DbProductLineCostField UpdateCostField()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineCostField orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineCostField SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID);
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed, update IsDeleted instead

    #endregion


    #region Extra Validation

    // none

    #endregion

  }
}
