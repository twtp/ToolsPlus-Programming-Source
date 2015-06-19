using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineWholesalePriceLevel", Namespace = "TP.Object")]
  public class DbProductLineWholesalePriceLevel : TP.Object.ProductLineWholesalePriceLevel
  {
    #region Cache

    private static AutoCache1Key1Index<DbProductLineWholesalePriceLevel, string, int> cache = new AutoCache1Key1Index<DbProductLineWholesalePriceLevel, string, int>
    {
      Constructor = _ => new DbProductLineWholesalePriceLevel(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<int>("ProductLineID"), _.Field<int>("WholesalePriceLevel")), // matches .PrimaryKeyString
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ProductLineID,
      SecondsToLive = 86400
    };

    #endregion


    #region Constructors

    protected DbProductLineWholesalePriceLevel() { }

    public DbProductLineWholesalePriceLevel(DataRow r)
    {
      this.ProductLineID = (int)r["ProductLineID"];
      this.Level = (int)r["WholesalePriceLevel"];
      this.Percentage = (int)r["Percentage"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.ProductLineID, this.Level); } }


    #region Getters

    // /productline/{id:int}/wholesale
    public static IEnumerable<DbProductLineWholesalePriceLevel> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineWholesalePriceLevel> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ProductLineID, WholesalePriceLevel, Percentage FROM ProductLineWholesalePricing WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineWholesalePriceLevel fillPreloads(DbProductLineWholesalePriceLevel wholesaleLevel, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (wholesaleLevel == null)
      {
        return wholesaleLevel;
      }
      var retval = (DbProductLineWholesalePriceLevel)wholesaleLevel.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbProductLineWholesalePriceLevel> fillPreloads(IEnumerable<DbProductLineWholesalePriceLevel> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // not externally accessible, only called via product line creation
    internal static IEnumerable<DbProductLineWholesalePriceLevel> CreateDefaultWholesalePriceLevels(int productLineID, System.Data.SqlClient.SqlTransaction txn) {
      // dang, this repeats the IntegerMinimum and IntegerMaximum on the level
      for (int i = 0; i <= 5; i++)
      {
        Connection.Execute(txn,
            "INSERT INTO ProductLineWholesalePricing (ProductLineID, WholesalePriceLevel, Percentage) VALUES (@id, @l, @p)",
            new QueryParameter("@id", SqlDbType.Int, productLineID),
            new QueryParameter("@l", SqlDbType.Int, i),
            new QueryParameter("@p", SqlDbType.Int, 20)
          );
        cache.Clear(string.Format("{0}/{1}", productLineID, i), productLineID);
      }
      return GetByProductLineID(txn, productLineID, null); // txn lookup, because this is ongoing
    }

    public DbProductLineWholesalePriceLevel UpdateWholesalePriceLevel()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineWholesalePriceLevel orig = GetByProductLineID(txn, this.ProductLineID).FirstOrDefault(_ => _.Level == this.Level);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineWholesalePricing SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.PrimaryKeyString, this.ProductLineID);
        return GetByProductLineID(this.ProductLineID).FirstOrDefault(_ => _.Level == this.Level);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed

    #endregion

  }
}
