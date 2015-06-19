using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemFreightEstimate", Namespace = "TP.Object")]
  public class DbItemFreightEstimate : TP.Object.ItemFreightEstimate
  {
    #region Cache

    private static AutoCache1Key1Index<DbItemFreightEstimate, string, int> cache = new AutoCache1Key1Index<DbItemFreightEstimate, string, int>
    {
      Constructor = _ => new DbItemFreightEstimate(_),
      Identifier = _ => string.Format("{0}/{1}/{2}", _.Field<int>("ItemID"), _.Field<string>("ZipCode"), _.Field<string>("Service")), // matches .PrimaryKeyString and cache.Clear in .Create
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ItemID
    };

    #endregion


    #region Constructors

    public DbItemFreightEstimate(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.ItemID = (int)r["ItemID"];
      this.PostalCode = (string)r["ZipCode"];
      this.Cost = (decimal)r["Cost"];
      this.Service = (string)r["Service"];
      this.LastUpdated = (DateTime)r["LastUpdated"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}/{2}", this.ItemID, this.PostalCode, this.Service); } }


    #region Getters

    public static IEnumerable<DbItemFreightEstimate> GetByItemID(int itemID, string[] preloads = null) { return GetByItemID(null, itemID, preloads); }
    internal static IEnumerable<DbItemFreightEstimate> GetByItemID(System.Data.SqlClient.SqlTransaction txn, int itemID, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(itemID, () => Connection.Retrieve(txn, "SELECT ID, ItemID, ZipCode, Cost, Service, LastUpdated FROM FreightEstimates WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, itemID)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbItemFreightEstimate fillPreloads(DbItemFreightEstimate freight, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (freight == null)
      {
        return freight;
      }
      var retval = (DbItemFreightEstimate)freight.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbItemFreightEstimate> fillPreloads(IEnumerable<DbItemFreightEstimate> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // internal create, used only by estimate function
    internal static bool CreateItemFreightLine(System.Data.SqlClient.SqlTransaction txn, int itemID, string postalCode, decimal cost, string serviceIdentifier)
    {
      DbItem.AssertItemExists(itemID, txn);

      Connection.Execute(txn,
          "DELETE FROM FreightEstimates WHERE ItemID=@id AND ZipCode=@pc",
          new QueryParameter("id", SqlDbType.Int, itemID),
          new QueryParameter("@pc", SqlDbType.VarChar, postalCode)
        );
      int id = Connection.Insert(txn,
          "INSERT INTO FreightEstimates (ItemID, ZipCode, Cost, Service) VALUES (@id, @pc, @c, @s)",
          new QueryParameter("@id", SqlDbType.Int, itemID),
          new QueryParameter("@pc", SqlDbType.VarChar, postalCode),
          new QueryParameter("@c", SqlDbType.Decimal, cost),
          new QueryParameter("@s", SqlDbType.VarChar, serviceIdentifier)
        );

      cache.Clear(string.Format("{0}/{1}/{2}", itemID, postalCode, serviceIdentifier), itemID);
      return true;
    }

    // no updating, rolled into creation

    // /item/{id:int}/freight/estimate/delete/{fid:int}
    public bool DeleteFreightEstimate()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        Connection.Execute(txn,
          "DELETE FROM FreightEstimates WHERE ID=@id",
          new QueryParameter("id", SqlDbType.Int, this.TupleID)
        );
        Connection.CommitTransaction(txn);

        cache.Clear(this.PrimaryKeyString, this.ItemID);
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
