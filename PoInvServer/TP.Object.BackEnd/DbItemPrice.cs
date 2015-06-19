using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemPrice", Namespace = "TP.Object")]
  public class DbItemPrice : TP.Object.ItemPrice
  {
    #region Cache

    private static AutoCache1Key1Index<DbItemPrice, int, int> cache = new AutoCache1Key1Index<DbItemPrice, int, int> {
        Constructor = _ => new DbItemPrice(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ItemID
      };

    #endregion


    #region Constructors

    public DbItemPrice(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.ItemID = (int)r["ItemID"];
      this.SalesChannelID = (int)r["SalesChannelID"];
      this.TimeEffective = (DateTime)r["TimeEffective"];
      this.MaximumQuantity1 = (int)r["MaximumQuantity1"];
      this.PriceValue1 = (decimal)r["PriceValue1"];
      this.MaximumQuantity2 = (int)r["MaximumQuantity2"];
      this.PriceValue2 = (decimal)r["PriceValue2"];
      this.MaximumQuantity3 = (int)r["MaximumQuantity3"];
      this.PriceValue3 = (decimal)r["PriceValue3"];
      this.MaximumQuantity4 = (int)r["MaximumQuantity4"];
      this.PriceValue4 = (decimal)r["PriceValue4"];
      this.MaximumQuantity5 = (int)r["MaximumQuantity5"];
      this.PriceValue5 = (decimal)r["PriceValue5"];
      this.PersonID = (int)r["PersonID"];
      this.ApplicationModuleID = (TP.Object.ApplicationModule)r["ApplicationModuleID"];
    }

    #endregion


    #region Getters

    // not externally available
    public static DbItemPrice GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    public static DbItemPrice GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ItemID, SalesChannelID, TimeEffective, " + buildSQLFields() + "PersonID, ApplicationModuleID, ExportRequired FROM ItemPrice WHERE ID=@id ORDER BY TimeEffective DESC", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /item/{id:int}/price
    public static IEnumerable<DbItemPrice> GetCurrentByItemID(int id, string[] preloads = null) { return GetCurrentByItemID(null, id, preloads); }
    internal static IEnumerable<DbItemPrice> GetCurrentByItemID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ItemID, SalesChannelID, TimeEffective, " + buildSQLFields() + "PersonID, ApplicationModuleID, ExportRequired FROM vItemPriceCurrent WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /item/{id:int}/price/channel/{chanid:int}
    public static DbItemPrice GetChannelCurrentByItemID(int chanID, int id, string[] preloads = null) { return GetChannelCurrentByItemID(null, chanID, id, preloads); }
    internal static DbItemPrice GetChannelCurrentByItemID(System.Data.SqlClient.SqlTransaction txn, int chanID, int id, string[] preloads = null)
    {
      return GetCurrentByItemID(txn, id, preloads).SingleOrDefault(_ => _.SalesChannelID == chanID);
    }

    // /item/{id:int}/price/history
    public static IEnumerable<DbItemPrice> GetHistoryByItemID(int id, DateTime dateLimit, int recordLimit = -1)
    {
      // TODO: DbItemPrice.GetHistoryByItemID is not cached?
      var dt = Connection.Retrieve(
          "SELECT " + (recordLimit > 0 ? ("TOP " + recordLimit.ToString() + " ") : "") + "ID, ItemID, SalesChannelID, TimeEffective, " + buildSQLFields() + "PersonID, ApplicationModuleID, ExportRequired FROM ItemPrice WHERE ItemID=@id AND TimeEffective>=@date ORDER BY TimeEffective DESC",
          new QueryParameter("@id", SqlDbType.Int, id),
          new QueryParameter("@date", SqlDbType.DateTime, dateLimit)
        );
      var retval = new List<DbItemPrice>(dt.Rows.Count);
      retval.AddRange(from DataRow r in dt.Rows select new DbItemPrice(r));
      return retval.AsReadOnly();
    }

    // /item/{id:int}/price/channel/{chanid:int}/history
    public static IEnumerable<DbItemPrice> GetChannelHistoryByItemID(int chanID, int id, DateTime dateLimit, int recordLimit = -1)
    {
      // TODO: DbItemPrice.GetHistoryByItemID is not cached?
      var dt = Connection.Retrieve(
          "SELECT " + (recordLimit > 0 ? ("TOP " + recordLimit.ToString() + " ") : "") + "ID, ItemID, SalesChannelID, TimeEffective, " + buildSQLFields() + "PersonID, ApplicationModuleID, ExportRequired FROM ItemPrice WHERE ItemID=@id AND SalesChannelID=@chan AND TimeEffective>=@date ORDER BY TimeEffective DESC",
          new QueryParameter("@id", SqlDbType.Int, id),
          new QueryParameter("@chan", SqlDbType.Int, chanID),
          new QueryParameter("@date", SqlDbType.DateTime, dateLimit)
        );
      var retval = new List<DbItemPrice>(dt.Rows.Count);
      retval.AddRange(from DataRow r in dt.Rows select new DbItemPrice(r));
      return retval.AsReadOnly();
    }

    public static IEnumerable<DbItemPrice> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbItemPrice> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT ID, ItemID, SalesChannelID, TimeEffective, " + buildSQLFields() + "PersonID, ApplicationModuleID, ExportRequired FROM vItemPriceCurrent WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbItemPrice fillPreloads(DbItemPrice price, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (price == null)
      {
        return price;
      }
      var retval = (DbItemPrice)price.MemberwiseClone();
      if (preloads != null)
      {
        // preloads here
      }
      return retval;
    }

    private static IEnumerable<DbItemPrice> fillPreloads(IEnumerable<DbItemPrice> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }
    
    #endregion


    #region Links

    // TODO: FillSalesChannel

    public override bool FillItem()
    {
      this._item = DbItem.GetByID(this.ItemID);
      return true;
    }

    public override bool FillPerson()
    {
      this._person = DbPerson.GetByID(this.PersonID);
      return true;
    }

    #endregion


    #region Create/Update/Delete

    public static DbItemPrice CreateItemPrice(int itemID, int salesChannelID, int q1, decimal p1, int q2, decimal p2, int q3, decimal p3, int q4, decimal p4, int q5, decimal p5, int personID, TP.Object.ApplicationModule applicationModule)
    {
      var priceList = MakeList(q1, p1, q2, p2, q3, p3, q4, p4, q5, p5);
      AssertPricesAreValid(priceList);

      var txn = Connection.BeginTransaction();
      try
      {
        DbItem.AssertItemExists(itemID, txn);
        //TODO: DbSalesChannel.AssertSalesChannelExists(salesChannelID, txn);
        DbPerson.AssertPersonExists(personID, txn);

        var parameters = new List<QueryParameter>(4 + MAXIMUM_LEVELS);
        parameters.Add(new QueryParameter("@item", SqlDbType.Int, itemID));
        parameters.Add(new QueryParameter("@chan", SqlDbType.Int, salesChannelID));
        parameters.AddRange(buildSQLParameters(priceList));
        parameters.Add(new QueryParameter("@person", SqlDbType.Int, personID));
        parameters.Add(new QueryParameter("@app", SqlDbType.Int, (int)applicationModule));
        int id = Connection.Insert(txn,
            "INSERT INTO ItemPrice(ItemID, SalesChannelID, " + buildSQLFields() + "PersonID, ApplicationModuleID) VALUES (@item, @chan, " + buildSQLPlaceholders() + "@person, @app)",
            parameters.ToArray()
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, itemID);
        return GetByID(id);
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    private static string _cachedSQLFields = null;
    private static string buildSQLFields()
    {
      if (_cachedSQLFields == null)
      {
        var sb = new System.Text.StringBuilder();
        for (int i = 1; i <= MAXIMUM_LEVELS; i++) { sb.AppendFormat("MaximumQuantity{0}, PriceValue{0}, ", i); }
        _cachedSQLFields = sb.ToString();
      }
      return _cachedSQLFields;
    }
    private static string _cachedSQLPlaceholders = null;
    private static string buildSQLPlaceholders()
    {
      if (_cachedSQLPlaceholders == null)
      {
        var sb = new System.Text.StringBuilder();
        for (int i = 1; i <= MAXIMUM_LEVELS; i++) { sb.AppendFormat("@q{0}, @p{0}, ", i); }
        _cachedSQLPlaceholders = sb.ToString();
      }
      return _cachedSQLPlaceholders;
    }
    private static QueryParameter[] buildSQLParameters(List<Tuple<int, decimal>> priceList)
    {
      var retval = new List<QueryParameter>(MAXIMUM_LEVELS + 5);
      for (int i = 1; i <= MAXIMUM_LEVELS; i++)
      {
        retval.Add(new QueryParameter("@q" + i.ToString(), SqlDbType.Int, priceList.Count >= i ? priceList[i - 1].Item1 : 0));
        retval.Add(new QueryParameter("@p" + i.ToString(), SqlDbType.Decimal, priceList.Count >= i ? priceList[i - 1].Item2 : 0));
      }
      return retval.ToArray();
    }

    // no updating or deleting allowed

    public override bool FlagExported(Func<bool> recomparer, params string[] options)
    {
      // options is unused

      var txn = Connection.BeginTransaction();
      try
      {
        bool okayToUnflag = recomparer();
        if (okayToUnflag)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ItemPrice SET ExportRequired=0 WHERE ID=@id",
              new QueryParameter("@v", SqlDbType.Int, this.TupleID)
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        else
        {
          // modified
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ItemID);
        return okayToUnflag;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


    #region Extra Validation

    public static void AssertPricesAreValid(List<Tuple<int, decimal>> priceList)
    {
      if (IsValid(priceList) == false)
      {
        throw new ArgumentException("price/breaks are not valid");
      }
    }

    #endregion

  }
}
