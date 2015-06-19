using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemKitContent", Namespace = "TP.Object")]
  public class DbItemKitContent : TP.Object.ItemKitContent
  {
    #region Cache

    private static AutoCache1Key2Index<DbItemKitContent, string, string, string> cache = new AutoCache1Key2Index<DbItemKitContent, string, string, string>
    {
      Constructor = _ => new DbItemKitContent(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<string>("SalesKitNo"), _.Field<string>("ComponentItemCode")), // matches .PrimaryKeyString and cache.Clear in .Create
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ItemNumber,
      Index2 = _ => _.SubItemNumber
    };

    #endregion


    #region Constructors

    public DbItemKitContent(DataRow r) : base()
    {
      this.ItemNumber = r.Field<string>("SalesKitNo");
      this.SubItemNumber = r.Field<string>("ComponentItemCode");
      this.Quantity = r.Field<int>("QuantityPerAssembly");
      this.IsCoreComponent = r.Field<bool>("CoreComponent");
      this.ExportRequired = r.Field<bool>("ExportRequired");
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.ItemNumber, this.SubItemNumber); } }


    #region Getters

    public static DbItemKitContent GetByPrimaryKey(string itemNumber, string subItemNumber, string[] preloads = null) { return GetByPrimaryKey(null, itemNumber, subItemNumber, preloads); }
    internal static DbItemKitContent GetByPrimaryKey(System.Data.SqlClient.SqlTransaction txn, string itemNumber, string subItemNumber, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(string.Format("{0}/{1}", itemNumber, subItemNumber), () => Connection.Retrieve(txn, "SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent, ExportRequired FROM IM_SalesKitDetail WHERE SalesKitNo=@item AND ComponentItemCode=@subitem", new QueryParameter("@item", SqlDbType.VarChar, itemNumber), new QueryParameter("@subitem", SqlDbType.VarChar, subItemNumber)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbItemKitContent> GetByItemNumber(string itemNumber, string[] preloads = null) { return GetByItemNumber(null, itemNumber, preloads); }
    internal static IEnumerable<DbItemKitContent> GetByItemNumber(System.Data.SqlClient.SqlTransaction txn, string itemNumber, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(itemNumber, () => Connection.Retrieve(txn, "SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent, ExportRequired FROM IM_SalesKitDetail WHERE SalesKitNo=@item", new QueryParameter("@item", SqlDbType.VarChar, itemNumber)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbItemKitContent> GetByComponentItemNumber(string subItemNumber, string[] preloads = null) { return GetByComponentItemNumber(null, subItemNumber, preloads); }
    internal static IEnumerable<DbItemKitContent> GetByComponentItemNumber(System.Data.SqlClient.SqlTransaction txn, string subItemNumber, string[] preloads = null)
    {
      var ret = cache.FetchByIndex2(subItemNumber, () => Connection.Retrieve(txn, "SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent, ExportRequired FROM IM_SalesKitDetail WHERE ComponentItemCode=@item", new QueryParameter("@item", SqlDbType.VarChar, subItemNumber)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbItemKitContent> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbItemKitContent> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent, ExportRequired FROM IM_SalesKitDetail WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbItemKitContent fillPreloads(DbItemKitContent kc, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (kc == null)
      {
        return kc;
      }
      var retval = (DbItemKitContent)kc.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbItemKitContent> fillPreloads(IEnumerable<DbItemKitContent> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // internal create, only called from CreateItemAsKit
    internal static bool CreateKitDetailLine(System.Data.SqlClient.SqlTransaction txn, string itemNumber, string subItemNumber, int quantity, bool isCoreComponent)
    {
      DbItem.AssertItemExists(itemNumber, txn);
      DbItem.AssertItemExists(subItemNumber, txn);

      int id = Connection.Insert(txn,
          "INSERT INTO IM_SalesKitDetail (SalesKitNo, ComponentItemCode, QuantityPerAssembly, IsCoreComponent) VALUES (@k, @i, @q, @c)",
          new QueryParameter("@k", SqlDbType.VarChar, itemNumber),
          new QueryParameter("@i", SqlDbType.VarChar, subItemNumber),
          new QueryParameter("@q", SqlDbType.Int, quantity),
          new QueryParameter("@c", SqlDbType.Bit, isCoreComponent)
        );

      cache.Clear(string.Format("{0}/{1}", itemNumber, subItemNumber), itemNumber, subItemNumber);
      return true;
    }

    // update not allowed
    // delete not allowed

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
              "UPDATE IM_SalesKitDetail SET ExportRequired=0 WHERE SalesKitNo=@k AND ComponentItemCode=@i",
              new QueryParameter("@k", SqlDbType.VarChar, this.ItemNumber),
              new QueryParameter("@i", SqlDbType.VarChar, this.SubItemNumber)
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

        cache.Clear(this.PrimaryKeyString, this.ItemNumber, this.SubItemNumber);
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

    // none

    #endregion

  }
}
