using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "PurchaseTerms", Namespace = "TP.Object")]
  public class DbPurchaseTerms : TP.Object.PurchaseTerms
  {
    #region Cache

    private static AutoCache1Key<DbPurchaseTerms, string> cache = new AutoCache1Key<DbPurchaseTerms, string> {
        Constructor = _ => new DbPurchaseTerms(_),
        Identifier = _ => _.Field<string>("TermsCode"),
        PrimaryKey = _ => _.Code,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbPurchaseTerms() { }

    public DbPurchaseTerms(DataRow r)
    {
      this.Code = (string)r["TermsCode"];
      this.Description = (string)r["TermsCodeDesc"];
      this.DaysBeforeDue = (int)r["DaysBeforeDue"];
      this.DueDateADayOfTheMonth = (string)r["DueDateADayOfTheMonth"];
      this.MinimumDaysAllowedInv = (int)r["MinimumDaysAllowedInv"];
      this.DaysDiscountAllowed = (int)r["DaysDiscountAllowed"];
      this.DiscountDateADayOfTheMonth = (string)r["DiscountDateADayOfTheMo"];
      this.MinimumDaysAllowedDisc = (int)r["MinimumDaysAllowedDisc"];
      this.DiscountRate = (decimal)r["DiscountRate"];
      this.ExportRequired = (bool)r["ExportRequired"];
    }

    #endregion


    #region Getters

    public static DbPurchaseTerms GetByID(string id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbPurchaseTerms GetByID(System.Data.SqlClient.SqlTransaction txn, string id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT TermsCode, TermsCodeDesc, DaysBeforeDue, DueDateADayOfTheMonth, MinimumDaysAllowedInv, DaysDiscountAllowed, DiscountDateADayOfTheMo, MinimumDaysAllowedDisc, DiscountRate, ExportRequired FROM AP_TermsCode WHERE TermsCode=@tc", new QueryParameter("@tc", SqlDbType.VarChar, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbPurchaseTerms> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbPurchaseTerms> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT TermsCode, TermsCodeDesc, DaysBeforeDue, DueDateADayOfTheMonth, MinimumDaysAllowedInv, DaysDiscountAllowed, DiscountDateADayOfTheMo, MinimumDaysAllowedDisc, DiscountRate, ExportRequired FROM AP_TermsCode WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbPurchaseTerms fillPreloads(DbPurchaseTerms terms, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (terms == null)
      {
        return terms;
      }
      var retval = (DbPurchaseTerms)terms.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbPurchaseTerms> fillPreloads(ICollection<DbPurchaseTerms> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    public static DbPurchaseTerms CreateNewPurchaseTerms(string code, string desc, int daysBeforeDue, int minAllowedDays, int discAllowedDays, int minDiscAllowedDays, decimal rate)
    {
      code = code.ToUpper();

      var txn = Connection.BeginTransaction();
      try
      {
        // check item exists
        var item = GetByID(code);
        if (item != null)
        {
          throw new ArgumentException("Terms with this ID already exist");
        }

        Connection.Execute(txn,
            "INSERT INTO AP_TermsCode (TermsCode, TermsCodeDesc, DaysBeforeDue, DueDateADayOfTheMonth, MinimumDaysAllowedInv, DaysDiscountAllowed, DiscountDateADayOfTheMo, MinimumDaysAllowedDisc, DiscountRate) VALUES (@tc, @d, @dbd, 'N', @mdai, @dda, 'N', @mdad, @r)",
            new QueryParameter("@tc", SqlDbType.VarChar, code),
            new QueryParameter("@d", SqlDbType.VarChar, desc),
            new QueryParameter("@dbd", SqlDbType.Int, daysBeforeDue),
            new QueryParameter("@mdai", SqlDbType.Int, minAllowedDays),
            new QueryParameter("@dda", SqlDbType.Int, discAllowedDays),
            new QueryParameter("@mdad", SqlDbType.Int, minDiscAllowedDays),
            new QueryParameter("@r", SqlDbType.Decimal, rate)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(code);
        return GetByID(code);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // update not allowed: if updates allowed, need to set trigger or changed.AddField("ExportRequired", SqlDbType.Bit, true);
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
              "UPDATE AP_TermsCode SET ExportRequired=0 WHERE TermsCode=@tc",
              new QueryParameter("@tc", SqlDbType.VarChar, this.Code)
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        else
        {
          // modified, do nothing
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.Code);
        return okayToUnflag;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion

  }
}
