using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "PurchaseOrderShipToAddress", Namespace = "TP.Object")]
  public class DbPurchaseOrderShipToAddress : TP.Object.PurchaseOrderShipToAddress
  {
    #region Cache

    private static AutoCache1Key<DbPurchaseOrderShipToAddress, string> cache = new AutoCache1Key<DbPurchaseOrderShipToAddress, string> {
        Constructor = _ => new DbPurchaseOrderShipToAddress(_),
        Identifier = _ => _.Field<string>("ShipToCode"),
        PrimaryKey = _ => _.ShipToCode,
        SecondsToLive = 86400
      };

    #endregion

    #region Constructors

    public DbPurchaseOrderShipToAddress(DataRow r) : base()
    {
      this.ShipToCode = (string)r["ShipToCode"];
      this.ShipToCodeDesc = this.nullableString(r["ShipToCodeDesc"]);
      this.ShipToName = this.nullableString(r["ShipToName"]);
      this.ShipToAddress1 = this.nullableString(r["ShipToAddress1"]);
      this.ShipToAddress2 = this.nullableString(r["ShipToAddress2"]);
      this.ShipToAddress3 = this.nullableString(r["ShipToAddress3"]);
      this.ShipToCity = this.nullableString(r["ShipToCity"]);
      this.ShipToState = this.nullableString(r["ShipToState"]);
      this.ShipToPostalCode = this.nullableString(r["ShipToZipCode"]);
      this.ShipToCountryCode = this.nullableString(r["ShipToCountryCode"]);
      this.ExportRequired = (bool)r["ExportRequired"];
    }

    #endregion


    #region Getters

    public static DbPurchaseOrderShipToAddress GetByID(string code, string[] preloads = null) { return GetByID(null, code, preloads); }
    internal static DbPurchaseOrderShipToAddress GetByID(System.Data.SqlClient.SqlTransaction txn, string code, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(code, () => Connection.Retrieve(txn, "SELECT ShipToCode, ShipToCodeDesc, ShipToName, ShipToAddress1, ShipToAddress2, ShipToAddress3, ShipToCity, ShipToState, ShipToZipCode, ShipToCountryCode, ExportRequired FROM PO_ShipToAddress WHERE ShipToCode=@code", new QueryParameter("@code", SqlDbType.VarChar, code)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbPurchaseOrderShipToAddress fillPreloads(DbPurchaseOrderShipToAddress ln, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (ln == null)
      {
        return ln;
      }
      var retval = (DbPurchaseOrderShipToAddress)ln.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here
      }
      return retval;
    }

    private static IEnumerable<DbPurchaseOrderShipToAddress> fillPreloads(ICollection<DbPurchaseOrderShipToAddress> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    public static DbPurchaseOrderShipToAddress CreateShipToAddress(string code, string desc, string name, string addr1, string addr2, string addr3, string city, string state, string postalCode, string countryCode)
    {
      if (code == null || code.Length == 0)
      {
        throw new ArgumentException("invalid ship-to code");
      }
      // should length-check, etc all these fields. ehh.

      var txn = Connection.BeginTransaction();
      try
      {
        Connection.Execute(txn,
            "INSERT INTO PO_ShipToAddress (ShipToCode, ShipToCodeDesc, ShipToName, ShipToAddress1, ShipToAddress2, ShipToAddress3, ShipToCity, ShipToState, ShipToZipCode, ShipToCountryCode, ExportRequired) VALUES (@code, @desc, @name, @addr1, @addr2, @addr3, @city, @state, @zip, @country, 1)",
            new QueryParameter("@code", SqlDbType.VarChar, code),
            new QueryParameter("@desc", SqlDbType.VarChar, desc),
            new QueryParameter("@name", SqlDbType.VarChar, name),
            new QueryParameter("@addr1", SqlDbType.VarChar, addr1),
            new QueryParameter("@addr2", SqlDbType.VarChar, addr2),
            new QueryParameter("@addr3", SqlDbType.VarChar, addr3),
            new QueryParameter("@city", SqlDbType.VarChar, city),
            new QueryParameter("@state", SqlDbType.VarChar, state),
            new QueryParameter("@zip", SqlDbType.VarChar, postalCode),
            new QueryParameter("@country", SqlDbType.VarChar, countryCode)
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

    // no update
    // no delete

    #endregion


    #region Extra Validation

    // any?

    #endregion

  }
}
