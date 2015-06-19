using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "MasWarehouse", Namespace = "TP.Object")]
  public class DbMasWarehouse : TP.Object.MasWarehouse
  {
    #region Cache

    private static AutoCache1Key<DbMasWarehouse, string> cache = new  AutoCache1Key<DbMasWarehouse, string> {
        Constructor = _ => new DbMasWarehouse(_),
        Identifier = _ => _.Field<string>("WarehouseCode"),
        PrimaryKey = _ => _.WarehouseCode,
        SecondsToLive = 86400
      };

    #endregion

    #region Constructors

    protected DbMasWarehouse() { }

    public DbMasWarehouse(DataRow r)
    {
      this.WarehouseCode = (string)r["WarehouseCode"];
      this.Description = this.nullableString(r["WarehouseDesc"]);
      this.Name = this.nullableString(r["WarehouseName"]);
      this.Address1 = this.nullableString(r["WarehouseAddress1"]);
      this.Address2 = this.nullableString(r["WarehouseAddress2"]);
      this.Address3 = this.nullableString(r["WarehouseAddress3"]);
      this.City = this.nullableString(r["WarehouseCity"]);
      this.State = this.nullableString(r["WarehouseState"]);
      this.PostalCode = this.nullableString(r["WarehouseZipCode"]);
      this.CountryCode = this.nullableString(r["WarehouseCountryCode"]);
      this.ExportRequired = (bool)r["ExportRequired"];
    }

    #endregion


    #region Getters

    public static DbMasWarehouse GetByCode(string code, string[] preloads = null) { return GetByCode(null, code, preloads); }
    internal static DbMasWarehouse GetByCode(System.Data.SqlClient.SqlTransaction txn, string code, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(code, () => Connection.Retrieve(txn, "SELECT WarehouseCode, WarehouseDesc, WarehouseName, WarehouseAddress1, WarehouseAddress2, WarehouseAddress3, WarehouseCity, WarehouseState, WarehouseZipCode, WarehouseCountryCode, ExportRequired FROM IM_Warehouse WHERE WarehouseCode=@c", new QueryParameter("@c", SqlDbType.VarChar, code)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbMasWarehouse> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbMasWarehouse> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT WarehouseCode, WarehouseDesc, WarehouseName, WarehouseAddress1, WarehouseAddress2, WarehouseAddress3, WarehouseCity, WarehouseState, WarehouseZipCode, WarehouseCountryCode, ExportRequired FROM IM_Warehouse WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbMasWarehouse fillPreloads(DbMasWarehouse whse, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (whse == null)
      {
        return whse;
      }
      var retval = (DbMasWarehouse)whse.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbMasWarehouse> fillPreloads(IEnumerable<DbMasWarehouse> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    public static DbMasWarehouse CreateNewMasWarehouse(string code, string desc, string name, string addr1, string addr2, string addr3, string city, string state, string postalCode, string countryCode)
    {
      if (code == null)
      {
        throw new ArgumentException("code can't be null!");
      }
      if (code.Length != 3)
      {
        throw new ArgumentException("code must be exactly 3 characters long");
      }
      code = code.ToUpper();

      var txn = Connection.BeginTransaction();
      try
      {
        // check item exists
        var item = GetByCode(code);
        if (item != null)
        {
          throw new ArgumentException("Warehouse with this ID already exist");
        }

        Connection.Execute(txn,
            "INSERT INTO IM_Warehouse (WarehouseCode, WarehouseDesc, WarehouseName, WarehouseAddress1, WarehouseAddress2, WarehouseAddress3, WarehouseCity, WarehouseState, WarehouseZipCode, WarehouseCountryCode, ExportRequired) VALUES (@code, @desc, @name, @addr1, @addr2, @addr3, @city, @state, @zip, @country, 1)",
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
        return GetByCode(code);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // update not allowed
    // delete not allowed

    #endregion

  }
}
