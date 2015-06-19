using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Vendor", Namespace = "TP.Object")]
  public class DbVendor : TP.Object.Vendor
  {
    #region Cache

    private static AutoCache1Key<DbVendor, string> cache = new AutoCache1Key<DbVendor, string> {
        Constructor = _ => new DbVendor(_),
        Identifier = _ => _.Field<string>("VendorNo"),
        PrimaryKey = _ => _.VendorNo,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbVendor() { }

    public DbVendor(DataRow r)
    {
      this.DivisionNo = (string)r["APDivisionNo"];
      this.VendorNo = (string)r["VendorNo"];
      this.Name = this.nullableString(r["VendorName"]);
      this.AddressLine1 = this.nullableString(r["AddressLine1"]);
      this.AddressLine2 = this.nullableString(r["AddressLine2"]);
      this.City = this.nullableString(r["City"]);
      this.State = this.nullableString(r["State"]);
      this.PostalCode = this.nullableString(r["ZipCode"]);
      this.CountryCode = this.nullableString(r["CountryCode"]);
      this.TelephoneNo = this.nullableString(r["TelephoneNo"]);
      this.DefaultPurchaseTermsCode = this.nullableString(r["TermsCode"]);
      this.Reference = this.nullableString(r["Reference"]);
      this.ExportRequired = (bool)r["ExportRequired"];
    }

    #endregion


    #region Getters

    // /vendor/{vendorno:length(3,7)}
    public static DbVendor GetByID(string id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbVendor GetByID(System.Data.SqlClient.SqlTransaction txn, string id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT APDivisionNo, VendorNo, VendorName, AddressLine1, AddressLine2, City, State, ZipCode, CountryCode, TelephoneNo, TermsCode, Reference, ExportRequired FROM AP_Vendor WHERE APDivisionNo='00' AND VendorNo=@id", new QueryParameter("@id", SqlDbType.VarChar, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbVendor> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbVendor> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT APDivisionNo, VendorNo, VendorName, AddressLine1, AddressLine2, City, State, ZipCode, CountryCode, TelephoneNo, TermsCode, Reference, ExportRequired FROM AP_Vendor WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    private static DbVendor fillPreloads(DbVendor vendor, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (vendor == null)
      {
        return vendor;
      }
      var retval = (DbVendor)vendor.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("terms"))
        {
          retval.FillPurchaseTerms(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbVendor> fillPreloads(ICollection<DbVendor> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbVendor has-one DbPurchaseTerms
    public override bool FillPurchaseTerms() { return FillPurchaseTerms(null); }
    internal bool FillPurchaseTerms(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this.DefaultPurchaseTermsCode == null)
      {
        return false;
      }
      if (this._purchaseterms == null)
      {
        this._purchaseterms = DbPurchaseTerms.GetByID(txn, this.DefaultPurchaseTermsCode);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /vendor/create
    public static DbVendor CreateVendor(string vendorNo, string vendorName)
    {
      vendorNo = vendorNo.ToUpper();
      AssertIsValidVendorNo(vendorNo);
      AssertIsValidVendorName(vendorName);

      var txn = Connection.BeginTransaction();
      try
      {
        DbVendor check = GetByID(txn, vendorNo);
        if (check != null)
        {
          throw new ArgumentException("vendor number already exists!");
        }

        Connection.Execute(txn,
            "INSERT INTO AP_Vendor (APDivisionNo, VendorNo, VendorName) VALUES ('00', @no, @name)",
            new QueryParameter("@no", SqlDbType.VarChar, vendorNo),
            new QueryParameter("@name", SqlDbType.VarChar, vendorName)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(vendorNo);
        return GetByID(vendorNo, null);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /vendor/update/{vendorno:length(3,7)
    public DbVendor UpdateVendor()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbVendor orig = GetByID(txn, this.VendorNo);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          changed.AddField("ExportRequired", SqlDbType.Bit, true); // should this be a trigger?
          int affected = Connection.Execute(txn,
              "UPDATE AP_Vendor SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.VendorNo);
        return GetByID(this.VendorNo);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

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
              "UPDATE AP_Vendor SET ExportRequired=0 WHERE VendorNo=@v",
              new QueryParameter("@v", SqlDbType.VarChar, this.VendorNo)
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

        cache.Clear(this.VendorNo);
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

    internal static void AssertIsValidVendorNo(string vendorNo)
    {
      if (!IsVendorNoValid(vendorNo))
      {
        throw new ArgumentException("vendor number is not valid");
      }
      return;
    }

    internal static void AssertIsValidVendorName(string vendorName)
    {
      if (!IsVendorNameValid(vendorName))
      {
        throw new ArgumentException("vendor name is not valid");
      }
      return;
    }

    #endregion

  }
}
