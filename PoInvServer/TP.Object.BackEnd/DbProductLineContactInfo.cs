using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineContactInfo", Namespace = "TP.Object")]
  public class DbProductLineContactInfo : TP.Object.ProductLineContactInfo
  {
    #region Cache

    private static AutoCache1Key1Index<DbProductLineContactInfo, string, int> cache = new AutoCache1Key1Index<DbProductLineContactInfo, string, int>
    {
      Constructor = _ => new DbProductLineContactInfo(_),
      Identifier = _ => string.Format("{0}/{1}", _.Field<int>("ProductLineID"), _.Field<int>("InfoTypeID")), // matches .PrimaryKeyString
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ProductLineID,
      SecondsToLive = 86400
    };

    #endregion


    #region Constructors

    protected DbProductLineContactInfo() { }

    public DbProductLineContactInfo(DataRow r)
    {
      this.ProductLineID = (int)r["ProductLineID"];
      this.InfoType = (TP.Object.ProductLineContactInfoType)r["InfoTypeID"];
      this.TelephoneNo = this.nullableString(r["Phone"]);
      this.FaxNo = this.nullableString(r["Fax"]);
      this.HomepageURL = this.nullableString(r["URL"]);
      this.EmailAddress = this.nullableString(r["Email"]);
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}", this.ProductLineID, (int)this.InfoType); } }


    #region Getters

    // /productline/{id:int}/contact/
    public static IEnumerable<DbProductLineContactInfo> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineContactInfo> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ProductLineID, InfoTypeID, Phone, Fax, URL, Email FROM ProductLineContactInfo WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineContactInfo fillPreloads(DbProductLineContactInfo contactInfo, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (contactInfo == null)
      {
        return contactInfo;
      }
      var retval = (DbProductLineContactInfo)contactInfo.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbProductLineContactInfo> fillPreloads(IEnumerable<DbProductLineContactInfo> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // /productline/{id:int}/contact/create/{typeid:int}
    public static DbProductLineContactInfo CreateContactInfo(int productLineID, TP.Object.ProductLineContactInfoType infoType, string telephoneNo = null, string faxNo = null, string url = null, string email = null)
    {
      DbProductLine.AssertProductLineExists(productLineID);
      AssertIsValidTelephoneNo(telephoneNo);
      AssertIsValidFaxNo(faxNo);
      AssertIsValidUrl(url);
      AssertIsValidEmailAddress(email);

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineContactInfo check = GetByProductLineID(txn, productLineID).SingleOrDefault(_ => _.InfoType == infoType);
        if (check != null)
        {
          throw new ArgumentException("contact info tuple already exists!");
        }

        Connection.Execute(txn,
            "INSERT INTO ProductLineContactInfo (ProductLineID, InfoTypeID, Phone, Fax, URL, Email) VALUES (@pl, @t, @p, @f, @u, @e)",
            new QueryParameter("@pl", SqlDbType.Int, productLineID),
            new QueryParameter("@t", SqlDbType.Int, (int)infoType),
            new QueryParameter("@p", SqlDbType.VarChar, telephoneNo),
            new QueryParameter("@f", SqlDbType.VarChar, faxNo),
            new QueryParameter("@u", SqlDbType.VarChar, url),
            new QueryParameter("@e", SqlDbType.VarChar, email)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(string.Format("{0}/{1}", productLineID, (int)infoType), productLineID);
        return GetByProductLineID(productLineID).SingleOrDefault(_ => _.InfoType == infoType);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/{id:int}/contact/update
    public DbProductLineContactInfo UpdateContactInfo()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineContactInfo orig = GetByProductLineID(txn, this.ProductLineID).SingleOrDefault(_ => _.InfoType == this.InfoType);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineContactInfo SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.PrimaryKeyString, this.ProductLineID);
        return GetByProductLineID(this.ProductLineID).SingleOrDefault(_ => _.InfoType == this.InfoType);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public bool DeleteContactInfo()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        Connection.Execute(txn,
            "DELETE FROM ProductLineContactInfo WHERE ProductLineID=@pl AND InfoTypeID=@t",
            new QueryParameter("@pl", SqlDbType.Int, this.ProductLineID),
            new QueryParameter("@t", SqlDbType.Int, (int)this.InfoType)
          );
        Connection.CommitTransaction(txn);
        cache.Clear(this.PrimaryKeyString, this.ProductLineID);
        return true;
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


    #region Extra Validation

    internal static void AssertIsValidTelephoneNo(string telephoneNo)
    {
      if (!IsValidTelephoneNo(telephoneNo))
      {
        throw new ArgumentException("telephone number is not valid");
      }
    }

    internal static void AssertIsValidFaxNo(string faxNo)
    {
      if (!IsValidFaxNo(faxNo))
      {
        throw new ArgumentException("fax number is not valid");
      }
    }

    internal static void AssertIsValidUrl(string url)
    {
      if (!IsValidUrl(url))
      {
        throw new ArgumentException("url is not valid");
      }
    }

    internal static void AssertIsValidEmailAddress(string emailAddress)
    {
      if (!IsValidEmailAddress(emailAddress))
      {
        throw new ArgumentException("telephone number is not valid");
      }
    }

    #endregion


  }
}
