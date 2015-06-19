using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "SalesRep", Namespace = "TP.Object")]
  public class DbSalesRep : TP.Object.SalesRep
  {
    #region Cache

    private static AutoCache1Key<DbSalesRep, int> cache = new AutoCache1Key<DbSalesRep, int> {
        Constructor = _ => new DbSalesRep(_),
        Identifier = _ => _.Field<int>("ID"),
        PrimaryKey = _ => _.SalesRepID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbSalesRep() { }

    public DbSalesRep(DataRow r)
    {
      this.SalesRepID = (int)r["ID"];
      this.FirstName = this.nullableString(r["FirstName"]);
      this.LastName = this.nullableString(r["LastName"]);
      this.PositionTitle = this.nullableString(r["PositionTitle"]);
      this.Address = this.nullableString(r["Address"]);
      this.City = this.nullableString(r["City"]);
      this.State = this.nullableString(r["State"]);
      this.PostalCode = this.nullableString(r["ZipCode"]);
      this.TelephoneNo = this.nullableString(r["Phone"]);
      this.MobileTelephoneNo = this.nullableString(r["Cell"]);
      this.EmailAddress = this.nullableString(r["Email"]);
      this.Notes = this.nullableString(r["Notes"]);
    }

    #endregion


    #region Getters

    // /salesrep/{id:int}
    public static DbSalesRep GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbSalesRep GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, FirstName, LastName, PositionTitle, Address, City, State, ZipCode, Phone, Cell, Fax, Email, Notes FROM SalesReps WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbSalesRep fillPreloads(DbSalesRep rep, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (rep == null)
      {
        return rep;
      }
      var retval = (DbSalesRep)rep.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    #endregion


    #region Links

    // DbSalesRep has-many DbProductLineSalesRep
    public override bool FillProductLinesRepped() { return FillProductLinesRepped(null); }
    internal bool FillProductLinesRepped(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._productlines == null)
      {
        this._productlines = DbProductLineSalesRep.GetBySalesRepID(txn, this.SalesRepID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /salesrep/create
    public static DbSalesRep CreateSalesRep(string firstName, string lastName)
    {
      AssertIsValidFirstName(firstName);
      AssertIsValidLastName(lastName);

      var txn = Connection.BeginTransaction();
      try
      {
        // get by name to check for dupes?

        int id = Connection.Insert(txn,
            "INSERT INTO SalesReps (FirstName, LastName) VALUES (@f, @l)",
            new QueryParameter("@f", SqlDbType.VarChar, firstName),
            new QueryParameter("@l", SqlDbType.VarChar, lastName)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /salesrep/update/{id:int}
    public DbSalesRep UpdateSalesRep()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbSalesRep orig = GetByID(txn, this.SalesRepID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE SalesReps SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.SalesRepID);
        return GetByID(this.SalesRepID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed?

    #endregion


    #region Extra Validation

    public static void AssertSalesRepExists(int salesRepID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbSalesRep test = GetByID(txn, salesRepID);
      if (test == null)
      {
        throw new ArgumentException("sales rep does not exist");
      }
      return;
    }

    public static void AssertIsValidFirstName(string firstName)
    {
      if (!IsValidFirstName(firstName))
      {
        throw new ArgumentException("first name is not valid");
      }
      return;
    }

    public static void AssertIsValidLastName(string lastName)
    {
      if (!IsValidLastName(lastName))
      {
        throw new ArgumentException("last name is not valid");
      }
      return;
    }

    #endregion

  }
}
