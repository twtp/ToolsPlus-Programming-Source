using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Person", Namespace = "TP.Object")]
  public class DbPerson : TP.Object.Person
  {
    #region Cache

    private static AutoCache2Key1Index<DbPerson, int, string, int> cache = new AutoCache2Key1Index<DbPerson, int, string, int> {
        Constructor = _ => new DbPerson(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.PersonID,
        AlternateKey = _ => _.NTUsername,
        Index1 = _ => 0,
        SecondsToLive = 86400
    };

    #endregion


    #region Constructors

    public DbPerson(DataRow r) : base()
    {
      this.PersonID = (int)r["ID"];
      this.NTUsername = (string)r["NTUsername"];
      this.ShortName = (string)r["ShortName"];
      this.IsDeleted = (bool)r["Deleted"];
      this.IsSpecialUser = (bool)r["SpecialUser"];
    }

    #endregion


    #region Getters

    // /person/{id:int}
    public static DbPerson GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbPerson GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, NTUsername, ShortName, Deleted, SpecialUser FROM Users WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /person/{ntuser}
    public static DbPerson GetByNTUsername(string username, string[] preloads = null) { return GetByNTUsername(null, username, preloads); }
    internal static DbPerson GetByNTUsername(System.Data.SqlClient.SqlTransaction txn, string username, string[] preloads = null)
    {
      username = username.ToLower();
      var ret = cache.FetchByAlternateKey(username, () => Connection.Retrieve(txn, "SELECT ID, NTUsername, ShortName, Deleted, SpecialUser FROM Users WHERE NTUsername=@n", new QueryParameter("@n", SqlDbType.VarChar, username)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbPerson FindOrCreatePerson(string username, string[] preloads = null) { return FindOrCreatePerson(null, username, preloads); }
    internal static DbPerson FindOrCreatePerson(System.Data.SqlClient.SqlTransaction txn, string username, string[] preloads = null)
    {
      if (string.IsNullOrEmpty(username)) { return null; }

      username = username.ToLower();
      var all = GetAll(txn, preloads);
      var found = all.SingleOrDefault(_ => _.NTUsername == username);
      if (found != null)
      {
        return found;
      }
      
      int id = CreatePerson(txn, username, username);
      cache.Clear(id, username, 0);
      return GetByID(txn, id);
    }

    public static IEnumerable<DbPerson> GetAll(string[] preloads = null) { return GetAll(null, preloads); }
    internal static IEnumerable<DbPerson> GetAll(System.Data.SqlClient.SqlTransaction txn, string[] preloads = null)
    {
      var ret = cache.FetchIndex1(0, () => Connection.Retrieve(txn, "SELECT ID, NTUsername, ShortName, Deleted, SpecialUser FROM Users"));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbPerson fillPreloads(DbPerson person, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (person == null)
      {
        return person;
      }
      var retval = (DbPerson)person.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here
      }
      return retval;
    }

    private static IEnumerable<DbPerson> fillPreloads(IEnumerable<DbPerson> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // TODO: DbPerson has-many ModulePermission

    #endregion


    #region Create/Update/Delete

    // /person/create
    public static DbPerson CreatePerson(string ntUsername, string shortName)
    {
      ntUsername = ntUsername.ToLower();
      AssertIsValidNTUsername(ntUsername);

      var txn = Connection.BeginTransaction();
      try
      {
        int id = CreatePerson(txn, ntUsername, shortName);
        Connection.CommitTransaction(txn);

        cache.Clear(id, ntUsername, 0);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }
    internal static int CreatePerson(System.Data.SqlClient.SqlTransaction txn, string ntUsername, string shortName)
    {
      ntUsername = ntUsername.ToLower();
      AssertIsValidNTUsername(ntUsername);

      DbPerson check = GetByNTUsername(txn, ntUsername);
      if (check != null)
      {
        throw new ArgumentException("nt username already exists!");
      }

      int id = Connection.Insert(txn,
          "INSERT INTO Users (NTUsername, ShortName) VALUES (@u, @n)",
          new QueryParameter("@u", SqlDbType.VarChar, ntUsername),
          new QueryParameter("@n", SqlDbType.VarChar, shortName)
        );

      return id;
    }

    // TODO: DbPerson.Create(string ntUsername, string shortName, DbPerson copyFrom) -- copies all permissions from given user

    // /person/update/{id:int}
    public DbPerson UpdatePerson()
    {
      this.NTUsername = this.NTUsername.ToLower();
      AssertIsValidNTUsername(this.NTUsername);

      var txn = Connection.BeginTransaction();
      try
      {
        DbPerson orig = GetByID(txn, this.PersonID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE Users SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(orig.PersonID, orig.NTUsername, 0); // username is mutable, so clear both
        cache.Clear(this.PersonID, this.NTUsername, 0);
        return GetByID(this.PersonID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed

    #endregion


    #region Extra Validation

    internal static void AssertPersonExists(int personID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbPerson p = GetByID(txn, personID);
      if (p == null)
      {
        throw new ArgumentException("person does not exist");
      }
      return;
    }

    internal static void AssertIsValidNTUsername(string ntUsername)
    {
      if (!IsValidNTUsername(ntUsername))
      {
        throw new ArgumentException("nt username is not valid!");
      }
      return;
    }

    #endregion

  }
}
