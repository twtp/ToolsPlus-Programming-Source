using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Computer", Namespace = "TP.Object")]
  public class DbComputer : TP.Object.Computer
  {
    #region Cache

    private static AutoCache2Key1Index<DbComputer, int, string, int> cache = new AutoCache2Key1Index<DbComputer, int, string, int>
    {
      Constructor = _ => new DbComputer(_),
      Identifier = _ => (int)_["ComputerID"],
      PrimaryKey = _ => _.ComputerID,
      AlternateKey = _ => _.ComputerName,
      Index1 = _ => 0,
      SecondsToLive = 86400
    };

    #endregion


    #region Constructors

    public DbComputer(DataRow r) : base()
    {
      this.ComputerID = (int)r["ComputerID"];
      this.ComputerName = (string)r["ComputerName"];
    }

    #endregion


    #region Getters

    // /computer/{id:int}
    public static DbComputer GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbComputer GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ComputerID, ComputerName FROM Computers WHERE ComputerID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /computer/{ntuser}
    public static DbComputer GetByComputerName(string computerName, string[] preloads = null) { return GetByComputerName(null, computerName, preloads); }
    internal static DbComputer GetByComputerName(System.Data.SqlClient.SqlTransaction txn, string computerName, string[] preloads = null)
    {
      computerName = computerName.ToUpper();
      var ret = cache.FetchByAlternateKey(computerName, () => Connection.Retrieve(txn, "SELECT ComputerID, ComputerName FROM Computers WHERE ComputerName=@n", new QueryParameter("@n", SqlDbType.VarChar, computerName)));
      return fillPreloads(ret, preloads, txn);
    }

    public static DbComputer FindOrCreateComputer(string computerName, string[] preloads = null) { return FindOrCreateComputer(null, computerName, preloads); }
    internal static DbComputer FindOrCreateComputer(System.Data.SqlClient.SqlTransaction txn, string computerName, string[] preloads = null)
    {
      if (string.IsNullOrEmpty(computerName)) { return null; }

      computerName = computerName.ToUpper();
      var all = GetAll(txn, preloads);
      var found = all.SingleOrDefault(_ => _.ComputerName == computerName);
      if (found != null)
      {
        return found;
      }

      int id = CreateComputer(txn, computerName);
      cache.Clear(id, computerName, 0);
      return GetByID(txn, id);
    }

    public static IEnumerable<DbComputer> GetAll(string[] preloads = null) { return GetAll(null, preloads); }
    internal static IEnumerable<DbComputer> GetAll(System.Data.SqlClient.SqlTransaction txn, string[] preloads = null)
    {
      var ret = cache.FetchIndex1(0, () => Connection.Retrieve(txn, "SELECT ComputerID, ComputerName FROM Computers"));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbComputer fillPreloads(DbComputer comp, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (comp == null)
      {
        return comp;
      }
      var retval = (DbComputer)comp.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here
      }
      return retval;
    }

    private static IEnumerable<DbComputer> fillPreloads(IEnumerable<DbComputer> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // /computer/create
    public static DbComputer CreateComputer(string computerName)
    {
      computerName = computerName.ToUpper();

      var txn = Connection.BeginTransaction();
      try
      {
        int id = CreateComputer(txn, computerName);
        Connection.CommitTransaction(txn);

        cache.Clear(id, computerName, 0);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }
    internal static int CreateComputer(System.Data.SqlClient.SqlTransaction txn, string computerName)
    {
      computerName = computerName.ToUpper();

      DbComputer check = GetByComputerName(txn, computerName);
      if (check != null)
      {
        throw new ArgumentException("computer name already exists!");
      }

      int id = Connection.Insert(txn,
          "INSERT INTO Computers (ComputerName) VALUES (@n)",
          new QueryParameter("@n", SqlDbType.VarChar, computerName)
        );

      return id;
    }

    // update not allowed
    // delete not allowed

    #endregion


    #region Extra Validation

    // none

    #endregion

  }
}
