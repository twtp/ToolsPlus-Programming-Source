using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "PurchaseOrderTracking", Namespace = "TP.Object")]
  public class DbPurchaseOrderTracking : TP.Object.PurchaseOrderTracking
  {
    #region Cache

    private static AutoCache1Key1Index<DbPurchaseOrderTracking, int, int> cache = new AutoCache1Key1Index<DbPurchaseOrderTracking, int, int> {
        Constructor = _ => new DbPurchaseOrderTracking(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.HeaderID
      };

    #endregion


    #region Constructors

    public DbPurchaseOrderTracking(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.HeaderID = (int)r["HeaderID"];
      this.TrackingID = (string)r["TrackingID"];
      this.Carrier = (string)r["Carrier"];
      this.IsCustomerNotified = (bool)r["CustomerNotified"];
    }

    #endregion


    #region Getters

    public static DbPurchaseOrderTracking GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbPurchaseOrderTracking GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, TrackingID, Carrier, CustomerNotified FROM PurchaseOrderTracking WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbPurchaseOrderTracking> GetByPurchaseOrderID(int id, string[] preloads = null) { return GetByPurchaseOrderID(null, id, preloads); }
    internal static IEnumerable<DbPurchaseOrderTracking> GetByPurchaseOrderID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, HeaderID, TrackingID, Carrier, CustomerNotified FROM PurchaseOrderTracking WHERE HeaderID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbPurchaseOrderTracking fillPreloads(DbPurchaseOrderTracking ln, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (ln == null)
      {
        return ln;
      }
      var retval = (DbPurchaseOrderTracking)ln.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("purchaseorder"))
        {
          retval.FillPurchaseOrder(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbPurchaseOrderTracking> fillPreloads(IEnumerable<DbPurchaseOrderTracking> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbPurchaseOrderTracking has-one DbPurchaseOrder
    public override bool FillPurchaseOrder() { return FillPurchaseOrder(null); }
    internal bool FillPurchaseOrder(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._po == null)
      {
        this._po = DbPurchaseOrder.GetByID(txn, this.HeaderID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    public static DbPurchaseOrderTracking AddPackageToPurchaseOrder(int poID, string trackingID, string carrier, bool customerNotified)
    {
      if (trackingID == null) { throw new ArgumentException("tracking number required!"); }
      if (carrier == null) { throw new ArgumentException("carrier name required!"); }

      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrder.AssertPurchaseOrderExists(poID, txn);

        int id = Connection.Insert(txn,
            "INSERT INTO PurchaseOrderTracking (HeaderID, TrackingID, Carrier, CustomerNotified) VALUES (@id, @t, @c, @cn)",
            new QueryParameter("@id", SqlDbType.Int, poID),
            new QueryParameter("@t", SqlDbType.VarChar, trackingID),
            new QueryParameter("@c", SqlDbType.VarChar, carrier),
            new QueryParameter("@cn", SqlDbType.Bit, customerNotified)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, poID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public DbPurchaseOrderTracking UpdatePurchaseOrderTracking()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbPurchaseOrderTracking orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE PurchaseOrderTracking SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.HeaderID);
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion


    #region Extra Validation

    // any?

    #endregion

  }
}
