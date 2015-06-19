using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineDropshipCharge", Namespace = "TP.Object")]
  public class DbProductLineDropshipCharge : TP.Object.ProductLineDropshipCharge
  {
    #region Cache

    private static AutoCache1Key1Index<DbProductLineDropshipCharge, int, int> cache = new AutoCache1Key1Index<DbProductLineDropshipCharge, int, int> {
        Constructor = _ => new DbProductLineDropshipCharge(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ProductLineID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbProductLineDropshipCharge() { }

    public DbProductLineDropshipCharge(DataRow r)
    {
      this.TupleID = (int)r["ID"];
      this.ProductLineID = (int)r["ProductLineID"];
      this.AddressZoningType = (TP.Object.AddressZoningType)r["AddressType"];
      this.LiftGateCharge = this.nullable<decimal>(r["LiftGateCharge"]);
      this.LiftGateWeightThreshold = this.nullable<int>(r["LiftGateWeightThreshold"]);
      this.AdderAmount = this.nullable<decimal>(r["AdderAmount"]);
      this.AdderCalculationType = (TP.Object.DropshipChargeCalculationType)r["AdderType"];
      this.AdderThreshold = this.nullable<decimal>(r["AdderThreshold"]);
      this.AdderMaximum = this.nullable<decimal>(r["AdderMaximum"]);
    }

    #endregion


    #region Getters

    // not externally accessible
    public static DbProductLineDropshipCharge GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLineDropshipCharge GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, AddressType, LiftGateCharge, LiftGateWeightThreshold, AdderAmount, AdderType, AdderThreshold, AdderMaximum FROM ProductLineDropshipCharges WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{id:int}/dropshipcharges
    public static IEnumerable<DbProductLineDropshipCharge> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineDropshipCharge> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, AddressType, LiftGateCharge, LiftGateWeightThreshold, AdderAmount, AdderType, AdderThreshold, AdderMaximum FROM ProductLineDropshipCharges WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineDropshipCharge fillPreloads(DbProductLineDropshipCharge charge, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (charge == null)
      {
        return charge;
      }
      var retval = (DbProductLineDropshipCharge)charge.MemberwiseClone();
      if (preloads != null)
      {
        // charge matrix is ALWAYS filled
        retval.FillChargeMatrix(txn);
      }
      return retval;
    }

    private static IEnumerable<DbProductLineDropshipCharge> fillPreloads(IEnumerable<DbProductLineDropshipCharge> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    public override bool FillChargeMatrix() { return FillChargeMatrix(null); }
    internal bool FillChargeMatrix(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._matrix == null)
      {
        this._matrix = DbProductLineDropshipMatrix.GetByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    public static DbProductLineDropshipCharge CreateChargeLine(int productLineID, TP.Object.AddressZoningType zoning)
    {
      DbProductLine.AssertProductLineExists(productLineID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineDropshipCharge check = GetByProductLineID(txn, productLineID).SingleOrDefault(_ => _.AddressZoningType == zoning);
        if (check != null)
        {
          throw new ArgumentException("dropship charge tuple already exists!");
        }

        int id = Connection.Insert(txn,
            "INSERT INTO ProductLineDropshipCharges (ProductLineID, AddressType) VALUES (@id, @at)",
            new QueryParameter("@id", SqlDbType.Int, productLineID),
            new QueryParameter("@at", SqlDbType.Int, (int)zoning)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(id, productLineID);
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public DbProductLineDropshipCharge UpdateDropshipCharges()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLineDropshipCharge orig = GetByID(txn, this.TupleID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLineDropshipCharges SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.TupleID, this.ProductLineID);
        return GetByID(this.TupleID);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // delete not allowed

    #endregion


  }
}
