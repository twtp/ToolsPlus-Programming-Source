using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ProductLineDropshipMatrixLine", Namespace = "TP.Object")]
  public class DbProductLineDropshipMatrixLine : TP.Object.ProductLineDropshipMatrixLine
  {
    #region Cache

    private static AutoCache1Key1Index<DbProductLineDropshipMatrixLine, int, int> cache = new AutoCache1Key1Index<DbProductLineDropshipMatrixLine, int, int> {
        Constructor = _ => new DbProductLineDropshipMatrixLine(_),
        Identifier = _ => (int)_["ID"],
        PrimaryKey = _ => _.TupleID,
        Index1 = _ => _.ProductLineID,
        SecondsToLive = 86400
      };

    #endregion


    #region Constructors

    protected DbProductLineDropshipMatrixLine() { }

    public DbProductLineDropshipMatrixLine(DataRow r)
    {
      this.TupleID = (int)r["ID"];
      this.ProductLineID = (int)r["ProductLineID"];
      this.ThresholdCost = (decimal)r["ThresholdCost"];
      this.CalculationType = (TP.Object.DropshipChargeCalculationType)r["AmountCalculationType"];
      this.Amount = (decimal)r["Amount"];
    }

    #endregion


    #region Getters

    public static DbProductLineDropshipMatrixLine GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLineDropshipMatrixLine GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, ThresholdCost, AmountCalculationType, Amount FROM ProductLineDropshipFlatFees WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbProductLineDropshipMatrixLine> GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static IEnumerable<DbProductLineDropshipMatrixLine> GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLineID, ThresholdCost, AmountCalculationType, Amount FROM ProductLineDropshipFlatFees WHERE ProductLineID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbProductLineDropshipMatrixLine fillPreloads(DbProductLineDropshipMatrixLine charge, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (charge == null)
      {
        return charge;
      }
      var retval = (DbProductLineDropshipMatrixLine)charge.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbProductLineDropshipMatrixLine> fillPreloads(IEnumerable<DbProductLineDropshipMatrixLine> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // this should only be called from DbProductLineDropshipMatrix.UpdateMatrix()
    // very little sanity checking here, since most of it is related to how the
    // lines interact.
    internal static DbProductLineDropshipMatrixLine CreateLine(System.Data.SqlClient.SqlTransaction txn, int productLineID, decimal threshold, TP.Object.DropshipChargeCalculationType calcType, decimal amount)
    {
      DbProductLine.AssertProductLineExists(productLineID, txn);

      int id = Connection.Insert(txn,
          "INSERT INTO ProductLineDropshipFlatFees (ProductLineID, ThresholdCost, AmountCalculationType, Amount) VALUES (@pl, @th, @ct, @am)",
          new QueryParameter("@pl", SqlDbType.Int, productLineID),
          new QueryParameter("@th", SqlDbType.Decimal, threshold),
          new QueryParameter("@ct", SqlDbType.Int, (int)calcType),
          new QueryParameter("@am", SqlDbType.Decimal, amount)
        );

      cache.Clear(id, productLineID);
      return GetByID(txn, id); // txn lookup, because this is ongoing
    }

    // this should only be called from DbProductLineDropshipMatrix.UpdateMatrix()
    // very little sanity checking here, since most of it is related to how the
    // lines interact.
    internal static DbProductLineDropshipMatrixLine UpdateLine(System.Data.SqlClient.SqlTransaction txn, int id, TP.Object.DropshipChargeCalculationType calcType, decimal threshold, decimal amount)
    {
      DbProductLineDropshipMatrixLine orig = GetByID(txn, id);
      DbProductLineDropshipMatrixLine upd = GetByID(txn, id);
      upd.CalculationType = calcType;
      upd.ThresholdCost = threshold;
      upd.Amount = amount;

      Extensions.ObjectChanges changed = Extensions.PropertiesModified(upd, orig);
      if (changed.Count > 0)
      {
        int affected = Connection.Execute(txn,
            "UPDATE ProductLineDropshipFlatFees SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
            changed.UpdateValues
          );
        if (affected == 0)
        {
          throw new ApplicationException("error updating!");
        }
      }

      cache.Clear(orig.TupleID, orig.ProductLineID);
      return GetByID(txn, upd.TupleID); // txn lookup, because this is ongoing
    }

    internal bool DeleteLine(System.Data.SqlClient.SqlTransaction txn)
    {
      Connection.Execute(txn, "DELETE FROM ProductLineDropshipFlatFees WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, this.TupleID));
      cache.Clear(this.TupleID, this.ProductLineID);
      return true;
    }

    #endregion

  }














  [System.Runtime.Serialization.DataContract(Name = "ProductLineDropshipMatrix", Namespace = "TP.Object")]
  public class DbProductLineDropshipMatrix : TP.Object.ProductLineDropshipMatrix
  {
    #region Constructors

    protected DbProductLineDropshipMatrix() { }

    public DbProductLineDropshipMatrix(IEnumerable<DbProductLineDropshipMatrixLine> coll)
    {
      this._list = new List<Object.ProductLineDropshipMatrixLine>(coll);
    }

    #endregion


    #region Getters

    public static DbProductLineDropshipMatrix GetByProductLineID(int id, string[] preloads = null) { return GetByProductLineID(null, id, preloads); }
    internal static DbProductLineDropshipMatrix GetByProductLineID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      return new DbProductLineDropshipMatrix(DbProductLineDropshipMatrixLine.GetByProductLineID(txn, id));
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    public static DbProductLineDropshipMatrix OverwriteMatrix(int productLineID, TP.Object.DropshipChargeCalculationType calcType, List<Tuple<decimal,decimal>> values)
    {
     sanityCheck(values);

      var txn = Connection.BeginTransaction();
      try {
        DbProductLineDropshipMatrix orig = DbProductLineDropshipMatrix.GetByProductLineID(txn, productLineID);

        int min = Math.Min(values.Count, orig._list.Count);
        for (int i = 0; i < min; i++)
        {
          DbProductLineDropshipMatrixLine.UpdateLine(txn, orig._list[i].TupleID, calcType, values[i].Item1, values[i].Item2);
        }

        if (values.Count == orig._list.Count)
        {
          // nothing more to do
        }
        else if (values.Count > orig._list.Count)
        {
          // append new lines
          for (int i = min; i < values.Count; i++)
          {
            DbProductLineDropshipMatrixLine.CreateLine(txn, productLineID, values[i].Item1, calcType, values[i].Item2);
          }
        }
        else {
          // delete existing lines
          for (int i = min; i < orig._list.Count; i++)
          {
            ((DbProductLineDropshipMatrixLine)orig._list[i]).DeleteLine(txn);
          }
        }

        Connection.CommitTransaction(txn);

        return DbProductLineDropshipMatrix.GetByProductLineID(productLineID);
      }
      catch {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    #endregion
    

    private static void sanityCheck(List<Tuple<decimal,decimal>> values) {
      decimal threshold = 0;
      foreach (var iter in values) {
        if (threshold > iter.Item1)
        {
          throw new ArgumentException("thresholds must be incrementing!");
        }
        threshold = iter.Item1;
      }
      return;
    }
  }
}
