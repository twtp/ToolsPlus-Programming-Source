using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "ItemFreightActual", Namespace = "TP.Object")]
  public class DbItemFreightActual : TP.Object.ItemFreightActual
  {
    #region Cache

    private static AutoCache1Key1Index<DbItemFreightActual, string, int> cache = new AutoCache1Key1Index<DbItemFreightActual, string, int>
    {
      Constructor = _ => new DbItemFreightActual(_),
      Identifier = _ => string.Format("{0}/{1}/{2}", _.Field<int>("ItemID"), _.Field<string>("SalesOrderNumber"), _.Field<DateTime>("DateOfShipment")), // matches .PrimaryKeyString and cache.Clear in .Create
      PrimaryKey = _ => _.PrimaryKeyString,
      Index1 = _ => _.ItemID
    };

    #endregion


    #region Constructors

    public DbItemFreightActual(DataRow r) : base()
    {
      this.TupleID = (int)r["ID"];
      this.ItemID = (int)r["ItemID"];
      this.PostalCode = (string)r["ZipCode"];
      this.Cost = (decimal)r["Cost"];
      this.ShipmentDate = (DateTime)r["DateOfShipment"];
      this.SaleDate = (DateTime)r["DateOfSale"];
      this.SalesOrderNo = (string)r["SalesOrderNumber"];
      this.ItemUnitPrice = (decimal)r["ItemUnitPrice"];
    }

    #endregion


    private string PrimaryKeyString { get { return string.Format("{0}/{1}/{2}", this.ItemID, this.SalesOrderNo, this.ShipmentDate); } }


    #region Getters

    public static IEnumerable<DbItemFreightActual> GetByItemID(int itemID, string[] preloads = null) { return GetByItemID(null, itemID, preloads); }
    internal static IEnumerable<DbItemFreightActual> GetByItemID(System.Data.SqlClient.SqlTransaction txn, int itemID, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(itemID, () => Connection.Retrieve(txn, "SELECT ID, ItemID, ZipCode, Cost, DateOfShipment, DateOfSale, SalesOrderNumber, ItemUnitPrice FROM FreightActual WHERE ItemID=@id", new QueryParameter("@id", SqlDbType.Int, itemID)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbItemFreightActual fillPreloads(DbItemFreightActual freight, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (freight == null)
      {
        return freight;
      }
      var retval = (DbItemFreightActual)freight.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbItemFreightActual> fillPreloads(IEnumerable<DbItemFreightActual> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // internal create, used only by nightly data import
    internal static bool CreateItemFreightLine(System.Data.SqlClient.SqlTransaction txn, int itemID, string postalCode, decimal cost, DateTime shipDate, DateTime saleDate, string salesOrderNo, decimal itemUnitPrice)
    {
      DbItem.AssertItemExists(itemID, txn);

      int id = Connection.Insert(txn,
          "INSERT INTO FreightActual (ItemID, ZipCode, Cost, DateOfShipment, DateOfSale, SalesOrderNumber, ItemUnitPrice) VALUES (@id, @pc, @c, @d1, @d2, @so, @p)",
          new QueryParameter("@id", SqlDbType.Int, itemID),
          new QueryParameter("@pc", SqlDbType.VarChar, postalCode),
          new QueryParameter("@c", SqlDbType.Decimal, cost),
          new QueryParameter("@d1", SqlDbType.DateTime, shipDate),
          new QueryParameter("@d2", SqlDbType.DateTime, saleDate),
          new QueryParameter("@so", SqlDbType.VarChar, salesOrderNo),
          new QueryParameter("@p", SqlDbType.Decimal, itemUnitPrice)
        );

      cache.Clear(string.Format("{0}/{1}/{2}", itemID, salesOrderNo, shipDate), itemID);
      return true;
    }

    // update not allowed
    // delete not allowed

    #endregion


    #region Extra Validation

    // none

    #endregion

  }
}
