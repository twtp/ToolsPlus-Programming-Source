using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "Barcode", Namespace = "TP.Object")]
  public class DbBarcode : TP.Object.Barcode
  {
    #region Cache

    private static AutoCache1Key1Index<DbBarcode, string, int> cache = new AutoCache1Key1Index<DbBarcode, string, int> {
        Constructor = _ => new DbBarcode(_),
        Identifier = _ => (string)_["Barcode"],
        PrimaryKey = _ => _.Code,
        Index1 = _ => _.ComponentID
      };

    #endregion


    #region Constructors

    public DbBarcode(DataRow r) : base()
    {
      this.Code = (string)r["Barcode"];
      this.ComponentID = (int)r["ComponentID"];
      this.IsInternal = (bool)r["Internal"];
      this.Quantity = (int)r["Quantity"];
      this.PrintByStandardPack = (bool)(r["PrintByStdPack"]);
    }

    #endregion


    #region Getters

    // /barcode/{code}
    public static DbBarcode GetByCode(string code, string[] preloads = null) { return GetByCode(null, code, preloads); }
    internal static DbBarcode GetByCode(System.Data.SqlClient.SqlTransaction txn, string code, string[] preloads = null)
    {
      code = code.ToUpper();
      var ret = cache.FetchByPrimaryKey(code, () => Connection.Retrieve(txn, "SELECT Barcode, ComponentID, Internal, Quantity, PrintByStdPack FROM InventoryComponentBarcodes WHERE Barcode=@bc", new QueryParameter("@bc", SqlDbType.VarChar, code)));
      return fillPreloads(ret, preloads, txn);
    }

    // /barcode/component/{id:int}
    public static IEnumerable<DbBarcode> GetByComponentID(int cid, string[] preloads = null) { return GetByComponentID(null, cid, preloads); }
    internal static IEnumerable<DbBarcode> GetByComponentID(System.Data.SqlClient.SqlTransaction txn, int cid, string[] preloads = null)
    {
      var ret = cache.FetchByIndex1(cid, () => Connection.Retrieve(txn, "SELECT Barcode, ComponentID, Internal, Quantity, PrintByStdPack FROM InventoryComponentBarcodes WHERE ComponentID=@id", new QueryParameter("@id", SqlDbType.Int, cid)));
      return fillPreloads(ret, preloads, txn);
    }

    private static DbBarcode fillPreloads(DbBarcode bc, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (bc == null)
      {
        return bc;
      }
      var retval = (DbBarcode)bc.MemberwiseClone();
      if (preloads != null)
      {
        // preloads go here for any links
      }
      return retval;
    }

    private static IEnumerable<DbBarcode> fillPreloads(IEnumerable<DbBarcode> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // none

    #endregion


    #region Create/Update/Delete

    // /barcode/create
    public static DbBarcode CreateBarcode(string code, int componentID, int quantity)
    {
      AssertIsValidBarcode(code);
      DbComponent.AssertComponentExists(componentID);

      var txn = Connection.BeginTransaction();
      try
      {
        DbBarcode check = GetByCode(code);
        if (check != null)
        {
          throw new ArgumentException("barcode already exists, and is mapped to component id " + check.ComponentID);
        }

        Connection.Execute(txn,
            "INSERT INTO InventoryComponentBarcodes (Barcode, ComponentID, Quantity) VALUES (@bc, @cid, @q)",
            new QueryParameter("@bc", SqlDbType.VarChar, code),
            new QueryParameter("@cid", SqlDbType.Int, componentID),
            new QueryParameter("@q", SqlDbType.Int, quantity)
          );
        Connection.CommitTransaction(txn);

        cache.Clear(code, componentID);
        return GetByCode(code);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /barcode/create/TP
    public static DbBarcode CreateToolsPlusBarcode(int componentID)
    {
      var txn = Connection.BeginTransaction();
      try
      {
        AssertComponentHasNoInternalBarcodes(componentID, txn);

        DataTable dt = Connection.Retrieve(txn, "SELECT MAX(CAST(SUBSTRING(Barcode,3,10) as int)) FROM InventoryComponentBarcodes WHERE Internal=1");
        string newBC = dt.Rows.Count == 0 ? "1" : Convert.ToString(dt.Rows[0][0]);
        newBC = string.Format("TP{0:0000000000}", newBC);
        Connection.Execute(txn,
            "INSERT INTO InventoryComponentBarcodes (Barcode, ComponentID, Quantity, Internal) VALUES (@bc, @cid, 1, 1)",
            new QueryParameter("@bc", SqlDbType.VarChar, newBC),
            new QueryParameter("@cid", SqlDbType.Int, componentID)
          );

        Connection.CommitTransaction(txn);

        return GetByCode(newBC);
      }
      catch
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /barcode/update/{code}
    public DbBarcode UpdateBarcode()
    {
      AssertIsValidBarcode(this.Code);
      
      var txn = Connection.BeginTransaction();
      try
      {
        DbComponent.AssertComponentExists(this.ComponentID, txn);

        DbBarcode orig = GetByCode(this.Code);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE InventoryComponentBarcodes SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.Code, this.ComponentID); // component id is non-modifiable, okay
        return GetByCode(this.Code);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public bool DeleteBarcode()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        this.DeleteBarcode(txn);
        Connection.CommitTransaction(txn);
        // cache handled already
        return true;
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    public bool DeleteBarcode(System.Data.SqlClient.SqlTransaction txn)
    {
      DbBarcode orig = GetByCode(txn, this.Code);

      Extensions.ObjectChanges changed = this.PropertiesModified(orig);
      if (changed.Count > 0)
      {
        throw new ArgumentException("barcode has been modified, not deleting");
      }

      Connection.Execute(txn,
          "DELETE FROM InventoryComponentBarcodes WHERE Barcode=@bc",
          new QueryParameter("@bc", SqlDbType.VarChar, this.Code)
        );

      cache.Clear(orig.Code, orig.ComponentID);
      return true;
    }

    #endregion


    #region Extra Validation

    internal static void AssertIsValidBarcode(string code)
    {
      if (!IsValidCode(code))
      {
        throw new ArgumentException("barcode is not valid");
      }
      return;
    }

    private static void AssertComponentHasNoInternalBarcodes(int componentID, System.Data.SqlClient.SqlTransaction txn)
    {
      DbComponent.AssertComponentExists(componentID, txn);
      DbComponent c = DbComponent.GetByID(txn, componentID);
      c.FillBarcodeList(txn);
      TP.Object.Barcode found = c.Barcodes.SingleOrDefault(_ => _.IsInternal == true);
      if (found != null)
      {
        throw new ArgumentException("component id " + componentID + " already has a TP barcode, " + found.Code);
      }
      return;
    }

    #endregion

  }
}
