using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using TP.Database;

namespace TP.Object.BackEnd
{
  [System.Runtime.Serialization.DataContract(Name = "BaseProductLine", Namespace = "TP.Object")]
  public class DbBaseProductLine : TP.Object.BaseProductLine
  {
    
    #region Constructors

    public DbBaseProductLine(DataRow r) : base()
    {
      this.ProductLineID = (int)r["ID"];
      this.ProductLinePrefix = (string)r["ProductLine"];
      this.ProductLineName = this.nullableString(r["ManufFullName"]);
      this.IsDeleted = (bool)r["Hide"];
    }

    #endregion


    #region Getters

    // see DbProductLine

    #endregion

  }


  [System.Runtime.Serialization.DataContract(Name = "ProductLine", Namespace = "TP.Object")]
  public class DbProductLine : TP.Object.ProductLine
  {
    #region Cache

    private static AutoCache2Key<DbProductLine, int, string> cache = new AutoCache2Key<DbProductLine, int, string> {
        Constructor = _ => new DbProductLine(_),
        Identifier = _ => _.Field<int>("ID"),
        PrimaryKey = _ => _.ProductLineID,
        AlternateKey = _ => _.ProductLinePrefix,
        SecondsToLive = 86400
      };

    // no autocache
    private static CacheSimples cacheSimples = new CacheSimples();
    private class CacheSimples : CacheBase1Key<DbBaseProductLine, int>
    {
      public IEnumerable<DbBaseProductLine> Fetch(Func<DataTable> getter)
      {
        if (base.count() == 0)
        {
          DataTable dt = getter();
          foreach (DataRow r in dt.Rows)
          {
            base.put((int)r["ID"], new DbBaseProductLine(r));
          }
        }
        return base.values();
      }

      public void Clear()
      {
        base.wipe();
      }
    }

    #endregion


    #region Constructors

    protected DbProductLine() { }

    public DbProductLine(DataRow r) {
      this.ProductLineID = (int)r["ID"];
      this.ProductLinePrefix = (string)r["ProductLine"];
      this.DefaultVendorNo = this.nullableString(r["PrimaryVendorNumber"]);
      this.IsCoopVendor = (bool)r["Coop"];
      this.CoopVendor = this.nullableString(r["RealVendor"]);
      this.ProductLineName = this.nullableString(r["ManufFullName"]);
      this.ProductLineNameCleaned = this.nullableString(r["ManufFullNameCleaned"]);
      this.ProductLineNameWeb = this.nullableString(r["ManufFullNameWeb"]);
      this.NewItemsDefaultDropshippable = (bool)r["DefaultDropshippable"];
      this.DropshippableItemsDefaultToBackorderThreshold = (decimal)r["BOBelow"];
      this.NewItemsDefaultAvailabilityLimit = (int)r["AvailLimitDefault"];
      this.StandardLeadTime = (int)r["StdLeadTime"];
      this.ExportRequired = (bool)r["NewChk"];
      this.IsDeleted = (bool)r["Hide"];
      this.RepairNotes = this.nullableString(r["Repair"]);
      this.WarrantyNotes = this.nullableString(r["Warranty"]);
      this.CompanyInfoNotes = this.nullableString(r["CompanyInfo"]);
      this.ServiceCenterNotes = this.nullableString(r["ServiceCenters"]);
      this.CompanyAddress1 = this.nullableString(r["CompanyAddress1"]);
      this.CompanyAddress2 = this.nullableString(r["CompanyAddress2"]);
      this.CompanyCity = this.nullableString(r["CompanyCity"]);
      this.CompanyState = this.nullableString(r["CompanyState"]);
      this.CompanyPostalCode = this.nullableString(r["CompanyZip"]);
      this.CompanyCountry = this.nullableString(r["CompanyCountry"]);
      this.MinimumOrderAmount = this.nullableString(r["MinimumOrderAmount"]);
      this.PrepaidAmount = this.nullableString(r["PrepaidAmount"]);
      this.PrepaidSpecialAmount = this.nullableString(r["PrepaidSpecialAmount"]);
      this.ShipsFromPostalCode = this.nullableString(r["ShipsFromZip"]);
      this.ReconditionedURL = this.nullableString(r["ReconditionedURL"]);
      this.WebSectionID = this.nullable<int>(r["WebSectionID"]);
      this.DisplayMappViolationWarning = (bool)r["MappViolationWarning"];

    }

    #endregion


    #region Getters

    // /productline/{id:int}
    public static DbProductLine GetByID(int id, string[] preloads = null) { return GetByID(null, id, preloads); }
    internal static DbProductLine GetByID(System.Data.SqlClient.SqlTransaction txn, int id, string[] preloads = null)
    {
      var ret = cache.FetchByPrimaryKey(id, () => Connection.Retrieve(txn, "SELECT ID, ProductLine, PrimaryVendorNumber, Coop, RealVendor, ManufFullName, ManufFullNameCleaned, ManufFullNameWeb, DefaultDropshippable, BOBelow, AvailLimitDefault, StdLeadTime, NewChk, Hide, Repair, Warranty, CompanyInfo, ServiceCenters, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyCountry, MinimumOrderAmount, PrepaidAmount, PrepaidSpecialAmount, ShipsFromZip, ReconditionedURL, WebSectionID, MappViolationWarning FROM ProductLine WHERE ID=@id", new QueryParameter("@id", SqlDbType.Int, id)));
      return fillPreloads(ret, preloads, txn);
    }

    // /productline/{prefix}
    public static DbProductLine GetByPrefix(string prefix, string[] preloads = null) { return GetByPrefix(null, prefix, preloads); }
    internal static DbProductLine GetByPrefix(System.Data.SqlClient.SqlTransaction txn, string prefix, string[] preloads = null)
    {
      var ret = cache.FetchByAlternateKey(prefix, () => Connection.Retrieve(txn, "SELECT ID, ProductLine, PrimaryVendorNumber, Coop, RealVendor, ManufFullName, ManufFullNameCleaned, ManufFullNameWeb, DefaultDropshippable, BOBelow, AvailLimitDefault, StdLeadTime, NewChk, Hide, Repair, Warranty, CompanyInfo, ServiceCenters, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyCountry, MinimumOrderAmount, PrepaidAmount, PrepaidSpecialAmount, ShipsFromZip, ReconditionedURL, WebSectionID, MappViolationWarning FROM ProductLine WHERE ProductLine=@pl", new QueryParameter("@pl", SqlDbType.VarChar, prefix)));
      return fillPreloads(ret, preloads, txn);
    }

    public static IEnumerable<DbProductLine> GetExportRequired() { return GetExportRequired(null); }
    internal static IEnumerable<DbProductLine> GetExportRequired(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cache.FetchNonIndexed(() => Connection.Retrieve(txn, "SELECT ID, ProductLine, PrimaryVendorNumber, Coop, RealVendor, ManufFullName, ManufFullNameCleaned, ManufFullNameWeb, DefaultDropshippable, BOBelow, AvailLimitDefault, StdLeadTime, NewChk, Hide, Repair, Warranty, CompanyInfo, ServiceCenters, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyCountry, MinimumOrderAmount, PrepaidAmount, PrepaidSpecialAmount, ShipsFromZip, ReconditionedURL, WebSectionID, MappViolationWarning FROM ProductLine WHERE ExportRequired=1"));
      return fillPreloads(ret, null, txn);
    }

    // /productline/list
    public static IEnumerable<DbBaseProductLine> ListAll() { return ListAll(null); }
    internal static IEnumerable<DbBaseProductLine> ListAll(System.Data.SqlClient.SqlTransaction txn)
    {
      var ret = cacheSimples.Fetch(() => Connection.Retrieve(txn, "SELECT ID, ProductLine, ManufFullName, Hide FROM ProductLine"));
      return ret; // no clone, never preloading anything
    }

    private static DbProductLine fillPreloads(DbProductLine pl, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      if (pl == null)
      {
        return pl;
      }
      var retval = (DbProductLine)pl.MemberwiseClone();
      if (preloads != null)
      {
        if (preloads.Contains("wholesalepricing"))
        {
          retval.FillWholesalePriceLevels(txn);
        }
        if (preloads.Contains("vendor"))
        {
          retval.FillDefaultVendor(txn);
        }
        if (preloads.Contains("contactinfo"))
        {
          retval.FillContactInfo(txn);
        }
        if (preloads.Contains("dropshipcharges"))
        {
          retval.FillDropshipCharges(txn);
        }
        if (preloads.Contains("purchasingsubs"))
        {
          retval.FillIncomingSubstitutedProductLines(txn);
          retval.FillPurchasingSubstituteProductLine(txn);
        }
        if (preloads.Contains("salesreps"))
        {
          retval.FillSalesReps(txn); // this will additionally populate PLSR/SR
        }
        if (preloads.Contains("costfields"))
        {
          retval.FillCostFields(txn);
        }
      }
      return retval;
    }

    private static IEnumerable<DbProductLine> fillPreloads(IEnumerable<DbProductLine> coll, string[] preloads, System.Data.SqlClient.SqlTransaction txn)
    {
      return coll.Select(_ => fillPreloads(_, preloads, txn)).ToArray();
    }

    #endregion


    #region Links

    // DbProductLine has-one DbVendor
    public override bool FillDefaultVendor() { return FillDefaultVendor(null); }
    internal bool FillDefaultVendor(System.Data.SqlClient.SqlTransaction txn)
    {
      if (string.IsNullOrWhiteSpace(this.DefaultVendorNo))
      {
        this._vendor = null;
        return false;
      }
      else
      {
        if (this._vendor == null)
        {
          this._vendor = DbVendor.GetByID(txn, this.DefaultVendorNo);
        }
        return true;
      }
    }

    // DbProductLine has-many DbWholesalePriceLevel
    public override bool FillWholesalePriceLevels() { return FillWholesalePriceLevels(null); }
    internal bool FillWholesalePriceLevels(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._wholesalepricelevels == null)
      {
        this._wholesalepricelevels = DbProductLineWholesalePriceLevel.GetByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLine has-many DbContactInfo
    public override bool FillContactInfo() { return FillContactInfo(null); }
    internal bool FillContactInfo(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._contacts == null)
      {
        this._contacts = DbProductLineContactInfo.GetByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLine has-many DbDropshipCharge
    public override bool FillDropshipCharges() { return FillDropshipCharges(null); }
    internal bool FillDropshipCharges(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._dscharges == null)
      {
        this._dscharges = DbProductLineDropshipCharge.GetByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLine has-many DbProductLineSalesRep
    public override bool FillSalesReps() { return FillSalesReps(null); }
    internal bool FillSalesReps(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._salesreps == null)
      {
        this._salesreps = DbProductLineSalesRep.GetByProductLineID(txn, this.ProductLineID, new string[] { "reps" });
      }
      return true;
    }

    // DbProductLine has-one DbProductLineSubstitution
    public override bool FillPurchasingSubstituteProductLine() { return FillPurchasingSubstituteProductLine(null); }
    internal bool FillPurchasingSubstituteProductLine(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._sub == null)
      {
        this._sub = DbProductLineSubstitution.GetByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLine has-one DbProductLineSubstitution
    public override bool FillIncomingSubstitutedProductLines() { return FillIncomingSubstitutedProductLines(null); }
    internal bool FillIncomingSubstitutedProductLines(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._subincoming == null)
      {
        this._subincoming = DbProductLineSubstitution.GetByPurchasingProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    // DbProductLine has-one DbProductLineCostField
    public override bool FillCostFields() { return FillCostFields(null); }
    internal bool FillCostFields(System.Data.SqlClient.SqlTransaction txn)
    {
      if (this._costfields == null)
      {
        this._costfields = DbProductLineCostField.GetActiveByProductLineID(txn, this.ProductLineID);
      }
      return true;
    }

    #endregion


    #region Create/Update/Delete

    // /productline/create
    public static DbProductLine CreateProductLine(string prefix)
    {
      prefix = prefix.ToUpper();
      if (false == IsValidPrefix(prefix))
      {
        throw new ArgumentException("invalid product line prefix");
      }

      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLine check = GetByPrefix(txn, prefix);
        if (check != null)
        {
          throw new ArgumentException("product line already exists!");
        }

        int id = Connection.Insert(txn,
            "INSERT INTO ProductLine (ProductLine) VALUES (@pl)",
            new QueryParameter("@pl", SqlDbType.VarChar, prefix)
          );
        DbProductLineWholesalePriceLevel.CreateDefaultWholesalePriceLevels(id, txn);
        Connection.CommitTransaction(txn);

        cache.Clear(id, prefix);
        cacheSimples.Clear(); // dumb that we need to clear, should just create + add
        return GetByID(id);
      }
      catch (Exception)
      {
        Connection.RollbackTransaction(txn);
        throw;
      }
    }

    // /productline/update/{id:int}
    public DbProductLine UpdateProductLine()
    {
      var txn = Connection.BeginTransaction();
      try
      {
        DbProductLine orig = GetByID(txn, this.ProductLineID);

        Extensions.ObjectChanges changed = this.PropertiesModified(orig);
        if (changed.Count > 0)
        {
          int affected = Connection.Execute(txn,
              "UPDATE ProductLine SET " + changed.UpdateFieldString + " WHERE " + changed.UpdateWhereClause,
              changed.UpdateValues
            );
          if (affected == 0)
          {
            throw new ApplicationException("error updating!");
          }
        }
        Connection.CommitTransaction(txn);

        cache.Clear(this.ProductLineID, this.ProductLinePrefix);
        // nothing to do with baseproductline cache
        return GetByID(this.ProductLineID);
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
              "UPDATE ProductLine SET ExportRequired=0 WHERE ID=@id",
              new QueryParameter("@id", SqlDbType.Int, this.ProductLineID)
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

        cache.Clear(this.ProductLineID, this.ProductLinePrefix);
        // nothing to do with baseproductline cache
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

    internal static void AssertProductLineExists(int productLineID, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbProductLine pl = GetByID(txn, productLineID);
      if (pl == null)
      {
        throw new ArgumentException("product line does not exist");
      }
      return;
    }

    internal static void AssertProductLineExists(string productLinePrefix, System.Data.SqlClient.SqlTransaction txn = null)
    {
      DbProductLine pl = GetByPrefix(txn, productLinePrefix);
      if (pl == null)
      {
        throw new ArgumentException("product line does not exist");
      }
      return;
    }

    #endregion

  }
}
