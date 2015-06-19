using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

using TP.Object.Process.MASExport;

namespace TP.Object.BackEnd.Process
{
  public class MASExport
  {
    public MASExport() { }


    #region AP_TermsCode

    public JobDefinition<AP_TermsCode> GetTermsCodeExport()
    {
      var lines = this.getLines<PurchaseTerms, AP_TermsCode>(
          getter: () => DbPurchaseTerms.GetExportRequired(),
          mutator: (tc) => AP_TermsCode.FromObject(tc)
        );
      return new JobDefinition<AP_TermsCode>
      {
        JobFriendlyName = "POINV2_AP_TERMS",
        JobCompiledName = "VIWI0F",
        DataFileName = "poinv2_ap_terms.csv",
        MutexFileName = "poinv2_ap_terms.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateTermsCodeExport(JobDefinition<AP_TermsCode> job)
    {
      // This is a beast, but it's actually pretty straightforward. validator()
      // iterates over each item in records, and double checks against MAS for
      // correctness. Most of the SQL/MAS specific stuff gets pushed into the
      // definitions for the IMASExportObject classes, except for the primary
      // key stuff for the query. Could use an intermediate object to build
      // the QueryParameter I guess?
      //
      //   records: just the list of IMASExportObject that this function
      //                  handles, the iteration list
      //   getter: returns a data table from MAS, which should have 1 row for
      //                  the specific item we're iterating over
      //   key: unique identifier for this IMASExportObject, so we can have 
      //                  something to send to an error report, if required
      //   fieldsCheck: nested iteration that goes over each field we attempted
      //                  to export, accumulating an error count to return back
      //                  a bool success/error
      //   updater: gets the current database object for the item, and
      //                  recompares it to our exportable, if the fields are
      //                  still the same, the object can be un-flagged,
      //                  otherwise, it remains flagged for the next offload
      //
      string sql = "SELECT " + AP_TermsCode.MASColumns() + " FROM " + AP_TermsCode.MASTableName() + " WHERE TermsCode=@tc";
      Func<AP_TermsCode, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@tc", SqlDbType.VarChar, i.TermsCode));
      Func<AP_TermsCode, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, AP_TermsCode, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<AP_TermsCode> updater = (i) => i.FlagExportedIfUnchanged(() => DbPurchaseTerms.GetByID(i.TermsCode));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region AP_Vendor

    public JobDefinition<AP_Vendor> GetVendorExport()
    {
      var lines = this.getLines<Vendor, AP_Vendor>(
          getter: () => DbVendor.GetExportRequired(),
          mutator: (v) => AP_Vendor.FromObject(v)
        );

      return new JobDefinition<AP_Vendor>
      {
        JobFriendlyName = "POINV2_AP_VENDOR",
        JobCompiledName = "VIWI0H",
        DataFileName = "poinv2_ap_vendor.csv",
        MutexFileName = "poinv2_ap_vendor.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateVendorExport(JobDefinition<AP_Vendor> job)
    {
      string sql = "SELECT " + AP_Vendor.MASColumns() + " FROM " + AP_Vendor.MASTableName() + " WHERE  APDivisionNo='00' AND VendorNo=@v";
      Func<AP_Vendor, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@v", System.Data.SqlDbType.VarChar, i.VendorNo));
      Func<AP_Vendor, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, AP_Vendor, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<AP_Vendor> updater = (i) => i.FlagExportedIfUnchanged(() => DbVendor.GetByID(i.VendorNo));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region IM_ProductLine

    public JobDefinition<IM_ProductLine> GetProductLineExport()
    {
      var lines = this.getLines<ProductLine, IM_ProductLine>(
          getter: () => DbProductLine.GetExportRequired(),
          mutator: (tc) => IM_ProductLine.FromObject(tc)
        );
      return new JobDefinition<IM_ProductLine>
      {
        JobFriendlyName = "POINV2_IM_PRODU",
        JobCompiledName = "VIWI0J",
        DataFileName = "poinv2_im_productline.csv",
        MutexFileName = "poinv2_im_productline.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateProductLineExport(JobDefinition<IM_ProductLine> job)
    {
      string sql = "SELECT " + IM_ProductLine.MASColumns() + " FROM " + IM_ProductLine.MASTableName() + " WHERE ProductLine=@pl";
      Func<IM_ProductLine, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@pl", SqlDbType.VarChar, i.ProductLineCode));
      Func<IM_ProductLine, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, IM_ProductLine, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<IM_ProductLine> updater = (i) => i.FlagExportedIfUnchanged(() => DbProductLine.GetByPrefix(i.ProductLineCode));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region IM_Warehouse

    public JobDefinition<IM_Warehouse> GetWarehouseExport()
    {
      var lines = this.getLines<MasWarehouse, IM_Warehouse>(
          getter: () => TP.Object.BackEnd.DbMasWarehouse.GetExportRequired(),
          mutator: (k) => IM_Warehouse.FromObject(k)
        );

      return new JobDefinition<IM_Warehouse>
      {
        JobFriendlyName = "POINV2_IM_WAREH",
        JobCompiledName = "VIWI00",
        DataFileName = "poinv2_im_warehouse.csv",
        MutexFileName = "poinv2_im_warehouse.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateWarehouseExport(JobDefinition<IM_Warehouse> job)
    {
      string sql = "SELECT " + IM_Warehouse.MASColumns() + " FROM " + IM_Warehouse.MASTableName() + " WHERE WarehouseCode=@w";
      Func<IM_Warehouse, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@w", SqlDbType.VarChar, i.WarehouseCode));
      Func<IM_Warehouse, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, IM_Warehouse, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<IM_Warehouse> updater = (i) => i.FlagExportedIfUnchanged(() => DbMasWarehouse.GetByCode(i.WarehouseCode));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region CI_Item - TODO

    public JobDefinition<CI_Item> GetItemExport()
    {
      var lines = this.getLines<Item, CI_Item>(
          getter: () => DbItem.GetExportRequired(),
          mutator: (k) => CI_Item.FromObject(k)
        );

      return new JobDefinition<CI_Item>
      {
        JobFriendlyName = "POINV2_IM_ITEM",
        JobCompiledName = "VIWI00",
        DataFileName = "poinv2_im_item.csv",
        MutexFileName = "poinv2_im_item.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateItemExport(JobDefinition<CI_Item> job)
    {
      return this.validator(
          records: job.DataFileContent,
          getter: (i) => TP.Database.MASConnection.Retrieve("SELECT * FROM CI_Item WHERE ItemCode=@i", new TP.Database.QueryParameter("@i", System.Data.SqlDbType.VarChar, i.ItemCode)),
          key: (i) => i.ItemCode,
          fieldsCheck: (rv, i, r, k) =>
          {
            int errorCount = 0;
            //errorCount += validatorCheck(rv, k, "SalesKitNo", i.SalesKitNo, r.Field<string>("SalesKitNo"));
            //errorCount += validatorCheck(rv, k, "ComponentItemCode", i.ComponentItemCode, r.Field<string>("ComponentItemCode"));
            //errorCount += validatorCheck(rv, k, "QuantityPerAssembly", i.QuantityPerAssembly, (int)r.Field<decimal>("QuantityPerAssembly"));
            return errorCount == 0;
          },
          //updater: (v) => { v.OriginalObject.FlagExported(); }
          updater: null
        );
    }

    #endregion


    #region IM_SalesKit

    public JobDefinition<IM_SalesKit> GetSalesKitExport()
    {
      var lines = this.getLines<ItemKitContent, IM_SalesKit>(
          getter: () => DbItemKitContent.GetExportRequired(),
          mutator: (k) => IM_SalesKit.FromObject(k)
        );

      return new JobDefinition<IM_SalesKit>
      {
        JobFriendlyName = "POINV2_IM_SALES",
        JobCompiledName = "VIWI00",
        DataFileName = "poinv2_im_saleskit.csv",
        MutexFileName = "poinv2_im_saleskit.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateSalesKitExport(JobDefinition<IM_SalesKit> job)
    {
      string sql = "SELECT " + IM_SalesKit.MASColumns() + " FROM " + IM_SalesKit.MASTableName() + " WHERE SalesKitNo=@k AND ComponentItemCode=@c";
      Func<IM_SalesKit, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@k", System.Data.SqlDbType.VarChar, i.SalesKitNo), new TP.Database.QueryParameter("@c", System.Data.SqlDbType.VarChar, i.ComponentItemCode));
      Func<IM_SalesKit, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, IM_SalesKit, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<IM_SalesKit> updater = (i) => i.FlagExportedIfUnchanged(() => DbItemKitContent.GetByPrimaryKey(i.SalesKitNo, i.ComponentItemCode));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region IM_PriceCode

    public JobDefinition<IM_PriceCode> GetItemPriceExport()
    {
      var lines = this.getLines<ItemPrice, IM_PriceCode>(
          getter: () => DbItemPrice.GetExportRequired(),
          filter: (p) => p.SalesChannelID == 0 || p.SalesChannelID == 1, // TODO: factor this out
          mutator: (p) => IM_PriceCode.FromObject(p)
        );
      lines.Concat(this.getLines<Item, IM_PriceCode>(
          getter: () => DbItem.GetExportRequiredPriceLevels(),
          mutator: (i) => IM_PriceCode.FromObject(i)
        ));

      return new JobDefinition<IM_PriceCode>
      {
        JobFriendlyName = "POINV2_IM_PRICE",
        JobCompiledName = "VIWI00",
        DataFileName = "poinv2_im_pricecode.csv",
        MutexFileName = "poinv2_im_pricecode.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidateItemPriceExport(JobDefinition<IM_PriceCode> job)
    {
      string sql = "SELECT " + IM_PriceCode.MASColumns() + " FROM " + IM_PriceCode.MASTableName() + " WHERE ItemCode=@i AND CustomerPriceLevel=@p";
      Func<IM_PriceCode, DataTable> getter = (i) => TP.Database.MASConnection.Retrieve(sql, new TP.Database.QueryParameter("@i", System.Data.SqlDbType.VarChar, i.ItemCode), new TP.Database.QueryParameter("@p", System.Data.SqlDbType.VarChar, i.CustomerPriceLevel));
      Func<IM_PriceCode, string> key = (i) => i.PrimaryKey();
      Func<List<JobErrorReport>, IM_PriceCode, DataRow, string, bool> fieldsCheck = (rv, i, r, k) => 0 == i.Validate(r).Sum(f => validatorCheck(rv, k, f.Item1, f.Item2, f.Item3));
      Action<IM_PriceCode> updater = (i) => i.FlagExportedIfUnchanged(() => DbItemPrice.GetByID(i.InternalID));
      return this.validator(job.DataFileContent, getter, key, fieldsCheck, updater);
    }

    #endregion


    #region PO_PurchaseOrder

    public JobDefinition<PO_PurchaseOrder> GetPurchaseOrderExport()
    {
      var lines = this.getLines<PurchaseOrder, PO_PurchaseOrder>(
          getter: () => TP.Object.BackEnd.DbPurchaseOrder.GetExportRequired(),
          mutator: (po) => PO_PurchaseOrder.FromObject(po)
        );
      return new JobDefinition<PO_PurchaseOrder>
      {
        JobFriendlyName = "POINV2_PO_PURCH",
        JobCompiledName = "VIWI0G",
        DataFileName = "poinv2_po_purchaseorder.csv",
        MutexFileName = "poinv2_po_purchaseorder.txt",
        DataFileContent = lines,
        DisplayMode = DisplayMode.Manual,
      };
    }

    public JobErrorReport[] ValidatePurchaseOrderExport(JobDefinition<PO_PurchaseOrder> job)
    {
      return this.validator(
          records: job.DataFileContent,
          getter: (i) => TP.Database.MASConnection.Retrieve("SELECT PurchaseOrderNo, PurchaseOrderDate, RequiredExpireDate, VendorNo, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZipCode, TermsCode FROM PO_PurchaseOrderHeader WHERE PurchaseOrderNo=@po", new TP.Database.QueryParameter("@po", SqlDbType.VarChar, i.PurchaseOrderNo)),
          key: (i) => i.PurchaseOrderNo,
          fieldsCheck: (rv, i, r, k) =>
          {
            int errorCount = 0;
            errorCount += validatorCheck(rv, k, "PurchaseOrderNo", i.PurchaseOrderNo, r.Field<string>("PurchaseOrderNo"));
            errorCount += validatorCheck(rv, k, "PurchaseOrderDate", i.OrderDate, r.Field<DateTime>("PurchaseOrderDate"));
            errorCount += validatorCheck(rv, k, "RequiredExpireDate", i.ExpireDate, r.Field<DateTime>("RequiredExpireDate"));
            errorCount += validatorCheck(rv, k, "VendorNo", i.VendorNo, r.Field<string>("VendorNo"));
            errorCount += validatorCheck(rv, k, "ShipToCode", i.ShipToCode, r.Field<string>("ShipToCode"));
            if (i.ShipToCode == "1234")
            {
              errorCount += validatorCheck(rv, k, "ShipToName", i.ShipToName, r.Field<string>("ShipToName"));
              errorCount += validatorCheck(rv, k, "ShipToAddress1", i.ShipToAddress1, r.Field<string>("ShipToAddress1"));
              errorCount += validatorCheck(rv, k, "ShipToAddress2", i.ShipToAddress2, r.Field<string>("ShipToAddress2"));
              errorCount += validatorCheck(rv, k, "ShipToCity", i.ShipToCity, r.Field<string>("ShipToCity"));
              errorCount += validatorCheck(rv, k, "ShipToState", i.ShipToState, r.Field<string>("ShipToState"));
              errorCount += validatorCheck(rv, k, "ShipToZipCode", i.ShipToZipCode, r.Field<string>("ShipToZipCode"));
            }
            errorCount += validatorCheck(rv, k, "TermsCode", i.TermsCode, r.Field<string>("TermsCode"));
            
            // XXX: This does not check for comment lines, but i'm not sure
            // how those would fail to go over. Additionally, this assumes
            // that 1 item == 1 row, which may change in the future (adding
            // items at zero cost?)
            this.validator(
                records: i.DetailLines,
                getter: (l) => TP.Database.MASConnection.Retrieve("SELECT ItemCode, QuantityOrdered, UnitCost, CommentText FROM PO_PurchaseOrderDetail WHERE PurchaseOrderNo=@po AND ItemCode=@item", new TP.Database.QueryParameter("@po", SqlDbType.VarChar, i.PurchaseOrderNo), new TP.Database.QueryParameter("@item", SqlDbType.VarChar, l.ItemCode)),
                key: (l) => string.Format("{0}/{1}", i.PurchaseOrderNo, l.ItemCode),
                fieldsCheck: (l_rv, l_i, l_r, l_k) =>
                {
                  errorCount += validatorCheck(rv, l_k, "ItemCode", l_i.ItemCode, l_r.Field<string>("ItemCode"));
                  errorCount += validatorCheck(rv, l_k, "QuantityOrdered", l_i.QuantityOrdered, (int)l_r.Field<decimal>("QuantityOrdered"));
                  errorCount += validatorCheck(rv, l_k, "UnitCost", l_i.UnitCost, l_r.Field<decimal>("UnitCost"));
                  errorCount += validatorCheck(rv, l_k, "CommentText", l_i.CommentText, l_r.Field<string>("CommentText"));
                  return false;
                },
                updater: null
              );

            return errorCount == 0;
          },
          //updater: (po) => po.OriginalObject.FlagExported()
          updater: null
        );
    }

    #endregion





    private IEnumerable<TTo> getLines<TFrom, TTo>(Func<IEnumerable<TFrom>> getter, Func<TFrom, TTo> mutator, Func<TFrom, bool> filter = null) where TTo : IMASExportObject where TFrom : TP.Object.BaseBusinessObject
    {
      var list = getter();
      if (list.Count() == 0) { return null; }

      if (filter != null)
      {
        list = list.Where(filter);
      }

      return list.Select(mutator).ToArray();
    }

    private IEnumerable<TTo> getLines<TFrom, TTo>(Func<IEnumerable<TFrom>> getter, Func<TFrom, IEnumerable<TTo>> mutator, Func<TFrom, bool> filter = null) where TTo : IMASExportObject where TFrom : TP.Object.BaseBusinessObject
    {
      var list = getter();
      if (list.Count() == 0) { return null; }

      if (filter != null)
      {
        list = list.Where(filter);
      }

      return list.SelectMany(mutator).ToArray();
    }

    private JobErrorReport[] validator<T>(IEnumerable<T> records, Func<T, DataTable> getter, Func<T, string> key, Func<List<JobErrorReport>, T, DataRow, string, bool> fieldsCheck, Action<T> updater)
    {
      var retval = new List<JobErrorReport>();
      foreach (var i in records)
      {
        var dt = getter(i);
        if (dt.Rows.Count != 1)
        {
          retval.Add(new JobErrorReport { Key = key(i), Field = "CREATE" });
        }
        else
        {
          bool exportedOK = fieldsCheck(retval, i, dt.Rows[0], key(i));
          if (exportedOK && updater != null)
          {
            updater(i);
          }
        }
      }
      
      return retval.ToArray();
    }

    private int validatorCheck<T>(List<JobErrorReport> retval, string keyValue, string fieldName, T correctFieldValue, T currentFieldValue)
    {
      if (object.Equals(correctFieldValue, currentFieldValue))
      {
        System.Diagnostics.Debug.WriteLine(string.Format("validatorCheck {0}: {1}: VALID ({2} {3})", keyValue, fieldName, correctFieldValue.ToString(), currentFieldValue.ToString()));
        return 0;
      }
      else
      {
        // fix for MAS null handling
        if (correctFieldValue is string)
        {
          var cfv = correctFieldValue as string;
          if (cfv == string.Empty && currentFieldValue == null)
          {
            System.Diagnostics.Debug.WriteLine(string.Format("validatorCheck {0}: {1}: VALID ({2} {3})", keyValue, fieldName, correctFieldValue.ToString(), "NULL"));
            return 0;
          }
        }

        System.Diagnostics.Debug.WriteLine(string.Format("validatorCheck {0}: {1}: ERROR ({2} {3})", keyValue, fieldName, correctFieldValue.ToString(), currentFieldValue == null ? "NULL" : currentFieldValue.ToString()));
        retval.Add(new JobErrorReport { Key = keyValue, Field = fieldName, Value = currentFieldValue == null ? "NULL" : currentFieldValue.ToString() });
        return 1;
      }
    }
  }
}
