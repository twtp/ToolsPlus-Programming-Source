using System;
using System.Collections.Generic;
using System.Linq;
using Nancy;

using TP.Object.Process.MASExport;
using TP.Object.BackEnd.Process;

namespace PoInvServer
{

  public class MASExportQueryValidate : NancyModule
  {
    public MASExportQueryValidate() : base("/process/masexport")
    {
      #region AP_TermsCode

      Get["/query/AP_TermsCode"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<AP_TermsCode> job = m.GetTermsCodeExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/AP_TermsCode"] = _ =>
      {
        try
        {
          JobDefinition<AP_TermsCode> job = NancyExtensions.GetDbObjectParameter<JobDefinition<AP_TermsCode>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateTermsCodeExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region AP_Vendor

      Get["/query/AP_Vendor"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<AP_Vendor> job = m.GetVendorExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/AP_Vendor"] = _ =>
      {
        try
        {
          JobDefinition<AP_Vendor> job = NancyExtensions.GetDbObjectParameter<JobDefinition<AP_Vendor>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateVendorExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region IM_ProductLine

      Get["/query/IM_ProductLine"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<IM_ProductLine> job = m.GetProductLineExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/IM_ProductLine"] = _ =>
      {
        try
        {
          JobDefinition<IM_ProductLine> job = NancyExtensions.GetDbObjectParameter<JobDefinition<IM_ProductLine>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateProductLineExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region IM_Warehouse

      Get["/query/IM_Warehouse"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<IM_Warehouse> job = m.GetWarehouseExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/IM_Warehouse"] = _ =>
      {
        try
        {
          JobDefinition<IM_Warehouse> job = NancyExtensions.GetDbObjectParameter<JobDefinition<IM_Warehouse>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateWarehouseExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region CI_Item

      Get["/query/CI_Item"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<CI_Item> job = m.GetItemExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/CI_Item"] = _ =>
      {
        try
        {
          JobDefinition<CI_Item> job = NancyExtensions.GetDbObjectParameter<JobDefinition<CI_Item>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateItemExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region IM_SalesKit

      Get["/query/IM_SalesKit"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<IM_SalesKit> job = m.GetSalesKitExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/IM_SalesKit"] = _ =>
      {
        try
        {
          JobDefinition<IM_SalesKit> job = NancyExtensions.GetDbObjectParameter<JobDefinition<IM_SalesKit>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateSalesKitExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region IM_PriceCode

      Get["/query/IM_PriceCode"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<IM_PriceCode> job = m.GetItemPriceExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/IM_PriceCode"] = _ =>
      {
        try
        {
          JobDefinition<IM_PriceCode> job = NancyExtensions.GetDbObjectParameter<JobDefinition<IM_PriceCode>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidateItemPriceExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region PO_PurchaseOrder

      Get["/query/PO_PurchaseOrder"] = _ =>
      {
        try
        {
          MASExport m = new MASExport();
          JobDefinition<PO_PurchaseOrder> job = m.GetPurchaseOrderExport();
          return NancyExtensions.Respond(job);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Post["/validate/PO_PurchaseOrder"] = _ =>
      {
        try
        {
          JobDefinition<PO_PurchaseOrder> job = NancyExtensions.GetDbObjectParameter<JobDefinition<PO_PurchaseOrder>>(Request.Form, "TP.Object.Process.MASExport.JobDefinition");
          MASExport m = new MASExport();
          JobErrorReport[] report = m.ValidatePurchaseOrderExport(job);
          return NancyExtensions.Respond(report);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }
}