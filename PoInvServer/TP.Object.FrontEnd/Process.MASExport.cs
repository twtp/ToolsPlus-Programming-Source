using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using TP.Object.Process.MASExport;

namespace TP.Object.FrontEnd.Process
{
  public class MASExport
  {

    public event EventHandler<StatusChangeEventArgs> StatusChange;

    private const string MAS_EXECUTABLE = "pvxwin32.exe";
    private const string MAS_HOME_DIRECTORY = "s:\\mas45server\\MAS90\\HOME";
    private const string MAS_SOTA_INI = "..\\LAUNCHER\\SOTA.INI";
    private const string MAS_STARTUP_M4P = "..\\SOA\\STARTUP.M4P";
    private const string MAS_LOGIN_USER = "BD";
    private const string MAS_LOGIN_PASS = "brian";
    private const string MAS_LOGIN_COMPANY = "TOO";

    public MASExport() { }

    public void RunExport()
    {
      bool go;

      // AP_TermsCode: no prereqs
      go = RunExport_AP_TermsCode();
      if (!go) { return; }

      // AP_Vendor: requires AP_TermsCode
      go = RunExport_AP_Vendor();
      if (!go) { return; }

      // IM_ProductLine: no prereqs
      /*go = RunExport_IM_ProductLine();
      if (!go) { return; }

      // IM_Warehouse: no prereqs
      go = RunExport_IM_Warehouse();
      if (!go) { return; }

      // CI_Item: requires IM_ProductLine, IM_Warehouse
      /*go = RunExport_CI_Item();
      if (!go) { return; }

      // IM_SalesKit(Header|Detail): requires CI_Item
      go = RunExport_IM_SalesKit();
      if (!go) { return; }

      // IM_PriceCode: requires CI_Item
      go = RunExport_IM_PriceCode();
      if (!go) { return; }

      // PO_PurchaseOrder(Header|Detail): requires CI_Item, AP_Vendor
      go = RunExport_PO_PurchaseOrder();
      if (!go) { return; }*/

      return;
    }

    public bool RunExport_AP_TermsCode() { return this.runExport<AP_TermsCode>("AP_TermsCode"); }
    public bool RunExport_AP_Vendor() { return this.runExport<AP_Vendor>("AP_Vendor"); }
    public bool RunExport_IM_ProductLine() { return this.runExport<IM_ProductLine>("IM_ProductLine"); }
    public bool RunExport_IM_Warehouse() { return this.runExport<IM_Warehouse>("IM_Warehouse"); }
    public bool RunExport_CI_Item() { return this.runExport<CI_Item>("CI_Item"); }
    public bool RunExport_IM_SalesKit() { return this.runExport<IM_SalesKit>("IM_SalesKit"); }
    public bool RunExport_IM_PriceCode() { return this.runExport<IM_PriceCode>("IM_PriceCode"); }
    public bool RunExport_PO_PurchaseOrder() { return this.runExport<PO_PurchaseOrder>("PO_PurchaseOrder"); }

    private bool runExport<T>(string urlPart) where T : IMASExportObject
    {
      var def = this.query<T>(urlPart);
      if (def == null || def.DataFileContent == null || def.DataFileContent.Count() == 0)
      {
        this.onStatusUpdate(new StatusChangeEventArgs(urlPart, "Nothing to do here"));
        return true;
      }

      this.onStatusUpdate(new StatusChangeEventArgs(urlPart, string.Format("Found {0} row(s) to export, starting MAS...", def.DataFileContent.Count())));
      this.doExport<T>(def);
      var results = this.validate(def, urlPart);
      
      if (results.Count() == 0) {
        this.onStatusUpdate(new StatusChangeEventArgs(urlPart, "Export completed without errors, yay!"));
        return true;
      }
      
      this.onStatusUpdate(new StatusChangeEventArgs(urlPart, "ERROR: encountered problems exporting, see below!"));
      this.onStatusUpdate(new StatusChangeEventArgs(urlPart, results.Select(_ => _.ToString())));
      return false;
    }


    private JobDefinition<T> query<T>(string urlPart) where T : IMASExportObject
    {
      this.onStatusUpdate(new StatusChangeEventArgs(urlPart, "Asking for changes..."));
      return Extensions.Get<JobDefinition<T>>("process", "masexport", "query", urlPart);
    }

    private JobErrorReport[] validate<T>(JobDefinition<T> def, string urlPart) where T : IMASExportObject
    {
      return Extensions.Post<JobDefinition<T>, JobErrorReport[]>("TP.Object.Process.MASExport.JobDefinition", def, "process", "masexport", "validate", urlPart);
    }

    private void doExport<T>(JobDefinition<T> def, Action<StringBuilder> header = null) where T : IMASExportObject
    {
      if (def == null) { return; }

      StringBuilder sb = new StringBuilder();
      if (header != null)
      {
        header(sb);
        sb.AppendLine();
      }
      foreach (var ln in def.DataFileContent)
      {
        ln.ToCSVLine(sb);
        sb.AppendLine();
      }
      System.IO.File.WriteAllText(def.MutexFileURI, "mutex file");
      System.IO.File.WriteAllText(def.DataFileURI, sb.ToString());

      string displayMode = def.DisplayMode == DisplayMode.Display ? "DISPLAY" : "MANUAL";

      var psi = new System.Diagnostics.ProcessStartInfo
      {
        FileName = MAS_HOME_DIRECTORY + "\\" + MAS_EXECUTABLE,
        Arguments = string.Format("{0} {1} {2} {3} {4} {5} {6} {7}", MAS_SOTA_INI, MAS_STARTUP_M4P, "-ARG DIRECT UION", MAS_LOGIN_USER, MAS_LOGIN_PASS, MAS_LOGIN_COMPANY, def.JobCompiledName, displayMode),
        WorkingDirectory = MAS_HOME_DIRECTORY,
        UseShellExecute = false
      };
      System.Diagnostics.Process.Start(psi);

      while (System.IO.File.Exists(def.MutexFileURI))
      {
        System.Threading.Thread.Sleep(1000);
      }

      return;
    }

    private void onStatusUpdate(StatusChangeEventArgs e)
    {
      EventHandler<StatusChangeEventArgs> handler = this.StatusChange;
      if (handler != null)
      {
        handler(this, e);
      }
    }


  }
}
