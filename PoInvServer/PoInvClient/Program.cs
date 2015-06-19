using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PoInvClient
{
  static class Program
  {
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);


      /*var mf = new TP.Object.FrontEnd.Process.MASExport();
      mf.RunExport((a,b) => false);*/

      Application.Run(new MASExportMain());

      //Application.Run(new Form1());
      
    }
  }
}
