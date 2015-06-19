using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PoInvClient
{
  public partial class MASExportMain : Form
  {
    public MASExportMain()
    {
      InitializeComponent();
    }

    private System.Threading.Thread worker = null;

    private void ExportStartButton_Click(object sender, EventArgs e)
    {
      this.ProgressListView.Items.Clear();
      this.ExportStartButton.Enabled = false;
      this.worker = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(this.runExport));
      this.worker.Start();
    }

    private void runExport(object data)
    {
      var m = new TP.Object.FrontEnd.Process.MASExport();
      m.StatusChange += this.guiUpdateProgress;
      m.RunExport();
      m.StatusChange -= this.guiUpdateProgress;
      this.guiUpdateComplete();
    }

    private void guiUpdateProgress(object sender, TP.Object.Process.MASExport.StatusChangeEventArgs e)
    {
      if (this.InvokeRequired)
      {
        this.Invoke((MethodInvoker)delegate { this.guiUpdateProgress(sender, e); });
        return;
      }

      foreach (var m in e.Message)
      {
        var lvi = new ListViewItem();
        lvi.Text = e.Module;
        lvi.SubItems.Add(m);
        this.ProgressListView.Items.Add(lvi);
      }
    }

    private void guiUpdateStart()
    {
      if (this.InvokeRequired)
      {
        this.Invoke((MethodInvoker)delegate { this.guiUpdateStart(); });
        return;
      }

      this.ProgressListView.Items.Clear();
      this.ExportStartButton.Enabled = false;
    }

    private void guiUpdateComplete()
    {
      if (this.InvokeRequired)
      {
        this.Invoke((MethodInvoker)delegate { this.guiUpdateComplete(); });
        return;
      }

      this.ExportStartButton.Enabled = true;
    }
  }
}
