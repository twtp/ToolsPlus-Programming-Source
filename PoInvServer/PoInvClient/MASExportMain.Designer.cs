namespace PoInvClient
{
  partial class MASExportMain
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.ProgressListView = new System.Windows.Forms.ListView();
      this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.ExportStartButton = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // ProgressListView
      // 
      this.ProgressListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
      this.ProgressListView.GridLines = true;
      this.ProgressListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
      this.ProgressListView.HideSelection = false;
      this.ProgressListView.Location = new System.Drawing.Point(12, 12);
      this.ProgressListView.Name = "ProgressListView";
      this.ProgressListView.Size = new System.Drawing.Size(582, 346);
      this.ProgressListView.TabIndex = 0;
      this.ProgressListView.UseCompatibleStateImageBehavior = false;
      this.ProgressListView.View = System.Windows.Forms.View.Details;
      // 
      // columnHeader1
      // 
      this.columnHeader1.Text = "Module";
      this.columnHeader1.Width = 100;
      // 
      // columnHeader2
      // 
      this.columnHeader2.Text = "Message";
      this.columnHeader2.Width = 450;
      // 
      // ExportStartButton
      // 
      this.ExportStartButton.Location = new System.Drawing.Point(140, 364);
      this.ExportStartButton.Name = "ExportStartButton";
      this.ExportStartButton.Size = new System.Drawing.Size(92, 23);
      this.ExportStartButton.TabIndex = 1;
      this.ExportStartButton.Text = "Start";
      this.ExportStartButton.UseVisualStyleBackColor = true;
      this.ExportStartButton.Click += new System.EventHandler(this.ExportStartButton_Click);
      // 
      // MASExportMain
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(606, 399);
      this.Controls.Add(this.ExportStartButton);
      this.Controls.Add(this.ProgressListView);
      this.Name = "MASExportMain";
      this.Text = "MASExportMain";
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.ListView ProgressListView;
    private System.Windows.Forms.ColumnHeader columnHeader1;
    private System.Windows.Forms.ColumnHeader columnHeader2;
    private System.Windows.Forms.Button ExportStartButton;
  }
}