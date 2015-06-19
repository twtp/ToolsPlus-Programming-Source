namespace FeedGenerator
{
    partial class DataBackend
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
            this.feedsBox = new System.Windows.Forms.GroupBox();
            this.feedsLVI = new System.Windows.Forms.ListView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataLocationLVI = new System.Windows.Forms.ListView();
            this.button1 = new System.Windows.Forms.Button();
            this.tablesLst = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.columnsLst = new System.Windows.Forms.ListBox();
            this.queryRefreshBtn = new System.Windows.Forms.Button();
            this.queryFromLineTxt = new System.Windows.Forms.TextBox();
            this.feedsBox.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // feedsBox
            // 
            this.feedsBox.Controls.Add(this.feedsLVI);
            this.feedsBox.Location = new System.Drawing.Point(12, 12);
            this.feedsBox.Name = "feedsBox";
            this.feedsBox.Size = new System.Drawing.Size(252, 376);
            this.feedsBox.TabIndex = 3;
            this.feedsBox.TabStop = false;
            this.feedsBox.Text = "Feeds";
            // 
            // feedsLVI
            // 
            this.feedsLVI.Location = new System.Drawing.Point(6, 19);
            this.feedsLVI.Name = "feedsLVI";
            this.feedsLVI.Size = new System.Drawing.Size(238, 348);
            this.feedsLVI.TabIndex = 3;
            this.feedsLVI.UseCompatibleStateImageBehavior = false;
            this.feedsLVI.SelectedIndexChanged += new System.EventHandler(this.feedsLVI_SelectedIndexChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(270, 1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(616, 387);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.dataLocationLVI);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(608, 361);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Data Locations";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.queryFromLineTxt);
            this.tabPage2.Controls.Add(this.queryRefreshBtn);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.columnsLst);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.tablesLst);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(608, 361);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Query Builder";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataLocationLVI
            // 
            this.dataLocationLVI.Location = new System.Drawing.Point(6, 25);
            this.dataLocationLVI.Name = "dataLocationLVI";
            this.dataLocationLVI.Size = new System.Drawing.Size(596, 330);
            this.dataLocationLVI.TabIndex = 0;
            this.dataLocationLVI.UseCompatibleStateImageBehavior = false;
            this.dataLocationLVI.SelectedIndexChanged += new System.EventHandler(this.dataLocationLVI_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(539, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(63, 22);
            this.button1.TabIndex = 1;
            this.button1.Text = "Refresh";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tablesLst
            // 
            this.tablesLst.FormattingEnabled = true;
            this.tablesLst.Location = new System.Drawing.Point(51, 34);
            this.tablesLst.Name = "tablesLst";
            this.tablesLst.Size = new System.Drawing.Size(200, 173);
            this.tablesLst.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(102, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Tables To Join";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(363, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Columns Identified";
            // 
            // columnsLst
            // 
            this.columnsLst.FormattingEnabled = true;
            this.columnsLst.Location = new System.Drawing.Point(312, 34);
            this.columnsLst.Name = "columnsLst";
            this.columnsLst.Size = new System.Drawing.Size(200, 173);
            this.columnsLst.TabIndex = 2;
            // 
            // queryRefreshBtn
            // 
            this.queryRefreshBtn.Location = new System.Drawing.Point(518, 8);
            this.queryRefreshBtn.Name = "queryRefreshBtn";
            this.queryRefreshBtn.Size = new System.Drawing.Size(73, 27);
            this.queryRefreshBtn.TabIndex = 4;
            this.queryRefreshBtn.Text = "Refresh";
            this.queryRefreshBtn.UseVisualStyleBackColor = true;
            this.queryRefreshBtn.Click += new System.EventHandler(this.queryRefreshBtn_Click);
            // 
            // queryFromLineTxt
            // 
            this.queryFromLineTxt.Location = new System.Drawing.Point(51, 232);
            this.queryFromLineTxt.Multiline = true;
            this.queryFromLineTxt.Name = "queryFromLineTxt";
            this.queryFromLineTxt.Size = new System.Drawing.Size(461, 107);
            this.queryFromLineTxt.TabIndex = 5;
            // 
            // DataBackend
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(898, 398);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.feedsBox);
            this.Name = "DataBackend";
            this.Text = "DataBackend";
            this.Load += new System.EventHandler(this.DataBackend_Load);
            this.feedsBox.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox feedsBox;
        private System.Windows.Forms.ListView feedsLVI;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ListView dataLocationLVI;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button queryRefreshBtn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox columnsLst;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox tablesLst;
        private System.Windows.Forms.TextBox queryFromLineTxt;
    }
}