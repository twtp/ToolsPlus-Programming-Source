namespace MasterCategoryConnectors
{
    partial class Form1
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
            this.MasterLVI = new System.Windows.Forms.ListView();
            this.ChildLVI = new System.Windows.Forms.ListView();
            this.hideChildConnectedChk = new System.Windows.Forms.CheckBox();
            this.hideMasterConnectedChk = new System.Windows.Forms.CheckBox();
            this.childCmb = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.connectBtn = new System.Windows.Forms.Button();
            this.disconnectBtn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.connectionsLVI = new System.Windows.Forms.ListView();
            this.hideChildHighPriceChk = new System.Windows.Forms.CheckBox();
            this.MasterCountLbl = new System.Windows.Forms.Label();
            this.ChildCountLbl = new System.Windows.Forms.Label();
            this.ConnectedCountLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // MasterLVI
            // 
            this.MasterLVI.BackColor = System.Drawing.Color.White;
            this.MasterLVI.HideSelection = false;
            this.MasterLVI.Location = new System.Drawing.Point(12, 30);
            this.MasterLVI.Name = "MasterLVI";
            this.MasterLVI.Size = new System.Drawing.Size(442, 397);
            this.MasterLVI.TabIndex = 0;
            this.MasterLVI.UseCompatibleStateImageBehavior = false;
            this.MasterLVI.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.MasterLVI_ColumnClick);
            this.MasterLVI.SelectedIndexChanged += new System.EventHandler(this.MasterLVI_SelectedIndexChanged);
            // 
            // ChildLVI
            // 
            this.ChildLVI.HideSelection = false;
            this.ChildLVI.Location = new System.Drawing.Point(568, 62);
            this.ChildLVI.Name = "ChildLVI";
            this.ChildLVI.Size = new System.Drawing.Size(422, 365);
            this.ChildLVI.TabIndex = 1;
            this.ChildLVI.UseCompatibleStateImageBehavior = false;
            this.ChildLVI.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.ChildLVI_ColumnClick);
            this.ChildLVI.SelectedIndexChanged += new System.EventHandler(this.ChildLVI_SelectedIndexChanged);
            // 
            // hideChildConnectedChk
            // 
            this.hideChildConnectedChk.AutoSize = true;
            this.hideChildConnectedChk.ForeColor = System.Drawing.Color.White;
            this.hideChildConnectedChk.Location = new System.Drawing.Point(131, 433);
            this.hideChildConnectedChk.Name = "hideChildConnectedChk";
            this.hideChildConnectedChk.Size = new System.Drawing.Size(169, 17);
            this.hideChildConnectedChk.TabIndex = 2;
            this.hideChildConnectedChk.Text = "Hide Child Connected Masters";
            this.hideChildConnectedChk.UseVisualStyleBackColor = true;
            this.hideChildConnectedChk.CheckedChanged += new System.EventHandler(this.hideChildConnectedChk_CheckedChanged);
            // 
            // hideMasterConnectedChk
            // 
            this.hideMasterConnectedChk.AutoSize = true;
            this.hideMasterConnectedChk.ForeColor = System.Drawing.Color.White;
            this.hideMasterConnectedChk.Location = new System.Drawing.Point(711, 433);
            this.hideMasterConnectedChk.Name = "hideMasterConnectedChk";
            this.hideMasterConnectedChk.Size = new System.Drawing.Size(169, 17);
            this.hideMasterConnectedChk.TabIndex = 3;
            this.hideMasterConnectedChk.Text = "Hide Master Connected Childs";
            this.hideMasterConnectedChk.UseVisualStyleBackColor = true;
            this.hideMasterConnectedChk.CheckedChanged += new System.EventHandler(this.hideMasterConnectedChk_CheckedChanged);
            // 
            // childCmb
            // 
            this.childCmb.FormattingEnabled = true;
            this.childCmb.Location = new System.Drawing.Point(569, 30);
            this.childCmb.Name = "childCmb";
            this.childCmb.Size = new System.Drawing.Size(421, 21);
            this.childCmb.TabIndex = 4;
            this.childCmb.SelectedIndexChanged += new System.EventHandler(this.childCmb_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(690, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Select A Child Category Collection";
            // 
            // connectBtn
            // 
            this.connectBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.connectBtn.Location = new System.Drawing.Point(460, 197);
            this.connectBtn.Name = "connectBtn";
            this.connectBtn.Size = new System.Drawing.Size(102, 37);
            this.connectBtn.TabIndex = 6;
            this.connectBtn.Text = "<-- Connect -->";
            this.connectBtn.UseVisualStyleBackColor = false;
            this.connectBtn.Click += new System.EventHandler(this.connectBtn_Click);
            // 
            // disconnectBtn
            // 
            this.disconnectBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.disconnectBtn.Location = new System.Drawing.Point(441, 437);
            this.disconnectBtn.Name = "disconnectBtn";
            this.disconnectBtn.Size = new System.Drawing.Size(140, 37);
            this.disconnectBtn.TabIndex = 7;
            this.disconnectBtn.Text = "--> Disconnect <--";
            this.disconnectBtn.UseVisualStyleBackColor = false;
            this.disconnectBtn.Click += new System.EventHandler(this.disconnectBtn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(167, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Master Categories";
            // 
            // connectionsLVI
            // 
            this.connectionsLVI.Location = new System.Drawing.Point(12, 480);
            this.connectionsLVI.Name = "connectionsLVI";
            this.connectionsLVI.Size = new System.Drawing.Size(978, 213);
            this.connectionsLVI.TabIndex = 9;
            this.connectionsLVI.UseCompatibleStateImageBehavior = false;
            this.connectionsLVI.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.connectionsLVI_ColumnClick);
            this.connectionsLVI.SelectedIndexChanged += new System.EventHandler(this.connectionsLVI_SelectedIndexChanged);
            // 
            // hideChildHighPriceChk
            // 
            this.hideChildHighPriceChk.AutoSize = true;
            this.hideChildHighPriceChk.ForeColor = System.Drawing.Color.White;
            this.hideChildHighPriceChk.Location = new System.Drawing.Point(711, 455);
            this.hideChildHighPriceChk.Name = "hideChildHighPriceChk";
            this.hideChildHighPriceChk.Size = new System.Drawing.Size(159, 17);
            this.hideChildHighPriceChk.TabIndex = 10;
            this.hideChildHighPriceChk.Text = "Hide High Priced Categories";
            this.hideChildHighPriceChk.UseVisualStyleBackColor = true;
            this.hideChildHighPriceChk.CheckedChanged += new System.EventHandler(this.hideChildHighPriceChk_CheckedChanged);
            // 
            // MasterCountLbl
            // 
            this.MasterCountLbl.AutoSize = true;
            this.MasterCountLbl.ForeColor = System.Drawing.Color.White;
            this.MasterCountLbl.Location = new System.Drawing.Point(366, 14);
            this.MasterCountLbl.Name = "MasterCountLbl";
            this.MasterCountLbl.Size = new System.Drawing.Size(88, 13);
            this.MasterCountLbl.TabIndex = 11;
            this.MasterCountLbl.Text = "(Displayed/Total)";
            // 
            // ChildCountLbl
            // 
            this.ChildCountLbl.AutoSize = true;
            this.ChildCountLbl.ForeColor = System.Drawing.Color.White;
            this.ChildCountLbl.Location = new System.Drawing.Point(902, 14);
            this.ChildCountLbl.Name = "ChildCountLbl";
            this.ChildCountLbl.Size = new System.Drawing.Size(88, 13);
            this.ChildCountLbl.TabIndex = 12;
            this.ChildCountLbl.Text = "(Displayed/Total)";
            // 
            // ConnectedCountLbl
            // 
            this.ConnectedCountLbl.AutoSize = true;
            this.ConnectedCountLbl.ForeColor = System.Drawing.Color.White;
            this.ConnectedCountLbl.Location = new System.Drawing.Point(902, 464);
            this.ConnectedCountLbl.Name = "ConnectedCountLbl";
            this.ConnectedCountLbl.Size = new System.Drawing.Size(88, 13);
            this.ConnectedCountLbl.TabIndex = 13;
            this.ConnectedCountLbl.Text = "(Displayed/Total)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(1002, 705);
            this.Controls.Add(this.ConnectedCountLbl);
            this.Controls.Add(this.ChildCountLbl);
            this.Controls.Add(this.MasterCountLbl);
            this.Controls.Add(this.hideChildHighPriceChk);
            this.Controls.Add(this.connectionsLVI);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.disconnectBtn);
            this.Controls.Add(this.connectBtn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.childCmb);
            this.Controls.Add(this.hideMasterConnectedChk);
            this.Controls.Add(this.hideChildConnectedChk);
            this.Controls.Add(this.ChildLVI);
            this.Controls.Add(this.MasterLVI);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "Master Category Connections";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView MasterLVI;
        private System.Windows.Forms.ListView ChildLVI;
        private System.Windows.Forms.CheckBox hideChildConnectedChk;
        private System.Windows.Forms.CheckBox hideMasterConnectedChk;
        private System.Windows.Forms.ComboBox childCmb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button connectBtn;
        private System.Windows.Forms.Button disconnectBtn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListView connectionsLVI;
        private System.Windows.Forms.CheckBox hideChildHighPriceChk;
        private System.Windows.Forms.Label MasterCountLbl;
        private System.Windows.Forms.Label ChildCountLbl;
        private System.Windows.Forms.Label ConnectedCountLbl;
    }
}

