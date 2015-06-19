namespace CategoryConnector
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.masterLVI = new System.Windows.Forms.ListView();
            this.connectorsCmb = new System.Windows.Forms.ComboBox();
            this.childLVI = new System.Windows.Forms.ListView();
            this.connectBtn = new System.Windows.Forms.Button();
            this.connectorsLVI = new System.Windows.Forms.ListView();
            this.disconnectBtn = new System.Windows.Forms.Button();
            this.hideHighPriceChk = new System.Windows.Forms.CheckBox();
            this.hideConnectedChildsChk = new System.Windows.Forms.CheckBox();
            this.childCategoriesChk = new System.Windows.Forms.CheckBox();
            this.masterCategoriesLbl = new System.Windows.Forms.Label();
            this.childCategoriesLbl = new System.Windows.Forms.Label();
            this.currentConnectionsLbl = new System.Windows.Forms.Label();
            this.globalMastersChk = new System.Windows.Forms.CheckBox();
            this.childMastersChk = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // masterLVI
            // 
            this.masterLVI.FullRowSelect = true;
            this.masterLVI.GridLines = true;
            this.masterLVI.HideSelection = false;
            this.masterLVI.Location = new System.Drawing.Point(12, 32);
            this.masterLVI.Name = "masterLVI";
            this.masterLVI.Size = new System.Drawing.Size(497, 243);
            this.masterLVI.TabIndex = 0;
            this.masterLVI.UseCompatibleStateImageBehavior = false;
            this.masterLVI.View = System.Windows.Forms.View.Details;
            // 
            // connectorsCmb
            // 
            this.connectorsCmb.FormattingEnabled = true;
            this.connectorsCmb.Location = new System.Drawing.Point(669, 5);
            this.connectorsCmb.Name = "connectorsCmb";
            this.connectorsCmb.Size = new System.Drawing.Size(273, 21);
            this.connectorsCmb.TabIndex = 2;
            this.connectorsCmb.SelectedIndexChanged += new System.EventHandler(this.connectorsCmb_SelectedIndexChanged);
            // 
            // childLVI
            // 
            this.childLVI.CheckBoxes = true;
            this.childLVI.FullRowSelect = true;
            this.childLVI.GridLines = true;
            this.childLVI.HideSelection = false;
            this.childLVI.Location = new System.Drawing.Point(669, 32);
            this.childLVI.Name = "childLVI";
            this.childLVI.Size = new System.Drawing.Size(497, 243);
            this.childLVI.TabIndex = 3;
            this.childLVI.UseCompatibleStateImageBehavior = false;
            this.childLVI.View = System.Windows.Forms.View.Details;
            // 
            // connectBtn
            // 
            this.connectBtn.BackColor = System.Drawing.Color.Green;
            this.connectBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.connectBtn.ForeColor = System.Drawing.Color.White;
            this.connectBtn.Location = new System.Drawing.Point(516, 53);
            this.connectBtn.Name = "connectBtn";
            this.connectBtn.Size = new System.Drawing.Size(147, 46);
            this.connectBtn.TabIndex = 4;
            this.connectBtn.Text = "Connect";
            this.connectBtn.UseVisualStyleBackColor = false;
            this.connectBtn.Click += new System.EventHandler(this.connectBtn_Click);
            // 
            // connectorsLVI
            // 
            this.connectorsLVI.FullRowSelect = true;
            this.connectorsLVI.GridLines = true;
            this.connectorsLVI.HideSelection = false;
            this.connectorsLVI.Location = new System.Drawing.Point(12, 302);
            this.connectorsLVI.Name = "connectorsLVI";
            this.connectorsLVI.Size = new System.Drawing.Size(1154, 243);
            this.connectorsLVI.TabIndex = 5;
            this.connectorsLVI.UseCompatibleStateImageBehavior = false;
            this.connectorsLVI.View = System.Windows.Forms.View.Details;
            // 
            // disconnectBtn
            // 
            this.disconnectBtn.BackColor = System.Drawing.Color.Maroon;
            this.disconnectBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.disconnectBtn.ForeColor = System.Drawing.Color.White;
            this.disconnectBtn.Location = new System.Drawing.Point(517, 548);
            this.disconnectBtn.Name = "disconnectBtn";
            this.disconnectBtn.Size = new System.Drawing.Size(147, 46);
            this.disconnectBtn.TabIndex = 6;
            this.disconnectBtn.Text = "Disconnect";
            this.disconnectBtn.UseVisualStyleBackColor = false;
            this.disconnectBtn.Click += new System.EventHandler(this.disconnectBtn_Click);
            // 
            // hideHighPriceChk
            // 
            this.hideHighPriceChk.AutoSize = true;
            this.hideHighPriceChk.Location = new System.Drawing.Point(515, 105);
            this.hideHighPriceChk.Name = "hideHighPriceChk";
            this.hideHighPriceChk.Size = new System.Drawing.Size(153, 17);
            this.hideHighPriceChk.TabIndex = 7;
            this.hideHighPriceChk.Text = "Hide High Price Categories";
            this.hideHighPriceChk.UseVisualStyleBackColor = true;
            this.hideHighPriceChk.CheckedChanged += new System.EventHandler(this.hideHighPriceChk_CheckedChanged);
            // 
            // hideConnectedChildsChk
            // 
            this.hideConnectedChildsChk.AutoSize = true;
            this.hideConnectedChildsChk.Location = new System.Drawing.Point(515, 128);
            this.hideConnectedChildsChk.Name = "hideConnectedChildsChk";
            this.hideConnectedChildsChk.Size = new System.Drawing.Size(134, 17);
            this.hideConnectedChildsChk.TabIndex = 8;
            this.hideConnectedChildsChk.Text = "Hide Connected Childs";
            this.hideConnectedChildsChk.UseVisualStyleBackColor = true;
            this.hideConnectedChildsChk.CheckedChanged += new System.EventHandler(this.hideConnectedChildsChk_CheckedChanged);
            // 
            // childCategoriesChk
            // 
            this.childCategoriesChk.AutoSize = true;
            this.childCategoriesChk.Location = new System.Drawing.Point(515, 151);
            this.childCategoriesChk.Name = "childCategoriesChk";
            this.childCategoriesChk.Size = new System.Drawing.Size(126, 17);
            this.childCategoriesChk.TabIndex = 9;
            this.childCategoriesChk.Text = "Child Categories Only";
            this.childCategoriesChk.UseVisualStyleBackColor = true;
            this.childCategoriesChk.CheckedChanged += new System.EventHandler(this.childCategoriesChk_CheckedChanged);
            // 
            // masterCategoriesLbl
            // 
            this.masterCategoriesLbl.AutoSize = true;
            this.masterCategoriesLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.masterCategoriesLbl.Location = new System.Drawing.Point(12, 5);
            this.masterCategoriesLbl.Name = "masterCategoriesLbl";
            this.masterCategoriesLbl.Size = new System.Drawing.Size(178, 24);
            this.masterCategoriesLbl.TabIndex = 10;
            this.masterCategoriesLbl.Text = "Master Categories";
            // 
            // childCategoriesLbl
            // 
            this.childCategoriesLbl.AutoSize = true;
            this.childCategoriesLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.childCategoriesLbl.Location = new System.Drawing.Point(1021, 9);
            this.childCategoriesLbl.Name = "childCategoriesLbl";
            this.childCategoriesLbl.Size = new System.Drawing.Size(77, 16);
            this.childCategoriesLbl.TabIndex = 11;
            this.childCategoriesLbl.Text = "Categories:";
            // 
            // currentConnectionsLbl
            // 
            this.currentConnectionsLbl.AutoSize = true;
            this.currentConnectionsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentConnectionsLbl.Location = new System.Drawing.Point(12, 281);
            this.currentConnectionsLbl.Name = "currentConnectionsLbl";
            this.currentConnectionsLbl.Size = new System.Drawing.Size(159, 20);
            this.currentConnectionsLbl.TabIndex = 12;
            this.currentConnectionsLbl.Text = "Current Connections:";
            // 
            // globalMastersChk
            // 
            this.globalMastersChk.AutoSize = true;
            this.globalMastersChk.Enabled = false;
            this.globalMastersChk.Location = new System.Drawing.Point(515, 174);
            this.globalMastersChk.Name = "globalMastersChk";
            this.globalMastersChk.Size = new System.Drawing.Size(149, 17);
            this.globalMastersChk.TabIndex = 13;
            this.globalMastersChk.Text = "Hide Global Used Masters";
            this.globalMastersChk.UseVisualStyleBackColor = true;
            this.globalMastersChk.CheckedChanged += new System.EventHandler(this.globalMastersChk_CheckedChanged);
            // 
            // childMastersChk
            // 
            this.childMastersChk.AutoSize = true;
            this.childMastersChk.Enabled = false;
            this.childMastersChk.Location = new System.Drawing.Point(515, 197);
            this.childMastersChk.Name = "childMastersChk";
            this.childMastersChk.Size = new System.Drawing.Size(142, 17);
            this.childMastersChk.TabIndex = 14;
            this.childMastersChk.Text = "Hide Child Used Masters";
            this.childMastersChk.UseVisualStyleBackColor = true;
            this.childMastersChk.CheckedChanged += new System.EventHandler(this.childMastersChk_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(1183, 631);
            this.Controls.Add(this.childMastersChk);
            this.Controls.Add(this.globalMastersChk);
            this.Controls.Add(this.currentConnectionsLbl);
            this.Controls.Add(this.childCategoriesLbl);
            this.Controls.Add(this.masterCategoriesLbl);
            this.Controls.Add(this.childCategoriesChk);
            this.Controls.Add(this.hideConnectedChildsChk);
            this.Controls.Add(this.hideHighPriceChk);
            this.Controls.Add(this.disconnectBtn);
            this.Controls.Add(this.connectorsLVI);
            this.Controls.Add(this.connectBtn);
            this.Controls.Add(this.childLVI);
            this.Controls.Add(this.connectorsCmb);
            this.Controls.Add(this.masterLVI);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Category Connector";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView masterLVI;
        private System.Windows.Forms.ComboBox connectorsCmb;
        private System.Windows.Forms.ListView childLVI;
        private System.Windows.Forms.Button connectBtn;
        private System.Windows.Forms.ListView connectorsLVI;
        private System.Windows.Forms.Button disconnectBtn;
        private System.Windows.Forms.CheckBox hideHighPriceChk;
        private System.Windows.Forms.CheckBox hideConnectedChildsChk;
        private System.Windows.Forms.CheckBox childCategoriesChk;
        private System.Windows.Forms.Label masterCategoriesLbl;
        private System.Windows.Forms.Label childCategoriesLbl;
        private System.Windows.Forms.Label currentConnectionsLbl;
        private System.Windows.Forms.CheckBox globalMastersChk;
        private System.Windows.Forms.CheckBox childMastersChk;

    }
}

