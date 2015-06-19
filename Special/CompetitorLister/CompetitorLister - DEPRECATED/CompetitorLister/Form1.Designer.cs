namespace CompetitorLister
{
    partial class competitorFrm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(competitorFrm));
            this.competitorLst = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.updateLbl = new System.Windows.Forms.Label();
            this.applyBtn = new System.Windows.Forms.Button();
            this.whitelistChk = new System.Windows.Forms.CheckBox();
            this.blacklistChk = new System.Windows.Forms.CheckBox();
            this.enabledChk = new System.Windows.Forms.CheckBox();
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.competitorTxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // competitorLst
            // 
            this.competitorLst.FormattingEnabled = true;
            this.competitorLst.Location = new System.Drawing.Point(2, 36);
            this.competitorLst.Name = "competitorLst";
            this.competitorLst.Size = new System.Drawing.Size(188, 264);
            this.competitorLst.TabIndex = 0;
            this.competitorLst.SelectedIndexChanged += new System.EventHandler(this.competitorLst_SelectedIndexChanged);
            this.competitorLst.DoubleClick += new System.EventHandler(this.competitorLst_DoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.updateLbl);
            this.groupBox1.Controls.Add(this.applyBtn);
            this.groupBox1.Controls.Add(this.whitelistChk);
            this.groupBox1.Controls.Add(this.blacklistChk);
            this.groupBox1.Controls.Add(this.enabledChk);
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(195, 32);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(104, 180);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Actions";
            // 
            // updateLbl
            // 
            this.updateLbl.AutoSize = true;
            this.updateLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.updateLbl.ForeColor = System.Drawing.Color.Yellow;
            this.updateLbl.Location = new System.Drawing.Point(15, 161);
            this.updateLbl.Name = "updateLbl";
            this.updateLbl.Size = new System.Drawing.Size(76, 13);
            this.updateLbl.TabIndex = 4;
            this.updateLbl.Text = "Updated OK";
            this.updateLbl.Visible = false;
            // 
            // applyBtn
            // 
            this.applyBtn.ForeColor = System.Drawing.Color.Black;
            this.applyBtn.Location = new System.Drawing.Point(16, 116);
            this.applyBtn.Name = "applyBtn";
            this.applyBtn.Size = new System.Drawing.Size(75, 42);
            this.applyBtn.TabIndex = 3;
            this.applyBtn.Text = "Apply";
            this.applyBtn.UseVisualStyleBackColor = true;
            this.applyBtn.Click += new System.EventHandler(this.applyBtn_Click);
            // 
            // whitelistChk
            // 
            this.whitelistChk.AutoSize = true;
            this.whitelistChk.Location = new System.Drawing.Point(16, 83);
            this.whitelistChk.Name = "whitelistChk";
            this.whitelistChk.Size = new System.Drawing.Size(78, 17);
            this.whitelistChk.TabIndex = 2;
            this.whitelistChk.Text = "Whitelisted";
            this.whitelistChk.UseVisualStyleBackColor = true;
            // 
            // blacklistChk
            // 
            this.blacklistChk.AutoSize = true;
            this.blacklistChk.Location = new System.Drawing.Point(16, 55);
            this.blacklistChk.Name = "blacklistChk";
            this.blacklistChk.Size = new System.Drawing.Size(77, 17);
            this.blacklistChk.TabIndex = 1;
            this.blacklistChk.Text = "Blacklisted";
            this.blacklistChk.UseVisualStyleBackColor = true;
            // 
            // enabledChk
            // 
            this.enabledChk.AutoSize = true;
            this.enabledChk.Location = new System.Drawing.Point(16, 27);
            this.enabledChk.Name = "enabledChk";
            this.enabledChk.Size = new System.Drawing.Size(65, 17);
            this.enabledChk.TabIndex = 0;
            this.enabledChk.Text = "Enabled";
            this.enabledChk.UseVisualStyleBackColor = true;
            // 
            // updateTmr
            // 
            this.updateTmr.Interval = 3000;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::CompetitorLister.Properties.Resources.wiser2;
            this.pictureBox1.Location = new System.Drawing.Point(196, 215);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(101, 92);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // competitorTxt
            // 
            this.competitorTxt.Location = new System.Drawing.Point(77, 10);
            this.competitorTxt.Name = "competitorTxt";
            this.competitorTxt.Size = new System.Drawing.Size(220, 20);
            this.competitorTxt.TabIndex = 3;
            this.competitorTxt.TextChanged += new System.EventHandler(this.competitorTxt_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Competitor:";
            // 
            // competitorFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(304, 305);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.competitorTxt);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.competitorLst);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "competitorFrm";
            this.Text = "Competitors";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox competitorLst;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button applyBtn;
        private System.Windows.Forms.CheckBox whitelistChk;
        private System.Windows.Forms.CheckBox blacklistChk;
        private System.Windows.Forms.CheckBox enabledChk;
        private System.Windows.Forms.Label updateLbl;
        private System.Windows.Forms.Timer updateTmr;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox competitorTxt;
        private System.Windows.Forms.Label label1;

    }
}

