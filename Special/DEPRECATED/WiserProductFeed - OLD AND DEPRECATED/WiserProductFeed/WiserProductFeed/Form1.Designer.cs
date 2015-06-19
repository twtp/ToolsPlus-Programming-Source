namespace WiserProductFeed
{
    partial class WiserMainFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WiserMainFrm));
            this.feedLVI = new System.Windows.Forms.ListView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.statusLbl = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.missingASINChk = new System.Windows.Forms.CheckBox();
            this.delayLbl = new System.Windows.Forms.Label();
            this.delayTrk = new System.Windows.Forms.TrackBar();
            this.useBlacklistChk = new System.Windows.Forms.CheckBox();
            this.delayWriteChk = new System.Windows.Forms.CheckBox();
            this.button6 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.incommingLVI = new System.Windows.Forms.ListView();
            this.lastUpdateLbl = new System.Windows.Forms.Label();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.delayTrk)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // feedLVI
            // 
            this.feedLVI.FullRowSelect = true;
            this.feedLVI.GridLines = true;
            this.feedLVI.Location = new System.Drawing.Point(-2, 0);
            this.feedLVI.MultiSelect = false;
            this.feedLVI.Name = "feedLVI";
            this.feedLVI.Size = new System.Drawing.Size(1001, 182);
            this.feedLVI.TabIndex = 0;
            this.feedLVI.UseCompatibleStateImageBehavior = false;
            this.feedLVI.View = System.Windows.Forms.View.Details;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(18, 17);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(132, 31);
            this.button1.TabIndex = 1;
            this.button1.Text = "Create Feed";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(18, 48);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(132, 31);
            this.button2.TabIndex = 2;
            this.button2.Text = "Upload Feed";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // statusLbl
            // 
            this.statusLbl.Location = new System.Drawing.Point(2, 5);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(573, 18);
            this.statusLbl.TabIndex = 4;
            this.statusLbl.Text = "Status:";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(18, 141);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(132, 31);
            this.button3.TabIndex = 6;
            this.button3.Text = "Store Competitor Info";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(18, 79);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(132, 31);
            this.button4.TabIndex = 5;
            this.button4.Text = "Download Return Feed";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.missingASINChk);
            this.groupBox1.Controls.Add(this.delayLbl);
            this.groupBox1.Controls.Add(this.delayTrk);
            this.groupBox1.Controls.Add(this.useBlacklistChk);
            this.groupBox1.Controls.Add(this.delayWriteChk);
            this.groupBox1.Controls.Add(this.button6);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Location = new System.Drawing.Point(5, 275);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(293, 179);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Manual Controls";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // missingASINChk
            // 
            this.missingASINChk.AutoSize = true;
            this.missingASINChk.Location = new System.Drawing.Point(173, 116);
            this.missingASINChk.Name = "missingASINChk";
            this.missingASINChk.Size = new System.Drawing.Size(114, 17);
            this.missingASINChk.TabIndex = 15;
            this.missingASINChk.Text = "Get Missing ASINs";
            this.missingASINChk.UseVisualStyleBackColor = true;
            // 
            // delayLbl
            // 
            this.delayLbl.Location = new System.Drawing.Point(167, 39);
            this.delayLbl.Name = "delayLbl";
            this.delayLbl.Size = new System.Drawing.Size(120, 23);
            this.delayLbl.TabIndex = 14;
            this.delayLbl.Text = "Delay Setting:";
            this.delayLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // delayTrk
            // 
            this.delayTrk.Location = new System.Drawing.Point(173, 19);
            this.delayTrk.Maximum = 1000;
            this.delayTrk.Minimum = 1;
            this.delayTrk.Name = "delayTrk";
            this.delayTrk.Size = new System.Drawing.Size(104, 45);
            this.delayTrk.TabIndex = 13;
            this.delayTrk.Value = 92;
            this.delayTrk.Scroll += new System.EventHandler(this.delayTrk_Scroll);
            // 
            // useBlacklistChk
            // 
            this.useBlacklistChk.AutoSize = true;
            this.useBlacklistChk.Checked = true;
            this.useBlacklistChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.useBlacklistChk.Location = new System.Drawing.Point(173, 93);
            this.useBlacklistChk.Name = "useBlacklistChk";
            this.useBlacklistChk.Size = new System.Drawing.Size(112, 17);
            this.useBlacklistChk.TabIndex = 12;
            this.useBlacklistChk.Text = "Use Blacklist Filter";
            this.useBlacklistChk.UseVisualStyleBackColor = true;
            // 
            // delayWriteChk
            // 
            this.delayWriteChk.AutoSize = true;
            this.delayWriteChk.Checked = true;
            this.delayWriteChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.delayWriteChk.Location = new System.Drawing.Point(173, 70);
            this.delayWriteChk.Name = "delayWriteChk";
            this.delayWriteChk.Size = new System.Drawing.Size(81, 17);
            this.delayWriteChk.TabIndex = 11;
            this.delayWriteChk.Text = "Delay Write";
            this.delayWriteChk.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(18, 110);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(132, 31);
            this.button6.TabIndex = 10;
            this.button6.Text = "Display Incomming Feed";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button8);
            this.groupBox2.Controls.Add(this.button7);
            this.groupBox2.Controls.Add(this.button5);
            this.groupBox2.Location = new System.Drawing.Point(304, 276);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(169, 178);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Automation Control";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(19, 29);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(132, 31);
            this.button5.TabIndex = 2;
            this.button5.Text = "Full Auto Cycle";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(3, 60);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1007, 211);
            this.tabControl1.TabIndex = 9;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.feedLVI);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(999, 185);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Outgoing Feed";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.incommingLVI);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(999, 185);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Incomming Feed";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // incommingLVI
            // 
            this.incommingLVI.FullRowSelect = true;
            this.incommingLVI.GridLines = true;
            this.incommingLVI.Location = new System.Drawing.Point(-1, -1);
            this.incommingLVI.MultiSelect = false;
            this.incommingLVI.Name = "incommingLVI";
            this.incommingLVI.Size = new System.Drawing.Size(1001, 182);
            this.incommingLVI.TabIndex = 1;
            this.incommingLVI.UseCompatibleStateImageBehavior = false;
            this.incommingLVI.View = System.Windows.Forms.View.Details;
            // 
            // lastUpdateLbl
            // 
            this.lastUpdateLbl.AutoSize = true;
            this.lastUpdateLbl.Location = new System.Drawing.Point(3, 23);
            this.lastUpdateLbl.Name = "lastUpdateLbl";
            this.lastUpdateLbl.Size = new System.Drawing.Size(37, 13);
            this.lastUpdateLbl.TabIndex = 11;
            this.lastUpdateLbl.Text = "Detail:";
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(19, 96);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(132, 29);
            this.button7.TabIndex = 3;
            this.button7.Text = "format CSV";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(19, 131);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(132, 29);
            this.button8.TabIndex = 4;
            this.button8.Text = "bulk Insert";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // WiserMainFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 457);
            this.Controls.Add(this.lastUpdateLbl);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusLbl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "WiserMainFrm";
            this.Text = "Wiser Product Feed";
            this.Load += new System.EventHandler(this.WiserMainFrm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.delayTrk)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView feedLVI;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ListView incommingLVI;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.CheckBox delayWriteChk;
        private System.Windows.Forms.Label lastUpdateLbl;
        private System.Windows.Forms.CheckBox useBlacklistChk;
        private System.Windows.Forms.TrackBar delayTrk;
        private System.Windows.Forms.Label delayLbl;
        private System.Windows.Forms.CheckBox missingASINChk;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
    }
}

