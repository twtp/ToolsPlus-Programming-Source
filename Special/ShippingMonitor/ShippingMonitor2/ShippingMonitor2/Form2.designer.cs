namespace ShippingMonitor2
{
    partial class Form2
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
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.amEbayOrdersChk = new System.Windows.Forms.CheckBox();
            this.pmEbayOrdersChk = new System.Windows.Forms.CheckBox();
            this.toteInformationChk = new System.Windows.Forms.CheckBox();
            this.funFactsChk = new System.Windows.Forms.CheckBox();
            this.yahooOrdersChk = new System.Windows.Forms.CheckBox();
            this.otherOrdersChk = new System.Windows.Forms.CheckBox();
            this.willCallChk = new System.Windows.Forms.CheckBox();
            this.poChk = new System.Windows.Forms.CheckBox();
            this.transfersChk = new System.Windows.Forms.CheckBox();
            this.showPackersLVIChk = new System.Windows.Forms.CheckBox();
            this.restockingChk = new System.Windows.Forms.CheckBox();
            this.titleChk = new System.Windows.Forms.CheckBox();
            this.timeChk = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(87, 350);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 38);
            this.button1.TabIndex = 0;
            this.button1.Text = "Start";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(36, 12);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(204, 147);
            this.listBox1.TabIndex = 1;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // amEbayOrdersChk
            // 
            this.amEbayOrdersChk.AutoSize = true;
            this.amEbayOrdersChk.Checked = true;
            this.amEbayOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.amEbayOrdersChk.Location = new System.Drawing.Point(9, 223);
            this.amEbayOrdersChk.Name = "amEbayOrdersChk";
            this.amEbayOrdersChk.Size = new System.Drawing.Size(134, 17);
            this.amEbayOrdersChk.TabIndex = 2;
            this.amEbayOrdersChk.Text = "Show AM EBay Orders";
            this.amEbayOrdersChk.UseVisualStyleBackColor = true;
            this.amEbayOrdersChk.CheckedChanged += new System.EventHandler(this.amEbayOrdersChk_CheckedChanged);
            // 
            // pmEbayOrdersChk
            // 
            this.pmEbayOrdersChk.AutoSize = true;
            this.pmEbayOrdersChk.Checked = true;
            this.pmEbayOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.pmEbayOrdersChk.Location = new System.Drawing.Point(9, 246);
            this.pmEbayOrdersChk.Name = "pmEbayOrdersChk";
            this.pmEbayOrdersChk.Size = new System.Drawing.Size(134, 17);
            this.pmEbayOrdersChk.TabIndex = 3;
            this.pmEbayOrdersChk.Text = "Show PM EBay Orders";
            this.pmEbayOrdersChk.UseVisualStyleBackColor = true;
            this.pmEbayOrdersChk.CheckedChanged += new System.EventHandler(this.pmEbayOrdersChk_CheckedChanged);
            // 
            // toteInformationChk
            // 
            this.toteInformationChk.AutoSize = true;
            this.toteInformationChk.Checked = true;
            this.toteInformationChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toteInformationChk.Location = new System.Drawing.Point(174, 177);
            this.toteInformationChk.Name = "toteInformationChk";
            this.toteInformationChk.Size = new System.Drawing.Size(133, 17);
            this.toteInformationChk.TabIndex = 4;
            this.toteInformationChk.Text = "Show Tote Information";
            this.toteInformationChk.UseVisualStyleBackColor = true;
            this.toteInformationChk.CheckedChanged += new System.EventHandler(this.toteInformationChk_CheckedChanged);
            // 
            // funFactsChk
            // 
            this.funFactsChk.AutoSize = true;
            this.funFactsChk.Checked = true;
            this.funFactsChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.funFactsChk.Location = new System.Drawing.Point(8, 315);
            this.funFactsChk.Name = "funFactsChk";
            this.funFactsChk.Size = new System.Drawing.Size(103, 17);
            this.funFactsChk.TabIndex = 5;
            this.funFactsChk.Text = "Show Fun Facts";
            this.funFactsChk.UseVisualStyleBackColor = true;
            this.funFactsChk.CheckedChanged += new System.EventHandler(this.funFactsChk_CheckedChanged);
            // 
            // yahooOrdersChk
            // 
            this.yahooOrdersChk.AutoSize = true;
            this.yahooOrdersChk.Checked = true;
            this.yahooOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.yahooOrdersChk.Location = new System.Drawing.Point(8, 269);
            this.yahooOrdersChk.Name = "yahooOrdersChk";
            this.yahooOrdersChk.Size = new System.Drawing.Size(121, 17);
            this.yahooOrdersChk.TabIndex = 6;
            this.yahooOrdersChk.Text = "Show Yahoo Orders";
            this.yahooOrdersChk.UseVisualStyleBackColor = true;
            this.yahooOrdersChk.CheckedChanged += new System.EventHandler(this.yahooOrdersChk_CheckedChanged);
            // 
            // otherOrdersChk
            // 
            this.otherOrdersChk.AutoSize = true;
            this.otherOrdersChk.Checked = true;
            this.otherOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.otherOrdersChk.Location = new System.Drawing.Point(8, 292);
            this.otherOrdersChk.Name = "otherOrdersChk";
            this.otherOrdersChk.Size = new System.Drawing.Size(116, 17);
            this.otherOrdersChk.TabIndex = 7;
            this.otherOrdersChk.Text = "Show Other Orders";
            this.otherOrdersChk.UseVisualStyleBackColor = true;
            this.otherOrdersChk.CheckedChanged += new System.EventHandler(this.otherOrdersChk_CheckedChanged);
            // 
            // willCallChk
            // 
            this.willCallChk.AutoSize = true;
            this.willCallChk.Checked = true;
            this.willCallChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.willCallChk.Location = new System.Drawing.Point(174, 200);
            this.willCallChk.Name = "willCallChk";
            this.willCallChk.Size = new System.Drawing.Size(98, 17);
            this.willCallChk.TabIndex = 8;
            this.willCallChk.Text = "Show Will Calls";
            this.willCallChk.UseVisualStyleBackColor = true;
            this.willCallChk.CheckedChanged += new System.EventHandler(this.willCallChk_CheckedChanged);
            // 
            // poChk
            // 
            this.poChk.AutoSize = true;
            this.poChk.Checked = true;
            this.poChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.poChk.Location = new System.Drawing.Point(174, 223);
            this.poChk.Name = "poChk";
            this.poChk.Size = new System.Drawing.Size(76, 17);
            this.poChk.TabIndex = 9;
            this.poChk.Text = "Show POs";
            this.poChk.UseVisualStyleBackColor = true;
            this.poChk.CheckedChanged += new System.EventHandler(this.poChk_CheckedChanged);
            // 
            // transfersChk
            // 
            this.transfersChk.AutoSize = true;
            this.transfersChk.Checked = true;
            this.transfersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.transfersChk.Location = new System.Drawing.Point(174, 246);
            this.transfersChk.Name = "transfersChk";
            this.transfersChk.Size = new System.Drawing.Size(100, 17);
            this.transfersChk.TabIndex = 10;
            this.transfersChk.Text = "Show Transfers";
            this.transfersChk.UseVisualStyleBackColor = true;
            this.transfersChk.CheckedChanged += new System.EventHandler(this.transfersChk_CheckedChanged);
            // 
            // showPackersLVIChk
            // 
            this.showPackersLVIChk.AutoSize = true;
            this.showPackersLVIChk.Checked = true;
            this.showPackersLVIChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showPackersLVIChk.Location = new System.Drawing.Point(174, 292);
            this.showPackersLVIChk.Name = "showPackersLVIChk";
            this.showPackersLVIChk.Size = new System.Drawing.Size(95, 17);
            this.showPackersLVIChk.TabIndex = 11;
            this.showPackersLVIChk.Text = "Show Packers";
            this.showPackersLVIChk.UseVisualStyleBackColor = true;
            this.showPackersLVIChk.CheckedChanged += new System.EventHandler(this.showPackersLVIChk_CheckedChanged);
            // 
            // restockingChk
            // 
            this.restockingChk.AutoSize = true;
            this.restockingChk.Checked = true;
            this.restockingChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.restockingChk.Location = new System.Drawing.Point(174, 269);
            this.restockingChk.Name = "restockingChk";
            this.restockingChk.Size = new System.Drawing.Size(110, 17);
            this.restockingChk.TabIndex = 12;
            this.restockingChk.Text = "Show Restocking";
            this.restockingChk.UseVisualStyleBackColor = true;
            this.restockingChk.CheckedChanged += new System.EventHandler(this.restockingChk_CheckedChanged);
            // 
            // titleChk
            // 
            this.titleChk.AutoSize = true;
            this.titleChk.Checked = true;
            this.titleChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.titleChk.Location = new System.Drawing.Point(9, 177);
            this.titleChk.Name = "titleChk";
            this.titleChk.Size = new System.Drawing.Size(115, 17);
            this.titleChk.TabIndex = 13;
            this.titleChk.Text = "Show Title Caption";
            this.titleChk.UseVisualStyleBackColor = true;
            this.titleChk.CheckedChanged += new System.EventHandler(this.titleChk_CheckedChanged);
            // 
            // timeChk
            // 
            this.timeChk.AutoSize = true;
            this.timeChk.Checked = true;
            this.timeChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.timeChk.Location = new System.Drawing.Point(9, 200);
            this.timeChk.Name = "timeChk";
            this.timeChk.Size = new System.Drawing.Size(79, 17);
            this.timeChk.TabIndex = 14;
            this.timeChk.Text = "Show Time";
            this.timeChk.UseVisualStyleBackColor = true;
            this.timeChk.CheckedChanged += new System.EventHandler(this.timeChk_CheckedChanged);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 395);
            this.Controls.Add(this.timeChk);
            this.Controls.Add(this.titleChk);
            this.Controls.Add(this.restockingChk);
            this.Controls.Add(this.showPackersLVIChk);
            this.Controls.Add(this.transfersChk);
            this.Controls.Add(this.poChk);
            this.Controls.Add(this.willCallChk);
            this.Controls.Add(this.otherOrdersChk);
            this.Controls.Add(this.yahooOrdersChk);
            this.Controls.Add(this.funFactsChk);
            this.Controls.Add(this.toteInformationChk);
            this.Controls.Add(this.pmEbayOrdersChk);
            this.Controls.Add(this.amEbayOrdersChk);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form2";
            this.Text = "Shipping Monitor Selection";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.CheckBox amEbayOrdersChk;
        private System.Windows.Forms.CheckBox pmEbayOrdersChk;
        private System.Windows.Forms.CheckBox toteInformationChk;
        private System.Windows.Forms.CheckBox funFactsChk;
        private System.Windows.Forms.CheckBox yahooOrdersChk;
        private System.Windows.Forms.CheckBox otherOrdersChk;
        private System.Windows.Forms.CheckBox willCallChk;
        private System.Windows.Forms.CheckBox poChk;
        private System.Windows.Forms.CheckBox transfersChk;
        private System.Windows.Forms.CheckBox showPackersLVIChk;
        private System.Windows.Forms.CheckBox restockingChk;
        private System.Windows.Forms.CheckBox titleChk;
        private System.Windows.Forms.CheckBox timeChk;
    }
}