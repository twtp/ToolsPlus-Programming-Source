namespace ShippingMonitor
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
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(81, 389);
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
            this.amEbayOrdersChk.Location = new System.Drawing.Point(36, 165);
            this.amEbayOrdersChk.Name = "amEbayOrdersChk";
            this.amEbayOrdersChk.Size = new System.Drawing.Size(134, 17);
            this.amEbayOrdersChk.TabIndex = 2;
            this.amEbayOrdersChk.Text = "Show AM EBay Orders";
            this.amEbayOrdersChk.UseVisualStyleBackColor = true;
            // 
            // pmEbayOrdersChk
            // 
            this.pmEbayOrdersChk.AutoSize = true;
            this.pmEbayOrdersChk.Checked = true;
            this.pmEbayOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.pmEbayOrdersChk.Location = new System.Drawing.Point(35, 188);
            this.pmEbayOrdersChk.Name = "pmEbayOrdersChk";
            this.pmEbayOrdersChk.Size = new System.Drawing.Size(134, 17);
            this.pmEbayOrdersChk.TabIndex = 3;
            this.pmEbayOrdersChk.Text = "Show PM EBay Orders";
            this.pmEbayOrdersChk.UseVisualStyleBackColor = true;
            // 
            // toteInformationChk
            // 
            this.toteInformationChk.AutoSize = true;
            this.toteInformationChk.Checked = true;
            this.toteInformationChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toteInformationChk.Location = new System.Drawing.Point(35, 280);
            this.toteInformationChk.Name = "toteInformationChk";
            this.toteInformationChk.Size = new System.Drawing.Size(133, 17);
            this.toteInformationChk.TabIndex = 4;
            this.toteInformationChk.Text = "Show Tote Information";
            this.toteInformationChk.UseVisualStyleBackColor = true;
            // 
            // funFactsChk
            // 
            this.funFactsChk.AutoSize = true;
            this.funFactsChk.Checked = true;
            this.funFactsChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.funFactsChk.Location = new System.Drawing.Point(35, 257);
            this.funFactsChk.Name = "funFactsChk";
            this.funFactsChk.Size = new System.Drawing.Size(103, 17);
            this.funFactsChk.TabIndex = 5;
            this.funFactsChk.Text = "Show Fun Facts";
            this.funFactsChk.UseVisualStyleBackColor = true;
            // 
            // yahooOrdersChk
            // 
            this.yahooOrdersChk.AutoSize = true;
            this.yahooOrdersChk.Checked = true;
            this.yahooOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.yahooOrdersChk.Location = new System.Drawing.Point(35, 211);
            this.yahooOrdersChk.Name = "yahooOrdersChk";
            this.yahooOrdersChk.Size = new System.Drawing.Size(121, 17);
            this.yahooOrdersChk.TabIndex = 6;
            this.yahooOrdersChk.Text = "Show Yahoo Orders";
            this.yahooOrdersChk.UseVisualStyleBackColor = true;
            // 
            // otherOrdersChk
            // 
            this.otherOrdersChk.AutoSize = true;
            this.otherOrdersChk.Checked = true;
            this.otherOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.otherOrdersChk.Location = new System.Drawing.Point(35, 234);
            this.otherOrdersChk.Name = "otherOrdersChk";
            this.otherOrdersChk.Size = new System.Drawing.Size(116, 17);
            this.otherOrdersChk.TabIndex = 7;
            this.otherOrdersChk.Text = "Show Other Orders";
            this.otherOrdersChk.UseVisualStyleBackColor = true;
            // 
            // willCallChk
            // 
            this.willCallChk.AutoSize = true;
            this.willCallChk.Checked = true;
            this.willCallChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.willCallChk.Location = new System.Drawing.Point(35, 303);
            this.willCallChk.Name = "willCallChk";
            this.willCallChk.Size = new System.Drawing.Size(98, 17);
            this.willCallChk.TabIndex = 8;
            this.willCallChk.Text = "Show Will Calls";
            this.willCallChk.UseVisualStyleBackColor = true;
            // 
            // poChk
            // 
            this.poChk.AutoSize = true;
            this.poChk.Checked = true;
            this.poChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.poChk.Location = new System.Drawing.Point(35, 326);
            this.poChk.Name = "poChk";
            this.poChk.Size = new System.Drawing.Size(76, 17);
            this.poChk.TabIndex = 9;
            this.poChk.Text = "Show POs";
            this.poChk.UseVisualStyleBackColor = true;
            // 
            // transfersChk
            // 
            this.transfersChk.AutoSize = true;
            this.transfersChk.Checked = true;
            this.transfersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.transfersChk.Location = new System.Drawing.Point(35, 349);
            this.transfersChk.Name = "transfersChk";
            this.transfersChk.Size = new System.Drawing.Size(100, 17);
            this.transfersChk.TabIndex = 10;
            this.transfersChk.Text = "Show Transfers";
            this.transfersChk.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 439);
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
    }
}