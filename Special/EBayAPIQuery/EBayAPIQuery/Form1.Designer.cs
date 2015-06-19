namespace EBayAPIQuery
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
            this.ebayEndPointTxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.contentHeaderTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ebayCallNameTxt = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.xmlRequestTxt = new System.Windows.Forms.TextBox();
            this.responseTxt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ebayEndPointTxt
            // 
            this.ebayEndPointTxt.Location = new System.Drawing.Point(95, 12);
            this.ebayEndPointTxt.Name = "ebayEndPointTxt";
            this.ebayEndPointTxt.Size = new System.Drawing.Size(272, 20);
            this.ebayEndPointTxt.TabIndex = 0;
            this.ebayEndPointTxt.Text = "https://api.ebay.com/ws/api.dll";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "EBay Endpoint";
            // 
            // contentHeaderTxt
            // 
            this.contentHeaderTxt.Location = new System.Drawing.Point(15, 56);
            this.contentHeaderTxt.Multiline = true;
            this.contentHeaderTxt.Name = "contentHeaderTxt";
            this.contentHeaderTxt.Size = new System.Drawing.Size(352, 150);
            this.contentHeaderTxt.TabIndex = 2;
            this.contentHeaderTxt.Text = resources.GetString("contentHeaderTxt.Text");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Content Header";
            // 
            // ebayCallNameTxt
            // 
            this.ebayCallNameTxt.Location = new System.Drawing.Point(104, 216);
            this.ebayCallNameTxt.Name = "ebayCallNameTxt";
            this.ebayCallNameTxt.Size = new System.Drawing.Size(263, 20);
            this.ebayCallNameTxt.TabIndex = 4;
            this.ebayCallNameTxt.Text = "GetItem";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 219);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "EBay Call Name:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(373, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "XML Request:";
            // 
            // xmlRequestTxt
            // 
            this.xmlRequestTxt.Location = new System.Drawing.Point(376, 26);
            this.xmlRequestTxt.Multiline = true;
            this.xmlRequestTxt.Name = "xmlRequestTxt";
            this.xmlRequestTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.xmlRequestTxt.Size = new System.Drawing.Size(511, 210);
            this.xmlRequestTxt.TabIndex = 7;
            this.xmlRequestTxt.Text = resources.GetString("xmlRequestTxt.Text");
            // 
            // responseTxt
            // 
            this.responseTxt.Location = new System.Drawing.Point(15, 278);
            this.responseTxt.Multiline = true;
            this.responseTxt.Name = "responseTxt";
            this.responseTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.responseTxt.Size = new System.Drawing.Size(871, 307);
            this.responseTxt.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 262);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Response:";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.button1.Location = new System.Drawing.Point(591, 242);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(123, 32);
            this.button1.TabIndex = 10;
            this.button1.Text = "Query";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(898, 597);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.responseTxt);
            this.Controls.Add(this.xmlRequestTxt);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ebayCallNameTxt);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.contentHeaderTxt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ebayEndPointTxt);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "EBay API Query";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox ebayEndPointTxt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox contentHeaderTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox ebayCallNameTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox xmlRequestTxt;
        private System.Windows.Forms.TextBox responseTxt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
    }
}

