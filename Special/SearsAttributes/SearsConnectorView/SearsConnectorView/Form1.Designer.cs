namespace SearsConnectorView
{
    partial class SearsConnectorFrm
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
            this.connectorsCmb = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.connectorLbl = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.trademarkLbl = new System.Windows.Forms.Label();
            this.multipleChoiceLbl = new System.Windows.Forms.Label();
            this.attributeNameLbl = new System.Windows.Forms.Label();
            this.customValueLbl = new System.Windows.Forms.Label();
            this.attributeTypeLbl = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.exampleItemTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.selectBtn = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.exampleLVI = new System.Windows.Forms.ListView();
            this.attributesLVI = new System.Windows.Forms.ListView();
            this.internalConnectorTxt = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // connectorsCmb
            // 
            this.connectorsCmb.FormattingEnabled = true;
            this.connectorsCmb.Location = new System.Drawing.Point(8, 8);
            this.connectorsCmb.Name = "connectorsCmb";
            this.connectorsCmb.Size = new System.Drawing.Size(721, 21);
            this.connectorsCmb.TabIndex = 0;
            this.connectorsCmb.SelectedIndexChanged += new System.EventHandler(this.connectorsCmb_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.internalConnectorTxt);
            this.groupBox1.Controls.Add(this.connectorLbl);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.listBox2);
            this.groupBox1.Controls.Add(this.trademarkLbl);
            this.groupBox1.Controls.Add(this.multipleChoiceLbl);
            this.groupBox1.Controls.Add(this.attributeNameLbl);
            this.groupBox1.Controls.Add(this.customValueLbl);
            this.groupBox1.Controls.Add(this.attributeTypeLbl);
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(357, 36);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(372, 257);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Attribute Details:";
            // 
            // connectorLbl
            // 
            this.connectorLbl.AutoSize = true;
            this.connectorLbl.Location = new System.Drawing.Point(10, 213);
            this.connectorLbl.Name = "connectorLbl";
            this.connectorLbl.Size = new System.Drawing.Size(97, 13);
            this.connectorLbl.TabIndex = 7;
            this.connectorLbl.Text = "Internal Connector:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Accepted Values:";
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(13, 113);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(348, 95);
            this.listBox2.TabIndex = 5;
            // 
            // trademarkLbl
            // 
            this.trademarkLbl.AutoSize = true;
            this.trademarkLbl.Location = new System.Drawing.Point(42, 72);
            this.trademarkLbl.Name = "trademarkLbl";
            this.trademarkLbl.Size = new System.Drawing.Size(61, 13);
            this.trademarkLbl.TabIndex = 4;
            this.trademarkLbl.Text = "Trademark:";
            // 
            // multipleChoiceLbl
            // 
            this.multipleChoiceLbl.AutoSize = true;
            this.multipleChoiceLbl.Location = new System.Drawing.Point(42, 59);
            this.multipleChoiceLbl.Name = "multipleChoiceLbl";
            this.multipleChoiceLbl.Size = new System.Drawing.Size(82, 13);
            this.multipleChoiceLbl.TabIndex = 3;
            this.multipleChoiceLbl.Text = "Multiple Choice:";
            // 
            // attributeNameLbl
            // 
            this.attributeNameLbl.AutoSize = true;
            this.attributeNameLbl.Location = new System.Drawing.Point(10, 20);
            this.attributeNameLbl.Name = "attributeNameLbl";
            this.attributeNameLbl.Size = new System.Drawing.Size(80, 13);
            this.attributeNameLbl.TabIndex = 2;
            this.attributeNameLbl.Text = "Attribute Name:";
            // 
            // customValueLbl
            // 
            this.customValueLbl.AutoSize = true;
            this.customValueLbl.Location = new System.Drawing.Point(42, 46);
            this.customValueLbl.Name = "customValueLbl";
            this.customValueLbl.Size = new System.Drawing.Size(108, 13);
            this.customValueLbl.TabIndex = 1;
            this.customValueLbl.Text = "Allows Custom Value:";
            // 
            // attributeTypeLbl
            // 
            this.attributeTypeLbl.AutoSize = true;
            this.attributeTypeLbl.Location = new System.Drawing.Point(42, 33);
            this.attributeTypeLbl.Name = "attributeTypeLbl";
            this.attributeTypeLbl.Size = new System.Drawing.Size(76, 13);
            this.attributeTypeLbl.TabIndex = 0;
            this.attributeTypeLbl.Text = "Attribute Type:";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.ForeColor = System.Drawing.Color.White;
            this.checkBox1.Location = new System.Drawing.Point(13, 19);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(153, 17);
            this.checkBox1.TabIndex = 3;
            this.checkBox1.Text = "Show Required Fields Only";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // exampleItemTxt
            // 
            this.exampleItemTxt.Location = new System.Drawing.Point(7, 32);
            this.exampleItemTxt.MaxLength = 30;
            this.exampleItemTxt.Name = "exampleItemTxt";
            this.exampleItemTxt.Size = new System.Drawing.Size(136, 20);
            this.exampleItemTxt.TabIndex = 4;
            this.exampleItemTxt.TextChanged += new System.EventHandler(this.exampleItemTxt_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(4, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Example Item:";
            // 
            // selectBtn
            // 
            this.selectBtn.ForeColor = System.Drawing.Color.LimeGreen;
            this.selectBtn.Location = new System.Drawing.Point(142, 30);
            this.selectBtn.Name = "selectBtn";
            this.selectBtn.Size = new System.Drawing.Size(27, 24);
            this.selectBtn.TabIndex = 6;
            this.selectBtn.Text = "✔";
            this.selectBtn.UseVisualStyleBackColor = true;
            this.selectBtn.Click += new System.EventHandler(this.selectBtn_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.exampleLVI);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.exampleItemTxt);
            this.groupBox2.Controls.Add(this.selectBtn);
            this.groupBox2.ForeColor = System.Drawing.Color.White;
            this.groupBox2.Location = new System.Drawing.Point(8, 299);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(343, 219);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Example Data:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.webBrowser1);
            this.groupBox3.ForeColor = System.Drawing.Color.White;
            this.groupBox3.Location = new System.Drawing.Point(358, 299);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(371, 219);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Example XML:";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(3, 16);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(365, 200);
            this.webBrowser1.TabIndex = 0;
            // 
            // exampleLVI
            // 
            this.exampleLVI.FullRowSelect = true;
            this.exampleLVI.GridLines = true;
            this.exampleLVI.Location = new System.Drawing.Point(6, 55);
            this.exampleLVI.MultiSelect = false;
            this.exampleLVI.Name = "exampleLVI";
            this.exampleLVI.Size = new System.Drawing.Size(326, 158);
            this.exampleLVI.TabIndex = 0;
            this.exampleLVI.UseCompatibleStateImageBehavior = false;
            this.exampleLVI.View = System.Windows.Forms.View.Details;
            // 
            // attributesLVI
            // 
            this.attributesLVI.FullRowSelect = true;
            this.attributesLVI.GridLines = true;
            this.attributesLVI.Location = new System.Drawing.Point(12, 42);
            this.attributesLVI.MultiSelect = false;
            this.attributesLVI.Name = "attributesLVI";
            this.attributesLVI.Size = new System.Drawing.Size(320, 207);
            this.attributesLVI.TabIndex = 9;
            this.attributesLVI.UseCompatibleStateImageBehavior = false;
            this.attributesLVI.View = System.Windows.Forms.View.Details;
            this.attributesLVI.SelectedIndexChanged += new System.EventHandler(this.attributesLVI_SelectedIndexChanged);
            // 
            // internalConnectorTxt
            // 
            this.internalConnectorTxt.Location = new System.Drawing.Point(13, 229);
            this.internalConnectorTxt.Name = "internalConnectorTxt";
            this.internalConnectorTxt.Size = new System.Drawing.Size(348, 20);
            this.internalConnectorTxt.TabIndex = 8;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.checkBox1);
            this.groupBox4.Controls.Add(this.attributesLVI);
            this.groupBox4.ForeColor = System.Drawing.Color.White;
            this.groupBox4.Location = new System.Drawing.Point(8, 36);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(343, 257);
            this.groupBox4.TabIndex = 10;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Attributes";
            // 
            // SearsConnectorFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(738, 527);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.connectorsCmb);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "SearsConnectorFrm";
            this.Text = "Sears Connectors Viewer";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox connectorsCmb;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label trademarkLbl;
        private System.Windows.Forms.Label multipleChoiceLbl;
        private System.Windows.Forms.Label attributeNameLbl;
        private System.Windows.Forms.Label customValueLbl;
        private System.Windows.Forms.Label attributeTypeLbl;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label connectorLbl;
        private System.Windows.Forms.TextBox exampleItemTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button selectBtn;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListView exampleLVI;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.ListView attributesLVI;
        private System.Windows.Forms.TextBox internalConnectorTxt;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

