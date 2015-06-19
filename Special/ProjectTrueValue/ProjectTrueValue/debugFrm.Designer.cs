namespace ProjectTrueValue
{
    partial class debugFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(debugFrm));
            this.catalogLV = new System.Windows.Forms.ListView();
            this.loadCatalogBtn = new System.Windows.Forms.Button();
            this.catalogPathLbl = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.processCatalogBtn = new System.Windows.Forms.Button();
            this.catalogCountLbl = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.inventoryLV = new System.Windows.Forms.ListView();
            this.processedLbl = new System.Windows.Forms.Label();
            this.loadInventoryBtn = new System.Windows.Forms.Button();
            this.processInventoryBtn = new System.Windows.Forms.Button();
            this.inventoryPathLbl = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.processingToFileLbl = new System.Windows.Forms.Label();
            this.loopStatusLbl = new System.Windows.Forms.Label();
            this.itemsInTVCatalogLbl = new System.Windows.Forms.Label();
            this.getItemsBtn = new System.Windows.Forms.Button();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.shippingLV = new System.Windows.Forms.ListView();
            this.updateShippingDataBtn = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label6 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // catalogLV
            // 
            this.catalogLV.Enabled = false;
            this.catalogLV.Location = new System.Drawing.Point(25, 125);
            this.catalogLV.Name = "catalogLV";
            this.catalogLV.Size = new System.Drawing.Size(706, 291);
            this.catalogLV.TabIndex = 0;
            this.catalogLV.UseCompatibleStateImageBehavior = false;
            // 
            // loadCatalogBtn
            // 
            this.loadCatalogBtn.Enabled = false;
            this.loadCatalogBtn.Location = new System.Drawing.Point(25, 29);
            this.loadCatalogBtn.Name = "loadCatalogBtn";
            this.loadCatalogBtn.Size = new System.Drawing.Size(117, 36);
            this.loadCatalogBtn.TabIndex = 1;
            this.loadCatalogBtn.Text = "Load Catalog";
            this.loadCatalogBtn.UseVisualStyleBackColor = true;
            this.loadCatalogBtn.Click += new System.EventHandler(this.loadCatelogBtn_Click);
            // 
            // catalogPathLbl
            // 
            this.catalogPathLbl.AutoSize = true;
            this.catalogPathLbl.Enabled = false;
            this.catalogPathLbl.Location = new System.Drawing.Point(148, 41);
            this.catalogPathLbl.Name = "catalogPathLbl";
            this.catalogPathLbl.Size = new System.Drawing.Size(218, 13);
            this.catalogPathLbl.TabIndex = 2;
            this.catalogPathLbl.Text = "c:\\users\\tomwestbrook\\desktop\\catalog.xml";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // processCatalogBtn
            // 
            this.processCatalogBtn.Enabled = false;
            this.processCatalogBtn.Location = new System.Drawing.Point(25, 71);
            this.processCatalogBtn.Name = "processCatalogBtn";
            this.processCatalogBtn.Size = new System.Drawing.Size(117, 36);
            this.processCatalogBtn.TabIndex = 3;
            this.processCatalogBtn.Text = "Process Catalog";
            this.processCatalogBtn.UseVisualStyleBackColor = true;
            this.processCatalogBtn.Click += new System.EventHandler(this.processCatalogBtn_Click);
            // 
            // catalogCountLbl
            // 
            this.catalogCountLbl.AutoSize = true;
            this.catalogCountLbl.Enabled = false;
            this.catalogCountLbl.Location = new System.Drawing.Point(596, 109);
            this.catalogCountLbl.Name = "catalogCountLbl";
            this.catalogCountLbl.Size = new System.Drawing.Size(24, 13);
            this.catalogCountLbl.TabIndex = 4;
            this.catalogCountLbl.Text = "0/0";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(765, 471);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.catalogLV);
            this.tabPage1.Controls.Add(this.catalogCountLbl);
            this.tabPage1.Controls.Add(this.loadCatalogBtn);
            this.tabPage1.Controls.Add(this.processCatalogBtn);
            this.tabPage1.Controls.Add(this.catalogPathLbl);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(757, 445);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Catalog Controls";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(189, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(460, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Form is disabled in code as it should never be run again unless the TrueValue tab" +
    "le is wiped first.";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.inventoryLV);
            this.tabPage2.Controls.Add(this.processedLbl);
            this.tabPage2.Controls.Add(this.loadInventoryBtn);
            this.tabPage2.Controls.Add(this.processInventoryBtn);
            this.tabPage2.Controls.Add(this.inventoryPathLbl);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(757, 445);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Inventory Controls";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(175, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(460, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Form is disabled in code as it should never be run again unless the TrueValue tab" +
    "le is wiped first.";
            // 
            // inventoryLV
            // 
            this.inventoryLV.Enabled = false;
            this.inventoryLV.Location = new System.Drawing.Point(25, 125);
            this.inventoryLV.Name = "inventoryLV";
            this.inventoryLV.Size = new System.Drawing.Size(706, 291);
            this.inventoryLV.TabIndex = 5;
            this.inventoryLV.UseCompatibleStateImageBehavior = false;
            // 
            // processedLbl
            // 
            this.processedLbl.AutoSize = true;
            this.processedLbl.Enabled = false;
            this.processedLbl.Location = new System.Drawing.Point(596, 109);
            this.processedLbl.Name = "processedLbl";
            this.processedLbl.Size = new System.Drawing.Size(24, 13);
            this.processedLbl.TabIndex = 9;
            this.processedLbl.Text = "0/0";
            // 
            // loadInventoryBtn
            // 
            this.loadInventoryBtn.Enabled = false;
            this.loadInventoryBtn.Location = new System.Drawing.Point(25, 29);
            this.loadInventoryBtn.Name = "loadInventoryBtn";
            this.loadInventoryBtn.Size = new System.Drawing.Size(117, 36);
            this.loadInventoryBtn.TabIndex = 6;
            this.loadInventoryBtn.Text = "Load Inventory";
            this.loadInventoryBtn.UseVisualStyleBackColor = true;
            this.loadInventoryBtn.Click += new System.EventHandler(this.loadInventoryBtn_Click);
            // 
            // processInventoryBtn
            // 
            this.processInventoryBtn.Enabled = false;
            this.processInventoryBtn.Location = new System.Drawing.Point(25, 71);
            this.processInventoryBtn.Name = "processInventoryBtn";
            this.processInventoryBtn.Size = new System.Drawing.Size(117, 36);
            this.processInventoryBtn.TabIndex = 8;
            this.processInventoryBtn.Text = "Process Inventory";
            this.processInventoryBtn.UseVisualStyleBackColor = true;
            this.processInventoryBtn.Click += new System.EventHandler(this.processInventoryBtn_Click);
            // 
            // inventoryPathLbl
            // 
            this.inventoryPathLbl.AutoSize = true;
            this.inventoryPathLbl.Enabled = false;
            this.inventoryPathLbl.Location = new System.Drawing.Point(148, 41);
            this.inventoryPathLbl.Name = "inventoryPathLbl";
            this.inventoryPathLbl.Size = new System.Drawing.Size(226, 13);
            this.inventoryPathLbl.TabIndex = 7;
            this.inventoryPathLbl.Text = "c:\\users\\tomwestbrook\\desktop\\inventory.xml";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label3);
            this.tabPage3.Controls.Add(this.processingToFileLbl);
            this.tabPage3.Controls.Add(this.loopStatusLbl);
            this.tabPage3.Controls.Add(this.itemsInTVCatalogLbl);
            this.tabPage3.Controls.Add(this.getItemsBtn);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(757, 445);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Items To POINV";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(76, 185);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(604, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Form should probably never ever get run. This would THEORETICALLY import the enti" +
    "re True Value catalog into our database.\r\n";
            // 
            // processingToFileLbl
            // 
            this.processingToFileLbl.AutoSize = true;
            this.processingToFileLbl.Enabled = false;
            this.processingToFileLbl.Location = new System.Drawing.Point(176, 77);
            this.processingToFileLbl.Name = "processingToFileLbl";
            this.processingToFileLbl.Size = new System.Drawing.Size(16, 13);
            this.processingToFileLbl.TabIndex = 3;
            this.processingToFileLbl.Text = "...";
            // 
            // loopStatusLbl
            // 
            this.loopStatusLbl.AutoSize = true;
            this.loopStatusLbl.Enabled = false;
            this.loopStatusLbl.Location = new System.Drawing.Point(305, 29);
            this.loopStatusLbl.Name = "loopStatusLbl";
            this.loopStatusLbl.Size = new System.Drawing.Size(16, 13);
            this.loopStatusLbl.TabIndex = 2;
            this.loopStatusLbl.Text = "...";
            // 
            // itemsInTVCatalogLbl
            // 
            this.itemsInTVCatalogLbl.AutoSize = true;
            this.itemsInTVCatalogLbl.Enabled = false;
            this.itemsInTVCatalogLbl.Location = new System.Drawing.Point(176, 29);
            this.itemsInTVCatalogLbl.Name = "itemsInTVCatalogLbl";
            this.itemsInTVCatalogLbl.Size = new System.Drawing.Size(16, 13);
            this.itemsInTVCatalogLbl.TabIndex = 1;
            this.itemsInTVCatalogLbl.Text = "...";
            // 
            // getItemsBtn
            // 
            this.getItemsBtn.Enabled = false;
            this.getItemsBtn.Location = new System.Drawing.Point(14, 21);
            this.getItemsBtn.Name = "getItemsBtn";
            this.getItemsBtn.Size = new System.Drawing.Size(156, 28);
            this.getItemsBtn.TabIndex = 0;
            this.getItemsBtn.Text = "Get Items From TV Catalog";
            this.getItemsBtn.UseVisualStyleBackColor = true;
            this.getItemsBtn.Click += new System.EventHandler(this.getItemsBtn_Click);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.label6);
            this.tabPage4.Controls.Add(this.progressBar1);
            this.tabPage4.Controls.Add(this.label5);
            this.tabPage4.Controls.Add(this.textBox1);
            this.tabPage4.Controls.Add(this.label4);
            this.tabPage4.Controls.Add(this.shippingLV);
            this.tabPage4.Controls.Add(this.updateShippingDataBtn);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(757, 445);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Shipping Data";
            this.tabPage4.UseVisualStyleBackColor = true;
            this.tabPage4.Click += new System.EventHandler(this.tabPage4_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(18, 403);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(721, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(23, 50);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Offset:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(62, 47);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(60, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "0";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(23, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "status";
            // 
            // shippingLV
            // 
            this.shippingLV.Location = new System.Drawing.Point(195, 18);
            this.shippingLV.Name = "shippingLV";
            this.shippingLV.Size = new System.Drawing.Size(544, 361);
            this.shippingLV.TabIndex = 1;
            this.shippingLV.UseCompatibleStateImageBehavior = false;
            // 
            // updateShippingDataBtn
            // 
            this.updateShippingDataBtn.Location = new System.Drawing.Point(26, 84);
            this.updateShippingDataBtn.Name = "updateShippingDataBtn";
            this.updateShippingDataBtn.Size = new System.Drawing.Size(148, 37);
            this.updateShippingDataBtn.TabIndex = 0;
            this.updateShippingDataBtn.Text = "Update Shipping Data";
            this.updateShippingDataBtn.UseVisualStyleBackColor = true;
            this.updateShippingDataBtn.Click += new System.EventHandler(this.updateShippingDataBtn_Click);
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(23, 387);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(47, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "progress";
            // 
            // debugFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(793, 493);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "debugFrm";
            this.Text = "Debug / Testing";
            this.Load += new System.EventHandler(this.debugFrm_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView catalogLV;
        private System.Windows.Forms.Button loadCatalogBtn;
        private System.Windows.Forms.Label catalogPathLbl;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button processCatalogBtn;
        private System.Windows.Forms.Label catalogCountLbl;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ListView inventoryLV;
        private System.Windows.Forms.Label processedLbl;
        private System.Windows.Forms.Button loadInventoryBtn;
        private System.Windows.Forms.Button processInventoryBtn;
        private System.Windows.Forms.Label inventoryPathLbl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label itemsInTVCatalogLbl;
        private System.Windows.Forms.Button getItemsBtn;
        private System.Windows.Forms.Label loopStatusLbl;
        private System.Windows.Forms.Label processingToFileLbl;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.ListView shippingLV;
        private System.Windows.Forms.Button updateShippingDataBtn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
    }
}