namespace EBayCategoryMapper
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
            this.biLst = new System.Windows.Forms.ListBox();
            this.eBILbl = new System.Windows.Forms.Label();
            this.eHGLbl = new System.Windows.Forms.Label();
            this.hgLst = new System.Windows.Forms.ListBox();
            this.mapBtn = new System.Windows.Forms.Button();
            this.statusLbl = new System.Windows.Forms.Label();
            this.biCountLbl = new System.Windows.Forms.Label();
            this.hgCountLbl = new System.Windows.Forms.Label();
            this.biSearchTxt = new System.Windows.Forms.TextBox();
            this.biUpBtn = new System.Windows.Forms.Button();
            this.biDownBtn = new System.Windows.Forms.Button();
            this.hgDownBtn = new System.Windows.Forms.Button();
            this.hgUpBtn = new System.Windows.Forms.Button();
            this.hgSearchTxt = new System.Windows.Forms.TextBox();
            this.refreshBtn = new System.Windows.Forms.Button();
            this.removalLVI = new System.Windows.Forms.ListView();
            this.removeMappingBtn = new System.Windows.Forms.Button();
            this.currentMapsLbl = new System.Windows.Forms.Label();
            this.removalTotLbl = new System.Windows.Forms.Label();
            this.hgDeletionBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // biLst
            // 
            this.biLst.FormattingEnabled = true;
            this.biLst.HorizontalScrollbar = true;
            this.biLst.Location = new System.Drawing.Point(10, 51);
            this.biLst.Name = "biLst";
            this.biLst.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.biLst.Size = new System.Drawing.Size(559, 303);
            this.biLst.Sorted = true;
            this.biLst.TabIndex = 0;
            // 
            // eBILbl
            // 
            this.eBILbl.AutoSize = true;
            this.eBILbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.eBILbl.Location = new System.Drawing.Point(24, 24);
            this.eBILbl.Name = "eBILbl";
            this.eBILbl.Size = new System.Drawing.Size(327, 24);
            this.eBILbl.TabIndex = 1;
            this.eBILbl.Text = "EBay Business && Industrial Categories";
            // 
            // eHGLbl
            // 
            this.eHGLbl.AutoSize = true;
            this.eHGLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.eHGLbl.Location = new System.Drawing.Point(580, 24);
            this.eHGLbl.Name = "eHGLbl";
            this.eHGLbl.Size = new System.Drawing.Size(391, 24);
            this.eHGLbl.TabIndex = 3;
            this.eHGLbl.Text = "EBay Home && Garden Unmapped Categories";
            // 
            // hgLst
            // 
            this.hgLst.FormattingEnabled = true;
            this.hgLst.HorizontalScrollbar = true;
            this.hgLst.Location = new System.Drawing.Point(575, 51);
            this.hgLst.Name = "hgLst";
            this.hgLst.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.hgLst.Size = new System.Drawing.Size(564, 303);
            this.hgLst.Sorted = true;
            this.hgLst.TabIndex = 2;
            // 
            // mapBtn
            // 
            this.mapBtn.BackColor = System.Drawing.Color.Green;
            this.mapBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.mapBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mapBtn.ForeColor = System.Drawing.Color.White;
            this.mapBtn.Location = new System.Drawing.Point(496, 354);
            this.mapBtn.Name = "mapBtn";
            this.mapBtn.Size = new System.Drawing.Size(155, 39);
            this.mapBtn.TabIndex = 4;
            this.mapBtn.Text = "Map!";
            this.mapBtn.UseVisualStyleBackColor = false;
            this.mapBtn.Click += new System.EventHandler(this.mapBtn_Click);
            // 
            // statusLbl
            // 
            this.statusLbl.AutoSize = true;
            this.statusLbl.ForeColor = System.Drawing.Color.OrangeRed;
            this.statusLbl.Location = new System.Drawing.Point(1, 3);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(43, 13);
            this.statusLbl.TabIndex = 5;
            this.statusLbl.Text = "Status: ";
            // 
            // biCountLbl
            // 
            this.biCountLbl.AutoSize = true;
            this.biCountLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.biCountLbl.Location = new System.Drawing.Point(201, 354);
            this.biCountLbl.Name = "biCountLbl";
            this.biCountLbl.Size = new System.Drawing.Size(71, 24);
            this.biCountLbl.TabIndex = 6;
            this.biCountLbl.Text = "Total: 0";
            // 
            // hgCountLbl
            // 
            this.hgCountLbl.AutoSize = true;
            this.hgCountLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hgCountLbl.Location = new System.Drawing.Point(832, 354);
            this.hgCountLbl.Name = "hgCountLbl";
            this.hgCountLbl.Size = new System.Drawing.Size(71, 24);
            this.hgCountLbl.TabIndex = 7;
            this.hgCountLbl.Text = "Total: 0";
            // 
            // biSearchTxt
            // 
            this.biSearchTxt.Location = new System.Drawing.Point(407, 29);
            this.biSearchTxt.Name = "biSearchTxt";
            this.biSearchTxt.Size = new System.Drawing.Size(100, 20);
            this.biSearchTxt.TabIndex = 8;
            this.biSearchTxt.TextChanged += new System.EventHandler(this.biSearchTxt_TextChanged);
            // 
            // biUpBtn
            // 
            this.biUpBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.biUpBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.biUpBtn.Location = new System.Drawing.Point(511, 27);
            this.biUpBtn.Name = "biUpBtn";
            this.biUpBtn.Size = new System.Drawing.Size(27, 23);
            this.biUpBtn.TabIndex = 10;
            this.biUpBtn.Text = "p";
            this.biUpBtn.UseVisualStyleBackColor = true;
            this.biUpBtn.Click += new System.EventHandler(this.biUpBtn_Click);
            // 
            // biDownBtn
            // 
            this.biDownBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.biDownBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.biDownBtn.Location = new System.Drawing.Point(539, 27);
            this.biDownBtn.Name = "biDownBtn";
            this.biDownBtn.Size = new System.Drawing.Size(27, 23);
            this.biDownBtn.TabIndex = 11;
            this.biDownBtn.Text = "q";
            this.biDownBtn.UseVisualStyleBackColor = true;
            this.biDownBtn.Click += new System.EventHandler(this.biDownBtn_Click);
            // 
            // hgDownBtn
            // 
            this.hgDownBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.hgDownBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.hgDownBtn.Location = new System.Drawing.Point(1110, 27);
            this.hgDownBtn.Name = "hgDownBtn";
            this.hgDownBtn.Size = new System.Drawing.Size(27, 23);
            this.hgDownBtn.TabIndex = 14;
            this.hgDownBtn.Text = "q";
            this.hgDownBtn.UseVisualStyleBackColor = true;
            this.hgDownBtn.Click += new System.EventHandler(this.hgDownBtn_Click);
            // 
            // hgUpBtn
            // 
            this.hgUpBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.hgUpBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.hgUpBtn.Location = new System.Drawing.Point(1082, 27);
            this.hgUpBtn.Name = "hgUpBtn";
            this.hgUpBtn.Size = new System.Drawing.Size(27, 23);
            this.hgUpBtn.TabIndex = 13;
            this.hgUpBtn.Text = "p";
            this.hgUpBtn.UseVisualStyleBackColor = true;
            this.hgUpBtn.Click += new System.EventHandler(this.hgUpBtn_Click);
            // 
            // hgSearchTxt
            // 
            this.hgSearchTxt.Location = new System.Drawing.Point(977, 29);
            this.hgSearchTxt.Name = "hgSearchTxt";
            this.hgSearchTxt.Size = new System.Drawing.Size(100, 20);
            this.hgSearchTxt.TabIndex = 12;
            this.hgSearchTxt.TextChanged += new System.EventHandler(this.hgSearchTxt_TextChanged);
            // 
            // refreshBtn
            // 
            this.refreshBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.refreshBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.refreshBtn.Location = new System.Drawing.Point(1, 25);
            this.refreshBtn.Name = "refreshBtn";
            this.refreshBtn.Size = new System.Drawing.Size(27, 23);
            this.refreshBtn.TabIndex = 15;
            this.refreshBtn.Text = "P";
            this.refreshBtn.UseVisualStyleBackColor = true;
            this.refreshBtn.Click += new System.EventHandler(this.refreshBtn_Click);
            // 
            // removalLVI
            // 
            this.removalLVI.FullRowSelect = true;
            this.removalLVI.GridLines = true;
            this.removalLVI.Location = new System.Drawing.Point(10, 406);
            this.removalLVI.Name = "removalLVI";
            this.removalLVI.Size = new System.Drawing.Size(1129, 355);
            this.removalLVI.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.removalLVI.TabIndex = 16;
            this.removalLVI.UseCompatibleStateImageBehavior = false;
            this.removalLVI.View = System.Windows.Forms.View.Details;
            // 
            // removeMappingBtn
            // 
            this.removeMappingBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.removeMappingBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.removeMappingBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.removeMappingBtn.ForeColor = System.Drawing.Color.White;
            this.removeMappingBtn.Location = new System.Drawing.Point(452, 767);
            this.removeMappingBtn.Name = "removeMappingBtn";
            this.removeMappingBtn.Size = new System.Drawing.Size(229, 43);
            this.removeMappingBtn.TabIndex = 17;
            this.removeMappingBtn.Text = "Remove Mapping";
            this.removeMappingBtn.UseVisualStyleBackColor = false;
            this.removeMappingBtn.Click += new System.EventHandler(this.removeMappingBtn_Click);
            // 
            // currentMapsLbl
            // 
            this.currentMapsLbl.AutoSize = true;
            this.currentMapsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentMapsLbl.Location = new System.Drawing.Point(11, 382);
            this.currentMapsLbl.Name = "currentMapsLbl";
            this.currentMapsLbl.Size = new System.Drawing.Size(160, 24);
            this.currentMapsLbl.TabIndex = 18;
            this.currentMapsLbl.Text = "Current Mappings";
            // 
            // removalTotLbl
            // 
            this.removalTotLbl.AutoSize = true;
            this.removalTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.removalTotLbl.Location = new System.Drawing.Point(164, 382);
            this.removalTotLbl.Name = "removalTotLbl";
            this.removalTotLbl.Size = new System.Drawing.Size(71, 24);
            this.removalTotLbl.TabIndex = 19;
            this.removalTotLbl.Text = "Total: 0";
            // 
            // hgDeletionBtn
            // 
            this.hgDeletionBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.hgDeletionBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.hgDeletionBtn.Location = new System.Drawing.Point(1008, 354);
            this.hgDeletionBtn.Name = "hgDeletionBtn";
            this.hgDeletionBtn.Size = new System.Drawing.Size(131, 30);
            this.hgDeletionBtn.TabIndex = 20;
            this.hgDeletionBtn.Text = "Delete Category(s)";
            this.hgDeletionBtn.UseVisualStyleBackColor = false;
            this.hgDeletionBtn.Click += new System.EventHandler(this.hgDeletionBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(1153, 814);
            this.Controls.Add(this.hgDeletionBtn);
            this.Controls.Add(this.removalTotLbl);
            this.Controls.Add(this.currentMapsLbl);
            this.Controls.Add(this.removeMappingBtn);
            this.Controls.Add(this.removalLVI);
            this.Controls.Add(this.refreshBtn);
            this.Controls.Add(this.hgDownBtn);
            this.Controls.Add(this.hgUpBtn);
            this.Controls.Add(this.hgSearchTxt);
            this.Controls.Add(this.biDownBtn);
            this.Controls.Add(this.biUpBtn);
            this.Controls.Add(this.biSearchTxt);
            this.Controls.Add(this.hgCountLbl);
            this.Controls.Add(this.biCountLbl);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.mapBtn);
            this.Controls.Add(this.eHGLbl);
            this.Controls.Add(this.hgLst);
            this.Controls.Add(this.eBILbl);
            this.Controls.Add(this.biLst);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "EBay Category Mapping";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox biLst;
        private System.Windows.Forms.Label eBILbl;
        private System.Windows.Forms.Label eHGLbl;
        private System.Windows.Forms.ListBox hgLst;
        private System.Windows.Forms.Button mapBtn;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.Label biCountLbl;
        private System.Windows.Forms.Label hgCountLbl;
        private System.Windows.Forms.TextBox biSearchTxt;
        private System.Windows.Forms.Button biUpBtn;
        private System.Windows.Forms.Button biDownBtn;
        private System.Windows.Forms.Button hgDownBtn;
        private System.Windows.Forms.Button hgUpBtn;
        private System.Windows.Forms.TextBox hgSearchTxt;
        private System.Windows.Forms.Button refreshBtn;
        private System.Windows.Forms.ListView removalLVI;
        private System.Windows.Forms.Button removeMappingBtn;
        private System.Windows.Forms.Label currentMapsLbl;
        private System.Windows.Forms.Label removalTotLbl;
        private System.Windows.Forms.Button hgDeletionBtn;
    }
}

