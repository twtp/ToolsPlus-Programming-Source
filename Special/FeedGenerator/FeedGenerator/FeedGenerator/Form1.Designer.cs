namespace FeedGenerator
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
            this.label1 = new System.Windows.Forms.Label();
            this.feedOptionsBox = new System.Windows.Forms.GroupBox();
            this.createKeywordsChk = new System.Windows.Forms.CheckBox();
            this.sendFeedsChk = new System.Windows.Forms.CheckBox();
            this.syncImagesChk = new System.Windows.Forms.CheckBox();
            this.feedsBox = new System.Windows.Forms.GroupBox();
            this.delFeedBtn = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.newFeedBtn = new System.Windows.Forms.Button();
            this.feedsLVI = new System.Windows.Forms.ListView();
            this.statusBox = new System.Windows.Forms.GroupBox();
            this.statusLVI = new System.Windows.Forms.ListView();
            this.debugBtn = new System.Windows.Forms.Button();
            this.manualBox = new System.Windows.Forms.GroupBox();
            this.createSingleFeedBtn = new System.Windows.Forms.Button();
            this.previewBox = new System.Windows.Forms.GroupBox();
            this.exportXML = new System.Windows.Forms.Button();
            this.exportCSV = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.searchTxt = new System.Windows.Forms.TextBox();
            this.previewLVI = new System.Windows.Forms.ListView();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.editFeedBtn = new System.Windows.Forms.Button();
            this.feedOptionsBox.SuspendLayout();
            this.feedsBox.SuspendLayout();
            this.statusBox.SuspendLayout();
            this.manualBox.SuspendLayout();
            this.previewBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(33, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(168, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Feeds Generator";
            // 
            // feedOptionsBox
            // 
            this.feedOptionsBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.feedOptionsBox.Controls.Add(this.createKeywordsChk);
            this.feedOptionsBox.Controls.Add(this.sendFeedsChk);
            this.feedOptionsBox.Controls.Add(this.syncImagesChk);
            this.feedOptionsBox.Location = new System.Drawing.Point(12, 36);
            this.feedOptionsBox.Name = "feedOptionsBox";
            this.feedOptionsBox.Size = new System.Drawing.Size(214, 73);
            this.feedOptionsBox.TabIndex = 1;
            this.feedOptionsBox.TabStop = false;
            this.feedOptionsBox.Text = "Feed Options";
            // 
            // createKeywordsChk
            // 
            this.createKeywordsChk.AutoSize = true;
            this.createKeywordsChk.Location = new System.Drawing.Point(103, 23);
            this.createKeywordsChk.Name = "createKeywordsChk";
            this.createKeywordsChk.Size = new System.Drawing.Size(111, 17);
            this.createKeywordsChk.TabIndex = 2;
            this.createKeywordsChk.Text = "Create Keywords";
            this.createKeywordsChk.UseVisualStyleBackColor = true;
            // 
            // sendFeedsChk
            // 
            this.sendFeedsChk.AutoSize = true;
            this.sendFeedsChk.Location = new System.Drawing.Point(14, 23);
            this.sendFeedsChk.Name = "sendFeedsChk";
            this.sendFeedsChk.Size = new System.Drawing.Size(85, 17);
            this.sendFeedsChk.TabIndex = 1;
            this.sendFeedsChk.Text = "Send Feeds";
            this.sendFeedsChk.UseVisualStyleBackColor = true;
            // 
            // syncImagesChk
            // 
            this.syncImagesChk.AutoSize = true;
            this.syncImagesChk.Location = new System.Drawing.Point(13, 46);
            this.syncImagesChk.Name = "syncImagesChk";
            this.syncImagesChk.Size = new System.Drawing.Size(88, 17);
            this.syncImagesChk.TabIndex = 0;
            this.syncImagesChk.Text = "Sync Images";
            this.syncImagesChk.UseVisualStyleBackColor = true;
            // 
            // feedsBox
            // 
            this.feedsBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.feedsBox.Controls.Add(this.editFeedBtn);
            this.feedsBox.Controls.Add(this.delFeedBtn);
            this.feedsBox.Controls.Add(this.label3);
            this.feedsBox.Controls.Add(this.newFeedBtn);
            this.feedsBox.Controls.Add(this.feedsLVI);
            this.feedsBox.Location = new System.Drawing.Point(12, 115);
            this.feedsBox.Name = "feedsBox";
            this.feedsBox.Size = new System.Drawing.Size(214, 376);
            this.feedsBox.TabIndex = 2;
            this.feedsBox.TabStop = false;
            this.feedsBox.Text = "Feeds";
            // 
            // delFeedBtn
            // 
            this.delFeedBtn.Location = new System.Drawing.Point(171, 9);
            this.delFeedBtn.Name = "delFeedBtn";
            this.delFeedBtn.Size = new System.Drawing.Size(37, 23);
            this.delFeedBtn.TabIndex = 5;
            this.delFeedBtn.Text = "del";
            this.delFeedBtn.UseVisualStyleBackColor = true;
            this.delFeedBtn.Click += new System.EventHandler(this.delFeedBtn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(11, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(133, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "Double Click To Enable/Disable";
            // 
            // newFeedBtn
            // 
            this.newFeedBtn.Location = new System.Drawing.Point(82, 9);
            this.newFeedBtn.Name = "newFeedBtn";
            this.newFeedBtn.Size = new System.Drawing.Size(37, 23);
            this.newFeedBtn.TabIndex = 4;
            this.newFeedBtn.Text = "new";
            this.newFeedBtn.UseVisualStyleBackColor = true;
            this.newFeedBtn.Click += new System.EventHandler(this.newFeedBtn_Click);
            // 
            // feedsLVI
            // 
            this.feedsLVI.Location = new System.Drawing.Point(6, 50);
            this.feedsLVI.Name = "feedsLVI";
            this.feedsLVI.Size = new System.Drawing.Size(203, 320);
            this.feedsLVI.TabIndex = 3;
            this.feedsLVI.UseCompatibleStateImageBehavior = false;
            this.feedsLVI.SelectedIndexChanged += new System.EventHandler(this.feedsLVI_SelectedIndexChanged);
            // 
            // statusBox
            // 
            this.statusBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.statusBox.Controls.Add(this.statusLVI);
            this.statusBox.Location = new System.Drawing.Point(447, 9);
            this.statusBox.Name = "statusBox";
            this.statusBox.Size = new System.Drawing.Size(460, 163);
            this.statusBox.TabIndex = 5;
            this.statusBox.TabStop = false;
            this.statusBox.Text = "Status";
            // 
            // statusLVI
            // 
            this.statusLVI.Location = new System.Drawing.Point(6, 19);
            this.statusLVI.Name = "statusLVI";
            this.statusLVI.Size = new System.Drawing.Size(451, 138);
            this.statusLVI.TabIndex = 4;
            this.statusLVI.UseCompatibleStateImageBehavior = false;
            // 
            // debugBtn
            // 
            this.debugBtn.Location = new System.Drawing.Point(963, 177);
            this.debugBtn.Name = "debugBtn";
            this.debugBtn.Size = new System.Drawing.Size(96, 31);
            this.debugBtn.TabIndex = 5;
            this.debugBtn.Text = "debugger";
            this.debugBtn.UseVisualStyleBackColor = true;
            this.debugBtn.Visible = false;
            this.debugBtn.Click += new System.EventHandler(this.debugBtn_Click);
            // 
            // manualBox
            // 
            this.manualBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.manualBox.Controls.Add(this.createSingleFeedBtn);
            this.manualBox.Location = new System.Drawing.Point(238, 9);
            this.manualBox.Name = "manualBox";
            this.manualBox.Size = new System.Drawing.Size(195, 163);
            this.manualBox.TabIndex = 6;
            this.manualBox.TabStop = false;
            this.manualBox.Text = "Manual Tasks";
            // 
            // createSingleFeedBtn
            // 
            this.createSingleFeedBtn.Location = new System.Drawing.Point(13, 27);
            this.createSingleFeedBtn.Name = "createSingleFeedBtn";
            this.createSingleFeedBtn.Size = new System.Drawing.Size(168, 23);
            this.createSingleFeedBtn.TabIndex = 0;
            this.createSingleFeedBtn.Text = "Create Feed From Selected";
            this.createSingleFeedBtn.UseVisualStyleBackColor = true;
            this.createSingleFeedBtn.Click += new System.EventHandler(this.createSingleFeedBtn_Click);
            // 
            // previewBox
            // 
            this.previewBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.previewBox.Controls.Add(this.exportXML);
            this.previewBox.Controls.Add(this.exportCSV);
            this.previewBox.Controls.Add(this.label2);
            this.previewBox.Controls.Add(this.searchTxt);
            this.previewBox.Controls.Add(this.previewLVI);
            this.previewBox.Location = new System.Drawing.Point(238, 184);
            this.previewBox.Name = "previewBox";
            this.previewBox.Size = new System.Drawing.Size(669, 307);
            this.previewBox.TabIndex = 7;
            this.previewBox.TabStop = false;
            this.previewBox.Text = "Feed Preview / Manipulator";
            // 
            // exportXML
            // 
            this.exportXML.Location = new System.Drawing.Point(111, 21);
            this.exportXML.Name = "exportXML";
            this.exportXML.Size = new System.Drawing.Size(84, 22);
            this.exportXML.TabIndex = 4;
            this.exportXML.Text = "Export XML";
            this.exportXML.UseVisualStyleBackColor = true;
            this.exportXML.Click += new System.EventHandler(this.exportXML_Click);
            // 
            // exportCSV
            // 
            this.exportCSV.Location = new System.Drawing.Point(16, 21);
            this.exportCSV.Name = "exportCSV";
            this.exportCSV.Size = new System.Drawing.Size(84, 22);
            this.exportCSV.TabIndex = 3;
            this.exportCSV.Text = "Export CSV";
            this.exportCSV.UseVisualStyleBackColor = true;
            this.exportCSV.Click += new System.EventHandler(this.exportCSV_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(576, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = " search";
            // 
            // searchTxt
            // 
            this.searchTxt.Location = new System.Drawing.Point(458, 23);
            this.searchTxt.Name = "searchTxt";
            this.searchTxt.Size = new System.Drawing.Size(112, 22);
            this.searchTxt.TabIndex = 1;
            // 
            // previewLVI
            // 
            this.previewLVI.AllowColumnReorder = true;
            this.previewLVI.Location = new System.Drawing.Point(6, 49);
            this.previewLVI.MultiSelect = false;
            this.previewLVI.Name = "previewLVI";
            this.previewLVI.Size = new System.Drawing.Size(657, 252);
            this.previewLVI.TabIndex = 0;
            this.previewLVI.UseCompatibleStateImageBehavior = false;
            this.previewLVI.ColumnReordered += new System.Windows.Forms.ColumnReorderedEventHandler(this.previewLVI_ColumnReordered);
            this.previewLVI.SelectedIndexChanged += new System.EventHandler(this.previewLVI_SelectedIndexChanged);
            this.previewLVI.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.previewLVI_MouseDoubleClick);
            // 
            // editFeedBtn
            // 
            this.editFeedBtn.Location = new System.Drawing.Point(125, 9);
            this.editFeedBtn.Name = "editFeedBtn";
            this.editFeedBtn.Size = new System.Drawing.Size(40, 23);
            this.editFeedBtn.TabIndex = 8;
            this.editFeedBtn.Text = "edit";
            this.editFeedBtn.UseVisualStyleBackColor = true;
            this.editFeedBtn.Click += new System.EventHandler(this.editFeedBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gray;
            this.ClientSize = new System.Drawing.Size(919, 500);
            this.Controls.Add(this.debugBtn);
            this.Controls.Add(this.previewBox);
            this.Controls.Add(this.manualBox);
            this.Controls.Add(this.statusBox);
            this.Controls.Add(this.feedsBox);
            this.Controls.Add(this.feedOptionsBox);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI Symbol", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Feeds Generator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.feedOptionsBox.ResumeLayout(false);
            this.feedOptionsBox.PerformLayout();
            this.feedsBox.ResumeLayout(false);
            this.feedsBox.PerformLayout();
            this.statusBox.ResumeLayout(false);
            this.manualBox.ResumeLayout(false);
            this.previewBox.ResumeLayout(false);
            this.previewBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox feedOptionsBox;
        private System.Windows.Forms.CheckBox sendFeedsChk;
        private System.Windows.Forms.CheckBox syncImagesChk;
        private System.Windows.Forms.GroupBox feedsBox;
        private System.Windows.Forms.ListView feedsLVI;
        private System.Windows.Forms.GroupBox statusBox;
        private System.Windows.Forms.GroupBox manualBox;
        private System.Windows.Forms.Button createSingleFeedBtn;
        private System.Windows.Forms.GroupBox previewBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox searchTxt;
        private System.Windows.Forms.ListView statusLVI;
        private System.Windows.Forms.CheckBox createKeywordsChk;
        private System.Windows.Forms.Button newFeedBtn;
        private System.Windows.Forms.Button debugBtn;
        private System.Windows.Forms.Button exportCSV;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button exportXML;
        public System.Windows.Forms.ListView previewLVI;
        private System.Windows.Forms.Button delFeedBtn;
        private System.Windows.Forms.Button editFeedBtn;
    }
}

