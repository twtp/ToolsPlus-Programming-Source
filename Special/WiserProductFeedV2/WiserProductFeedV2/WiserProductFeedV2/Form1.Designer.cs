namespace WiserProductFeedV2
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.statusLbl = new System.Windows.Forms.Label();
            this.uploadBtn = new System.Windows.Forms.Button();
            this.downloadBtn = new System.Windows.Forms.Button();
            this.formatBtn = new System.Windows.Forms.Button();
            this.bulkBtn = new System.Windows.Forms.Button();
            this.feedLVI = new System.Windows.Forms.ListView();
            this.createBtn = new System.Windows.Forms.Button();
            this.createLbl = new System.Windows.Forms.Label();
            this.uploadLbl = new System.Windows.Forms.Label();
            this.downloadLbl = new System.Windows.Forms.Label();
            this.formatLbl = new System.Windows.Forms.Label();
            this.insertLbl = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.updateLbl = new System.Windows.Forms.Label();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.endTmr = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // statusLbl
            // 
            this.statusLbl.BackColor = System.Drawing.Color.Transparent;
            this.statusLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusLbl.Location = new System.Drawing.Point(1, 48);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(365, 20);
            this.statusLbl.TabIndex = 1;
            this.statusLbl.Text = "Status:";
            // 
            // uploadBtn
            // 
            this.uploadBtn.Location = new System.Drawing.Point(59, 88);
            this.uploadBtn.Name = "uploadBtn";
            this.uploadBtn.Size = new System.Drawing.Size(47, 26);
            this.uploadBtn.TabIndex = 2;
            this.uploadBtn.Text = "upload";
            this.uploadBtn.UseVisualStyleBackColor = true;
            this.uploadBtn.Click += new System.EventHandler(this.uploadBtn_Click);
            // 
            // downloadBtn
            // 
            this.downloadBtn.Location = new System.Drawing.Point(112, 88);
            this.downloadBtn.Name = "downloadBtn";
            this.downloadBtn.Size = new System.Drawing.Size(66, 26);
            this.downloadBtn.TabIndex = 3;
            this.downloadBtn.Text = "download";
            this.downloadBtn.UseVisualStyleBackColor = true;
            this.downloadBtn.Click += new System.EventHandler(this.downloadBtn_Click);
            // 
            // formatBtn
            // 
            this.formatBtn.Location = new System.Drawing.Point(184, 88);
            this.formatBtn.Name = "formatBtn";
            this.formatBtn.Size = new System.Drawing.Size(47, 26);
            this.formatBtn.TabIndex = 4;
            this.formatBtn.Text = "format";
            this.formatBtn.UseVisualStyleBackColor = true;
            this.formatBtn.Click += new System.EventHandler(this.formatBtn_Click);
            // 
            // bulkBtn
            // 
            this.bulkBtn.Location = new System.Drawing.Point(237, 88);
            this.bulkBtn.Name = "bulkBtn";
            this.bulkBtn.Size = new System.Drawing.Size(56, 26);
            this.bulkBtn.TabIndex = 5;
            this.bulkBtn.Text = "bulk sql";
            this.bulkBtn.UseVisualStyleBackColor = true;
            this.bulkBtn.Click += new System.EventHandler(this.bulkBtn_Click);
            // 
            // feedLVI
            // 
            this.feedLVI.FullRowSelect = true;
            this.feedLVI.GridLines = true;
            this.feedLVI.Location = new System.Drawing.Point(38, 118);
            this.feedLVI.Name = "feedLVI";
            this.feedLVI.Size = new System.Drawing.Size(165, 19);
            this.feedLVI.TabIndex = 6;
            this.feedLVI.UseCompatibleStateImageBehavior = false;
            this.feedLVI.View = System.Windows.Forms.View.Details;
            // 
            // createBtn
            // 
            this.createBtn.Location = new System.Drawing.Point(6, 88);
            this.createBtn.Name = "createBtn";
            this.createBtn.Size = new System.Drawing.Size(47, 26);
            this.createBtn.TabIndex = 7;
            this.createBtn.Text = "create";
            this.createBtn.UseVisualStyleBackColor = true;
            this.createBtn.Click += new System.EventHandler(this.createBtn_Click);
            // 
            // createLbl
            // 
            this.createLbl.AutoSize = true;
            this.createLbl.Location = new System.Drawing.Point(0, 35);
            this.createLbl.Name = "createLbl";
            this.createLbl.Size = new System.Drawing.Size(38, 13);
            this.createLbl.TabIndex = 8;
            this.createLbl.Text = "Create";
            // 
            // uploadLbl
            // 
            this.uploadLbl.AutoSize = true;
            this.uploadLbl.Location = new System.Drawing.Point(55, 35);
            this.uploadLbl.Name = "uploadLbl";
            this.uploadLbl.Size = new System.Drawing.Size(41, 13);
            this.uploadLbl.TabIndex = 9;
            this.uploadLbl.Text = "Upload";
            // 
            // downloadLbl
            // 
            this.downloadLbl.AutoSize = true;
            this.downloadLbl.Location = new System.Drawing.Point(119, 35);
            this.downloadLbl.Name = "downloadLbl";
            this.downloadLbl.Size = new System.Drawing.Size(55, 13);
            this.downloadLbl.TabIndex = 10;
            this.downloadLbl.Text = "Download";
            // 
            // formatLbl
            // 
            this.formatLbl.AutoSize = true;
            this.formatLbl.Location = new System.Drawing.Point(199, 35);
            this.formatLbl.Name = "formatLbl";
            this.formatLbl.Size = new System.Drawing.Size(45, 13);
            this.formatLbl.TabIndex = 11;
            this.formatLbl.Text = "Format  ";
            // 
            // insertLbl
            // 
            this.insertLbl.AutoSize = true;
            this.insertLbl.Location = new System.Drawing.Point(263, 35);
            this.insertLbl.Name = "insertLbl";
            this.insertLbl.Size = new System.Drawing.Size(39, 13);
            this.insertLbl.TabIndex = 12;
            this.insertLbl.Text = "Insert  ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(39, 35);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(13, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = ">";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(102, 35);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(13, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = ">";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(180, 35);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(13, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = ">";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(244, 35);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(19, 13);
            this.label9.TabIndex = 16;
            this.label9.Text = ">  ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(302, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = ">  ";
            // 
            // updateLbl
            // 
            this.updateLbl.AutoSize = true;
            this.updateLbl.Location = new System.Drawing.Point(321, 35);
            this.updateLbl.Name = "updateLbl";
            this.updateLbl.Size = new System.Drawing.Size(45, 13);
            this.updateLbl.TabIndex = 17;
            this.updateLbl.Text = "Update ";
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::WiserProductFeedV2.Properties.Resources.wisepricer_headerbar;
            this.pictureBox3.Location = new System.Drawing.Point(238, -2);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(135, 71);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox3.TabIndex = 20;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = global::WiserProductFeedV2.Properties.Resources.wiser2;
            this.pictureBox1.Location = new System.Drawing.Point(371, -1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(80, 70);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::WiserProductFeedV2.Properties.Resources.wiser_enterprise;
            this.pictureBox2.Location = new System.Drawing.Point(-1, -1);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(246, 36);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 19;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // endTmr
            // 
            this.endTmr.Interval = 5000;
            this.endTmr.Tick += new System.EventHandler(this.endTmr_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(448, 64);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.updateLbl);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.insertLbl);
            this.Controls.Add(this.formatLbl);
            this.Controls.Add(this.downloadLbl);
            this.Controls.Add(this.uploadLbl);
            this.Controls.Add(this.createLbl);
            this.Controls.Add(this.createBtn);
            this.Controls.Add(this.feedLVI);
            this.Controls.Add(this.bulkBtn);
            this.Controls.Add(this.formatBtn);
            this.Controls.Add(this.downloadBtn);
            this.Controls.Add(this.uploadBtn);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.pictureBox3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Wiser Product Feed v2.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.Button uploadBtn;
        private System.Windows.Forms.Button downloadBtn;
        private System.Windows.Forms.Button formatBtn;
        private System.Windows.Forms.Button bulkBtn;
        private System.Windows.Forms.ListView feedLVI;
        private System.Windows.Forms.Button createBtn;
        private System.Windows.Forms.Label createLbl;
        private System.Windows.Forms.Label uploadLbl;
        private System.Windows.Forms.Label downloadLbl;
        private System.Windows.Forms.Label formatLbl;
        private System.Windows.Forms.Label insertLbl;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label updateLbl;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Timer endTmr;
    }
}

