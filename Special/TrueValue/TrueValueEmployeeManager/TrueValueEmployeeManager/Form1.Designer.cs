namespace TrueValueEmployeeManager
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
            this.employeesLst = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.databaseIDTxt = new System.Windows.Forms.TextBox();
            this.rdcIDTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.jobTitleTxt = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.firstNameTxt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lastNameTxt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.officeExtTxt = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cellPhoneTxt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.activeChk = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.saveBtn = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.refreshListLnk = new System.Windows.Forms.LinkLabel();
            this.clearFormLnk = new System.Windows.Forms.LinkLabel();
            this.phoneNumberTxt = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.previewPhotoImg = new System.Windows.Forms.PictureBox();
            this.label10 = new System.Windows.Forms.Label();
            this.imagePathLbl = new System.Windows.Forms.Label();
            this.browseLnk = new System.Windows.Forms.LinkLabel();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.previewPhotoImg)).BeginInit();
            this.SuspendLayout();
            // 
            // employeesLst
            // 
            this.employeesLst.FormattingEnabled = true;
            this.employeesLst.Location = new System.Drawing.Point(280, 37);
            this.employeesLst.Name = "employeesLst";
            this.employeesLst.Size = new System.Drawing.Size(171, 212);
            this.employeesLst.TabIndex = 0;
            this.employeesLst.SelectedIndexChanged += new System.EventHandler(this.employeesLst_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Database ID";
            // 
            // databaseIDTxt
            // 
            this.databaseIDTxt.Enabled = false;
            this.databaseIDTxt.Location = new System.Drawing.Point(101, 47);
            this.databaseIDTxt.Name = "databaseIDTxt";
            this.databaseIDTxt.Size = new System.Drawing.Size(147, 20);
            this.databaseIDTxt.TabIndex = 1;
            // 
            // rdcIDTxt
            // 
            this.rdcIDTxt.Location = new System.Drawing.Point(101, 73);
            this.rdcIDTxt.Name = "rdcIDTxt";
            this.rdcIDTxt.Size = new System.Drawing.Size(147, 20);
            this.rdcIDTxt.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "RDC ID";
            // 
            // jobTitleTxt
            // 
            this.jobTitleTxt.Location = new System.Drawing.Point(101, 99);
            this.jobTitleTxt.Name = "jobTitleTxt";
            this.jobTitleTxt.Size = new System.Drawing.Size(147, 20);
            this.jobTitleTxt.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Job Title";
            // 
            // firstNameTxt
            // 
            this.firstNameTxt.Location = new System.Drawing.Point(101, 125);
            this.firstNameTxt.Name = "firstNameTxt";
            this.firstNameTxt.Size = new System.Drawing.Size(147, 20);
            this.firstNameTxt.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "First Name";
            // 
            // lastNameTxt
            // 
            this.lastNameTxt.Location = new System.Drawing.Point(101, 151);
            this.lastNameTxt.Name = "lastNameTxt";
            this.lastNameTxt.Size = new System.Drawing.Size(147, 20);
            this.lastNameTxt.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Last Name";
            // 
            // officeExtTxt
            // 
            this.officeExtTxt.Location = new System.Drawing.Point(101, 177);
            this.officeExtTxt.Name = "officeExtTxt";
            this.officeExtTxt.Size = new System.Drawing.Size(147, 20);
            this.officeExtTxt.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 180);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(63, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Office Ext #";
            // 
            // cellPhoneTxt
            // 
            this.cellPhoneTxt.Location = new System.Drawing.Point(101, 229);
            this.cellPhoneTxt.Name = "cellPhoneTxt";
            this.cellPhoneTxt.Size = new System.Drawing.Size(147, 20);
            this.cellPhoneTxt.TabIndex = 8;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(20, 232);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Cell Phone #";
            // 
            // activeChk
            // 
            this.activeChk.AutoSize = true;
            this.activeChk.Location = new System.Drawing.Point(101, 367);
            this.activeChk.Name = "activeChk";
            this.activeChk.Size = new System.Drawing.Size(67, 17);
            this.activeChk.TabIndex = 9;
            this.activeChk.Text = "Is Active";
            this.activeChk.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.browseLnk);
            this.groupBox1.Controls.Add(this.imagePathLbl);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.previewPhotoImg);
            this.groupBox1.Controls.Add(this.phoneNumberTxt);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.clearFormLnk);
            this.groupBox1.Controls.Add(this.refreshListLnk);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.activeChk);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.employeesLst);
            this.groupBox1.Controls.Add(this.cellPhoneTxt);
            this.groupBox1.Controls.Add(this.databaseIDTxt);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.officeExtTxt);
            this.groupBox1.Controls.Add(this.rdcIDTxt);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.lastNameTxt);
            this.groupBox1.Controls.Add(this.jobTitleTxt);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.firstNameTxt);
            this.groupBox1.Location = new System.Drawing.Point(12, 20);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(486, 402);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Employee Information";
            // 
            // saveBtn
            // 
            this.saveBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.saveBtn.Location = new System.Drawing.Point(171, 440);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(172, 41);
            this.saveBtn.TabIndex = 10;
            this.saveBtn.Text = "Save / Update";
            this.saveBtn.UseVisualStyleBackColor = false;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(277, 18);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(89, 13);
            this.label8.TabIndex = 16;
            this.label8.Text = "Select Employee:";
            // 
            // refreshListLnk
            // 
            this.refreshListLnk.AutoSize = true;
            this.refreshListLnk.Location = new System.Drawing.Point(388, 18);
            this.refreshListLnk.Name = "refreshListLnk";
            this.refreshListLnk.Size = new System.Drawing.Size(63, 13);
            this.refreshListLnk.TabIndex = 11;
            this.refreshListLnk.TabStop = true;
            this.refreshListLnk.Text = "Refresh List";
            this.refreshListLnk.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.refreshListLnk_LinkClicked);
            // 
            // clearFormLnk
            // 
            this.clearFormLnk.AutoSize = true;
            this.clearFormLnk.Location = new System.Drawing.Point(98, 20);
            this.clearFormLnk.Name = "clearFormLnk";
            this.clearFormLnk.Size = new System.Drawing.Size(57, 13);
            this.clearFormLnk.TabIndex = 12;
            this.clearFormLnk.TabStop = true;
            this.clearFormLnk.Text = "Clear Form";
            this.clearFormLnk.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.clearFormLnk_LinkClicked);
            // 
            // phoneNumberTxt
            // 
            this.phoneNumberTxt.Location = new System.Drawing.Point(101, 203);
            this.phoneNumberTxt.Name = "phoneNumberTxt";
            this.phoneNumberTxt.Size = new System.Drawing.Size(147, 20);
            this.phoneNumberTxt.TabIndex = 7;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(20, 206);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(48, 13);
            this.label9.TabIndex = 19;
            this.label9.Text = "Phone #";
            // 
            // previewPhotoImg
            // 
            this.previewPhotoImg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.previewPhotoImg.Location = new System.Drawing.Point(23, 265);
            this.previewPhotoImg.Name = "previewPhotoImg";
            this.previewPhotoImg.Size = new System.Drawing.Size(85, 80);
            this.previewPhotoImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.previewPhotoImg.TabIndex = 20;
            this.previewPhotoImg.TabStop = false;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(114, 274);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(84, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "Employee Photo";
            // 
            // imagePathLbl
            // 
            this.imagePathLbl.AutoSize = true;
            this.imagePathLbl.Location = new System.Drawing.Point(114, 316);
            this.imagePathLbl.Name = "imagePathLbl";
            this.imagePathLbl.Size = new System.Drawing.Size(16, 13);
            this.imagePathLbl.TabIndex = 22;
            this.imagePathLbl.Text = "...";
            // 
            // browseLnk
            // 
            this.browseLnk.AutoSize = true;
            this.browseLnk.Location = new System.Drawing.Point(114, 294);
            this.browseLnk.Name = "browseLnk";
            this.browseLnk.Size = new System.Drawing.Size(42, 13);
            this.browseLnk.TabIndex = 23;
            this.browseLnk.TabStop = true;
            this.browseLnk.Text = "Browse";
            this.browseLnk.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.browseLnk_LinkClicked);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(510, 495);
            this.Controls.Add(this.saveBtn);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "True Value Employee Manager";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.previewPhotoImg)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox employeesLst;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox databaseIDTxt;
        private System.Windows.Forms.TextBox rdcIDTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox jobTitleTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox firstNameTxt;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox lastNameTxt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox officeExtTxt;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox cellPhoneTxt;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox activeChk;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.LinkLabel refreshListLnk;
        private System.Windows.Forms.LinkLabel clearFormLnk;
        private System.Windows.Forms.TextBox phoneNumberTxt;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.LinkLabel browseLnk;
        private System.Windows.Forms.Label imagePathLbl;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.PictureBox previewPhotoImg;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

