namespace MasterCategoriesToMasterTable
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.totLbl = new System.Windows.Forms.Label();
            this.statusLbl = new System.Windows.Forms.Label();
            this.mcLVI = new System.Windows.Forms.ListView();
            this.verifyChk = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 41);
            this.button1.TabIndex = 1;
            this.button1.Text = "Load Master Categories";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 100);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 41);
            this.button2.TabIndex = 2;
            this.button2.Text = "Insert Master Categories";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // totLbl
            // 
            this.totLbl.AutoSize = true;
            this.totLbl.Location = new System.Drawing.Point(484, 150);
            this.totLbl.Name = "totLbl";
            this.totLbl.Size = new System.Drawing.Size(43, 13);
            this.totLbl.TabIndex = 3;
            this.totLbl.Text = "Total: 0";
            // 
            // statusLbl
            // 
            this.statusLbl.AutoSize = true;
            this.statusLbl.Location = new System.Drawing.Point(22, 155);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(40, 13);
            this.statusLbl.TabIndex = 4;
            this.statusLbl.Text = "Status:";
            // 
            // mcLVI
            // 
            this.mcLVI.FullRowSelect = true;
            this.mcLVI.GridLines = true;
            this.mcLVI.Location = new System.Drawing.Point(152, 10);
            this.mcLVI.Name = "mcLVI";
            this.mcLVI.Size = new System.Drawing.Size(769, 131);
            this.mcLVI.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.mcLVI.TabIndex = 5;
            this.mcLVI.UseCompatibleStateImageBehavior = false;
            this.mcLVI.View = System.Windows.Forms.View.Details;
            // 
            // verifyChk
            // 
            this.verifyChk.AutoSize = true;
            this.verifyChk.Checked = true;
            this.verifyChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.verifyChk.Location = new System.Drawing.Point(25, 59);
            this.verifyChk.Name = "verifyChk";
            this.verifyChk.Size = new System.Drawing.Size(105, 17);
            this.verifyChk.TabIndex = 6;
            this.verifyChk.Text = "Verify Categories";
            this.verifyChk.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(933, 177);
            this.Controls.Add(this.verifyChk);
            this.Controls.Add(this.mcLVI);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.totLbl);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Master Category Generator For MasterCategories Table";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label totLbl;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.ListView mcLVI;
        private System.Windows.Forms.CheckBox verifyChk;
    }
}

